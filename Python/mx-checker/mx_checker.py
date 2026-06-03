"""
MX Kontrol Arayuzu
- Teknik destek ekibinin Python kurmadan tek EXE ile kullanabilmesi için tasarlandi.
- Dosya secme, cikti yolu secme, ilerleme ve log ekrani vardir.
- Desteklenen giris formatlari: .xlsx, .xlsm, .txt, .csv
"""

from __future__ import annotations

import csv
import queue
import threading
import time
from pathlib import Path
import concurrent.futures
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import dns.resolver
import dns.exception
import openpyxl


def load_domains(path: str) -> list[str]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Dosya bulunamadi: {path}")

    ext = p.suffix.lower()

    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        ws = wb.active
        domains: list[str] = []
        for row in ws.iter_rows(values_only=True):
            if not row:
                continue
            val = row[0]
            if val is not None and str(val).strip():
                domains.append(str(val).strip())
        wb.close()
        return normalize_domains(domains)

    if ext == ".csv":
        domains = []
        with open(p, newline="", encoding="utf-8", errors="ignore") as f:
            reader = csv.reader(f)
            for row in reader:
                if row and str(row[0]).strip():
                    domains.append(str(row[0]).strip())
        return normalize_domains(domains)

    if ext in (".txt", ""):
        with open(p, encoding="utf-8", errors="ignore") as f:
            return normalize_domains([line.strip() for line in f if line.strip()])

    raise ValueError(f"Desteklenmeyen dosya formati: {ext}. Yalnizca xlsx, xlsm, txt, csv desteklenir.")


def normalize_domains(items: list[str]) -> list[str]:
    cleaned: list[str] = []
    seen: set[str] = set()

    for item in items:
        d = str(item).strip().lower()
        if not d:
            continue

        # Baslik satiri gibi gorunen verileri ele
        if d in {"domain", "domainname", "alanadi", "alan adı"}:
            continue

        # Excel'den veya CSV'den gelen http/https kalintilarini temizle
        d = d.replace("http://", "").replace("https://", "")
        d = d.strip("/")

        if d and d not in seen:
            seen.add(d)
            cleaned.append(d)

    return cleaned


def check_mx(domain: str, timeout: float = 5.0, lifetime: float = 8.0) -> dict[str, str]:
    try:
        resolver = dns.resolver.Resolver()
        resolver.timeout = timeout
        resolver.lifetime = lifetime
        answers = resolver.resolve(domain, "MX")
        mx_list = sorted([f"{r.preference} {str(r.exchange).rstrip('.')}" for r in answers])
        return {"domain": domain, "status": "MX_VAR", "mx": " | ".join(mx_list)}

    except dns.resolver.NXDOMAIN:
        return {"domain": domain, "status": "MX_YOK", "mx": "Domain mevcut degil"}
    except dns.resolver.NoAnswer:
        return {"domain": domain, "status": "MX_YOK", "mx": "MX kaydi yok"}
    except dns.resolver.NoNameservers:
        return {"domain": domain, "status": "HATA", "mx": "Nameserver yok"}
    except dns.exception.Timeout:
        return {"domain": domain, "status": "HATA", "mx": "Zaman asimi"}
    except Exception as e:
        return {"domain": domain, "status": "HATA", "mx": str(e)[:180]}


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Toplu MX Kontrol")
        self.root.geometry("900x650")
        self.root.minsize(820, 560)

        self.queue: queue.Queue = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.is_running = False

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.worker_var = tk.StringVar(value="30")

        self.total = 0
        self.completed = 0
        self.mx_var_count = 0
        self.mx_yok_count = 0
        self.error_count = 0
        self.start_time = 0.0

        self._build_ui()
        self.root.after(150, self.process_queue)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)

        title = ttk.Label(outer, text="Toplu MX Kontrol Araci", font=("Segoe UI", 15, "bold"))
        title.pack(anchor="w", pady=(0, 12))

        file_frame = ttk.LabelFrame(outer, text="Dosyalar", padding=12)
        file_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(file_frame, text="Giris dosyasi").grid(row=0, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.input_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(file_frame, text="Sec", command=self.select_input).grid(row=1, column=1, sticky="ew")

        ttk.Label(file_frame, text="Cikti CSV dosyasi").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(file_frame, textvariable=self.output_var).grid(row=3, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(file_frame, text="Kaydet", command=self.select_output).grid(row=3, column=1, sticky="ew")

        file_frame.columnconfigure(0, weight=1)

        settings_frame = ttk.LabelFrame(outer, text="Ayarlar", padding=12)
        settings_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(settings_frame, text="Paralel sorgu sayisi").grid(row=0, column=0, sticky="w")
        ttk.Entry(settings_frame, textvariable=self.worker_var, width=10).grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(
            settings_frame,
            text="Onerilen: 20-50 arasi. Cok yuksek degerler bazi DNS sunucularinda hata oranini artirabilir.",
        ).grid(row=0, column=2, sticky="w", padx=(12, 0))

        action_frame = ttk.Frame(outer)
        action_frame.pack(fill="x", pady=(0, 10))

        self.start_btn = ttk.Button(action_frame, text="Baslat", command=self.start_process)
        self.start_btn.pack(side="left")

        self.open_output_btn = ttk.Button(action_frame, text="Cikti Klasorunu Ac", command=self.open_output_folder, state="disabled")
        self.open_output_btn.pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(outer, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=(0, 8))

        self.status_var = tk.StringVar(value="Hazir.")
        ttk.Label(outer, textvariable=self.status_var).pack(anchor="w", pady=(0, 10))

        stats_frame = ttk.LabelFrame(outer, text="Istatistik", padding=12)
        stats_frame.pack(fill="x", pady=(0, 10))

        self.lbl_total = ttk.Label(stats_frame, text="Toplam: 0")
        self.lbl_total.grid(row=0, column=0, sticky="w", padx=(0, 20))

        self.lbl_completed = ttk.Label(stats_frame, text="Tamamlanan: 0")
        self.lbl_completed.grid(row=0, column=1, sticky="w", padx=(0, 20))

        self.lbl_mx_var = ttk.Label(stats_frame, text="MX Var: 0")
        self.lbl_mx_var.grid(row=0, column=2, sticky="w", padx=(0, 20))

        self.lbl_mx_yok = ttk.Label(stats_frame, text="MX Yok: 0")
        self.lbl_mx_yok.grid(row=0, column=3, sticky="w", padx=(0, 20))

        self.lbl_error = ttk.Label(stats_frame, text="Hata: 0")
        self.lbl_error.grid(row=0, column=4, sticky="w")

        log_frame = ttk.LabelFrame(outer, text="Log", padding=8)
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_frame, wrap="word", height=18)
        self.log_text.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scroll.set)

        self.log("Uygulama hazir.")

    def log(self, message: str) -> None:
        now = time.strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{now}] {message}\n")
        self.log_text.see("end")

    def select_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Giris dosyasini secin",
            filetypes=[
                ("Desteklenen Dosyalar", "*.xlsx *.xlsm *.txt *.csv"),
                ("Excel", "*.xlsx *.xlsm"),
                ("Metin", "*.txt"),
                ("CSV", "*.csv"),
                ("Tum Dosyalar", "*.*"),
            ],
        )
        if path:
            self.input_var.set(path)
            if not self.output_var.get().strip():
                suggested = str(Path(path).with_name(f"{Path(path).stem}-mx-sonuclari.csv"))
                self.output_var.set(suggested)

    def select_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Cikti CSV dosyasini kaydet",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile="mx-sonuclari.csv",
        )
        if path:
            self.output_var.set(path)

    def open_output_folder(self) -> None:
        output = self.output_var.get().strip()
        if not output:
            return
        folder = str(Path(output).resolve().parent)
        try:
            import os
            os.startfile(folder)
        except Exception as e:
            messagebox.showerror("Hata", f"Klasor acilamadi:\n{e}")

    def start_process(self) -> None:
        if self.is_running:
            return

        input_file = self.input_var.get().strip()
        output_file = self.output_var.get().strip()

        if not input_file:
            messagebox.showwarning("Eksik Bilgi", "Lutfen giris dosyasini secin.")
            return

        if not output_file:
            suggested = str(Path(input_file).with_name(f"{Path(input_file).stem}-mx-sonuclari.csv"))
            self.output_var.set(suggested)
            output_file = suggested

        try:
            workers = int(self.worker_var.get().strip())
            if workers < 1 or workers > 200:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Hatali Deger", "Paralel sorgu sayisi 1 ile 200 arasinda bir tam sayi olmalidir.")
            return

        self.reset_stats()
        self.is_running = True
        self.start_btn.config(state="disabled")
        self.open_output_btn.config(state="disabled")
        self.status_var.set("Calisiyor...")
        self.log(f"Islem baslatildi. Giris: {input_file}")
        self.log(f"Cikti: {output_file}")
        self.log(f"Paralel sorgu sayisi: {workers}")

        self.worker_thread = threading.Thread(
            target=self.run_process,
            args=(input_file, output_file, workers),
            daemon=True,
        )
        self.worker_thread.start()

    def reset_stats(self) -> None:
        self.total = 0
        self.completed = 0
        self.mx_var_count = 0
        self.mx_yok_count = 0
        self.error_count = 0
        self.start_time = time.time()
        self.progress["value"] = 0
        self.progress["maximum"] = 100
        self.update_stats_labels()

    def update_stats_labels(self) -> None:
        self.lbl_total.config(text=f"Toplam: {self.total}")
        self.lbl_completed.config(text=f"Tamamlanan: {self.completed}")
        self.lbl_mx_var.config(text=f"MX Var: {self.mx_var_count}")
        self.lbl_mx_yok.config(text=f"MX Yok: {self.mx_yok_count}")
        self.lbl_error.config(text=f"Hata: {self.error_count}")

    def run_process(self, input_file: str, output_file: str, workers: int) -> None:
        try:
            domains = load_domains(input_file)
            if not domains:
                self.queue.put(("error", "Dosyada islenecek domain bulunamadi."))
                return

            self.queue.put(("init", len(domains)))
            output_path = Path(output_file)
            output_path.parent.mkdir(parents=True, exist_ok=True)

            with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=["#", "domain", "durum", "mx_kayitlari"])
                writer.writeheader()

                with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
                    futures = {executor.submit(check_mx, domain): idx for idx, domain in enumerate(domains, 1)}

                    for future in concurrent.futures.as_completed(futures):
                        num = futures[future]
                        try:
                            result = future.result()
                        except Exception as e:
                            result = {"domain": "-", "status": "HATA", "mx": str(e)[:180]}

                        writer.writerow(
                            {
                                "#": num,
                                "domain": result["domain"],
                                "durum": result["status"],
                                "mx_kayitlari": result["mx"],
                            }
                        )
                        self.queue.put(("progress", result))

            self.queue.put(("done", str(output_path)))

        except Exception as e:
            self.queue.put(("error", str(e)))

    def process_queue(self) -> None:
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]

                if kind == "init":
                    self.total = item[1]
                    self.progress["maximum"] = self.total
                    self.log(f"{self.total} domain yuklendi.")
                    self.update_stats_labels()

                elif kind == "progress":
                    result = item[1]
                    self.completed += 1
                    self.progress["value"] = self.completed

                    if result["status"] == "MX_VAR":
                        self.mx_var_count += 1
                    elif result["status"] == "MX_YOK":
                        self.mx_yok_count += 1
                    else:
                        self.error_count += 1

                    self.update_stats_labels()

                    elapsed = max(time.time() - self.start_time, 0.001)
                    speed = self.completed / elapsed
                    remaining = max(self.total - self.completed, 0)
                    eta = int(remaining / speed) if speed > 0 else 0

                    self.status_var.set(
                        f"Isleniyor... {self.completed}/{self.total} | "
                        f"MX Var: {self.mx_var_count} | MX Yok: {self.mx_yok_count} | "
                        f"Hata: {self.error_count} | Tahmini kalan: {eta} sn"
                    )

                    if self.completed <= 5 or self.completed % 100 == 0 or self.completed == self.total:
                        self.log(
                            f"{self.completed}/{self.total} tamamlandi | "
                            f"{result['domain']} -> {result['status']}"
                        )

                elif kind == "done":
                    output_path = item[1]
                    self.is_running = False
                    self.start_btn.config(state="normal")
                    self.open_output_btn.config(state="normal")
                    self.status_var.set("Tamamlandi.")
                    self.log(f"Islem tamamlandi. Cikti dosyasi: {output_path}")
                    messagebox.showinfo("Tamamlandi", f"MX kontrol islemi tamamlandi.\n\nCikti dosyasi:\n{output_path}")

                elif kind == "error":
                    error_message = item[1]
                    self.is_running = False
                    self.start_btn.config(state="normal")
                    self.status_var.set("Hata olustu.")
                    self.log(f"HATA: {error_message}")
                    messagebox.showerror("Hata", error_message)

        except queue.Empty:
            pass

        self.root.after(150, self.process_queue)


def main() -> None:
    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

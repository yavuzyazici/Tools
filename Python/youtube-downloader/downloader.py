# downloader.py
# YouTube -> MP3 (const URL) | yt-dlp + FFmpeg
# KULLANIM: py downloader.py
# UYARI: Yalnızca size ait/izinli veya açık lisanslı içerikler için kullanın.

from pathlib import Path
import sys
from yt_dlp import YoutubeDL

# === SABİTLER ===
URL = "https://www.youtube.com/watch?v=rhpZNno8dbM"   # <- kendi linkinizi girin
OUTPUT_DIR = "mp3"                                  # çıkış klasörü
BITRATE = "320"                                     # 128/160/192/256/320
DOWNLOAD_PLAYLIST = False                            # playlist ise tamamını indir
FFMPEG_PATH = r"E:\ffmpeg\bin"                      # PATH'e ekliyseniz "" bırakın

def main():
    if not URL or URL.startswith("https://www.youtube.com/watch?v=XXXX"):
        print("Lütfen URL sabitini geçerli bir YouTube linkiyle değiştirin.", file=sys.stderr)
        sys.exit(1)

    outdir = Path(OUTPUT_DIR).expanduser().resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    ydl_opts = {
        # Yalnızca ses akışlarını tercih et; m4a/webm sesleri öne al
        "format": "bestaudio[ext=m4a]/bestaudio[ext=webm]/bestaudio/best",
        "outtmpl": str(outdir / "%(title)s.%(ext)s"),
        "noplaylist": not DOWNLOAD_PLAYLIST,
        "restrictfilenames": True,
        "trim_file_name": 200,
        "prefer_ffmpeg": True,
        "overwrites": False,
        "keepvideo": False,  # MP3'e çevrildikten sonra kaynak dosyayı sil

        # SABR/İmzalama sorunlarında daha uyumlu istemciler:
        "extractor_args": {"youtube": {"player_client": ["android", "web"]}},

        # FFmpeg yolu (PATH'e ekli değilse)
        **({"ffmpeg_location": FFMPEG_PATH} if FFMPEG_PATH else {}),

        # MP3 dönüştürme
        "postprocessors": [
            {
                "key": "FFmpegExtractAudio",
                "preferredcodec": "mp3",
                "preferredquality": BITRATE,
            },
            {"key": "FFmpegMetadata"},
        ],
    }

    try:
        with YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(URL, download=True)
            if not info:
                print("İndirme başarısız oldu.", file=sys.stderr)
                sys.exit(1)
        print("MP3 indirildi ve dönüştürüldü →", outdir)
    except FileNotFoundError:
        print("FFmpeg/ffprobe bulunamadı. FFmpeg’i kurup PATH’e ekleyin "
              "veya FFMPEG_PATH sabitini doğru klasöre ayarlayın.", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Hata: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()

#!/usr/bin/env python3
import argparse
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog

import webview


# --- Configuration (match main.ps1 paths) ---

VIDEO_SOURCE_DIR = Path(r"C:\Users\LEGION\Videos")
TITLE_FILE_PATH = Path(r"C:\Users\LEGION\Desktop\title.txt")
WORK_DIR = Path(r"C:\Users\LEGION\Desktop\إخراج")
DEST_ORIG_VIDEO = Path(r"D:\01 - الفيديو")
DEST_COMP_VIDEO = Path(r"D:\02 - مضغوط")
DEST_AUDIO = Path(r"D:\03 - صوت")


def ensure_directories() -> None:
    for p in (WORK_DIR, DEST_ORIG_VIDEO, DEST_COMP_VIDEO, DEST_AUDIO):
        p.mkdir(parents=True, exist_ok=True)


# --- Helper functions (Python equivalents) ---

def escape_html_for_dialog(text: str) -> str:
    """Escape text for safe use in HTML (e.g. value="" attribute)."""
    if not text:
        return ""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def get_safe_filename(text: str) -> str:
    if not text:
        return ""
    # Basic Windows-invalid chars; good enough mirror of .NET's GetInvalidFileNameChars for our usage
    invalid_chars = '<>:"/\\|?*'
    sanitized = "".join((" " if ch in invalid_chars else ch) for ch in text)
    # Collapse whitespace
    sanitized = re.sub(r"\s+", " ", sanitized).strip()
    return sanitized


def normalize_leading_date_in_title(text: str) -> str:
    if not text:
        return ""
    m = re.match(r"^\s*(\d{4})[-/\s\\]+(\d{1,2})[-/\s\\]+(\d{1,2})(\s.*|$)", text)
    if not m:
        return text
    y, mth, d, rest = m.groups()
    mth = mth.zfill(2)
    d = d.zfill(2)
    return f"{y}{mth}{d}{rest}"


def get_video_fps(file_path: Path) -> float | None:
    """Get average FPS using ffprobe (mirrors Get-VideoFps)."""
    try:
        proc = subprocess.run(
            [
                "ffprobe",
                "-v",
                "error",
                "-select_streams",
                "v:0",
                "-show_entries",
                "stream=avg_frame_rate",
                "-of",
                "default=noprint_wrappers=1:nokey=1",
                str(file_path),
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            text=True,
            check=False,
        )
    except FileNotFoundError:
        return None

    rate_str = (proc.stdout.splitlines() or [""])[0].strip()
    if not rate_str:
        return None

    m = re.match(r"^(?P<num>\d+)\s*/\s*(?P<den>\d+)$", rate_str)
    if m:
        num = float(m.group("num"))
        den = float(m.group("den"))
        if den <= 0:
            return None
        return num / den

    try:
        return float(rate_str)
    except ValueError:
        return None


def is_file_locked(file_path: Path) -> bool:
    """Return True if file appears locked (similar intent to Test-IsFileLocked)."""
    try:
        # Try opening for read; if another process has an exclusive lock we should get an error.
        with open(file_path, "rb"):
            return False
    except OSError:
        return True


def get_audio_encode_args(input_path: Path, output_path: Path, meta_title: str, album_artist: str) -> list[str]:
    # Mirrors Get-AudioEncodeArgs in main.ps1
    return [
        "-y",
        "-i",
        str(input_path),
        "-map",
        "0:a:0",
        "-vn",
        "-c:a",
        "libopus",
        "-application",
        "voip",
        "-b:a",
        "18k",
        "-map_metadata",
        "-1",
        "-metadata",
        f"title={meta_title}",
        "-metadata",
        f"album_artist={album_artist}",
        "-metadata:s:a:0",
        f"title={meta_title}",
        str(output_path),
    ]


def get_h265480_encode_args(input_path: Path, output_path: Path, meta_title: str) -> list[str]:
    # Mirrors Get-H265480EncodeArgs in main.ps1
    return [
        "-y",
        "-i",
        str(input_path),
        "-vf",
        "scale=-2:480",
        "-map_metadata",
        "-1",
        "-c:v",
        "libx265",
        "-preset",
        "veryfast",
        "-crf",
        "24",
        "-r",
        "25",
        "-c:a",
        "copy",
        "-c:s",
        "copy",
        "-map_chapters",
        "0",
        "-metadata:s:a:0",
        f"title={meta_title}",
        str(output_path),
    ]


# --- GUI dialogs (tkinter) ---

def _init_tk_root() -> tk.Tk:
    root = tk.Tk()
    # We mostly use it for dialogs; no default root window UI.
    root.withdraw()
    return root


def show_file_selector(initial_directory: Path | None = None) -> Path | None:
    root = _init_tk_root()
    try:
        initialdir = str(initial_directory) if initial_directory and initial_directory.exists() else os.path.join(
            os.environ.get("USERPROFILE", ""), "Videos"
        )
        file_path = filedialog.askopenfilename(
            title="Select a Video File",
            initialdir=initialdir,
            filetypes=(
                ("Video Files", "*.mp4 *.mkv *.mov *.avi *.flv *.obs"),
                ("All Files", "*.*"),
            ),
        )
        if not file_path:
            return None
        return Path(file_path)
    finally:
        root.destroy()


# --- HTML title dialog (pywebview) ---

_DIALOG_HTML = """<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
<meta charset="UTF-8">
<style>
  body {
    font-family: 'Segoe UI', Tahoma, sans-serif;
    font-size: 14px;
    padding: 24px;
    margin: 0;
    background: #1e1e1e;
    color: #e0e0e0;
  }
  .row { margin-bottom: 16px; }
  label { display: inline-block; width: 90px; font-weight: bold; color: #e0e0e0; }
  input[type="text"] {
    width: 280px;
    padding: 8px 10px;
    border: 1px solid #555;
    border-radius: 4px;
    font-size: 14px;
    background: #3c3c3c;
    color: #e0e0e0;
  }
  input[type="text"]::placeholder { color: #888; }
  input[type="checkbox"] {
    margin-left: 8px;
    vertical-align: middle;
    accent-color: #0e639c;
  }
  .buttons { margin-top: 28px; padding-top: 16px; }
  .buttons button {
    padding: 12px 28px;
    margin-left: 14px;
    cursor: pointer;
    font-size: 14px;
    border-radius: 6px;
    border: 1px solid #555;
  }
  .buttons button:first-of-type {
    margin-left: 0;
    background: #0e639c;
    color: #fff;
    border-color: #0e639c;
  }
  .buttons button:first-of-type:hover { background: #1177bb; }
  .buttons button:last-of-type {
    background: #3c3c3c;
    color: #e0e0e0;
  }
  .buttons button:last-of-type:hover { background: #505050; }
</style>
</head>
<body>
  <div class="row"><label>العنوان:</label><input type="text" id="titleInput" value="{{TITLE}}" dir="rtl"></div>
  <div class="row"><label>الفنان:</label><input type="text" id="artistInput" value="{{ARTIST}}" dir="rtl"></div>
  <div class="row"><label>تخطي الصوت:</label><input type="checkbox" id="skipAudio"{{SKIP_AUDIO_CHECKED}}></div>
  <div class="row"><label>تخطي الفيديو:</label><input type="checkbox" id="skipVideo"{{SKIP_VIDEO_CHECKED}}></div>
  <div class="buttons">
    <button type="button" id="btnOk">موافق</button>
    <button type="button" id="btnCancel">إلغاء</button>
  </div>
<script>
window.addEventListener('pywebviewready', function() {
  function submit() {
    var title = document.getElementById('titleInput').value || '';
    var artist = document.getElementById('artistInput').value || '';
    var skipAudio = document.getElementById('skipAudio').checked;
    var skipVideo = document.getElementById('skipVideo').checked;
    pywebview.api.submit(title, artist, skipAudio, skipVideo);
  }
  function cancel() { pywebview.api.cancel(); }
  document.getElementById('btnOk').onclick = submit;
  document.getElementById('btnCancel').onclick = cancel;
  var titleEl = document.getElementById('titleInput');
  var artistEl = document.getElementById('artistInput');
  var firstTime = { title: true, artist: true };
  titleEl.onfocus = function() { if (firstTime.title) { firstTime.title = false; titleEl.select(); } };
  titleEl.onmouseup = function() { if (firstTime.title) titleEl.select(); };
  artistEl.onfocus = function() { if (firstTime.artist) { firstTime.artist = false; artistEl.select(); } };
  artistEl.onmouseup = function() { if (firstTime.artist) artistEl.select(); };
});
</script>
</body>
</html>
"""


class _DialogAPI:
    """Exposed to JS as pywebview.api. submit/cancel store result and close the window."""

    def __init__(self) -> None:
        self._window: webview.Window | None = None
        self._result: dict | None = None

    def set_window(self, window: webview.Window) -> None:
        self._window = window

    def submit(self, title: str, artist: str, skip_audio: bool, skip_video: bool) -> None:
        self._result = {
            "ok": True,
            "title": (title or "").strip(),
            "artist": (artist or "").strip(),
            "skip_audio": bool(skip_audio),
            "skip_video": bool(skip_video),
        }
        if self._window:
            self._window.destroy()

    def cancel(self) -> None:
        self._result = {"ok": False}
        if self._window:
            self._window.destroy()

    def get_result(self) -> dict | None:
        return self._result


def show_title_confirmation_dialog(
    title_text: str,
    artist_text: str,
    initial_skip_audio: bool = False,
    initial_skip_video: bool = False,
) -> dict | None:
    title_escaped = escape_html_for_dialog(title_text or "")
    artist_escaped = escape_html_for_dialog(artist_text or "")
    skip_audio_checked = " checked" if initial_skip_audio else ""
    skip_video_checked = " checked" if initial_skip_video else ""

    html = (
        _DIALOG_HTML.replace("{{TITLE}}", title_escaped)
        .replace("{{ARTIST}}", artist_escaped)
        .replace("{{SKIP_AUDIO_CHECKED}}", skip_audio_checked)
        .replace("{{SKIP_VIDEO_CHECKED}}", skip_video_checked)
    )

    api = _DialogAPI()
    window = webview.create_window(
        "Confirm Title",
        html=html,
        js_api=api,
        width=520,
        height=320,
    )
    api.set_window(window)
    webview.start()
    return api.get_result()


# --- Main workflow ---

def run(debug_program: bool, cli_skip_audio: bool) -> None:
    # Step 0: ensure UTF-8 for Arabic output
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass

    ensure_directories()

    # --- Step 1: Wait for Title File (Must have 2 lines) ---
    print("Checking Title file for 2 lines...")
    do_skip_audio = False
    do_skip_video = False

    while True:
        if TITLE_FILE_PATH.exists():
            with TITLE_FILE_PATH.open("r", encoding="utf-8") as f:
                lines = [line.rstrip("\n") for line in f]
            valid_lines = [ln for ln in lines if ln.strip()]

            if len(valid_lines) >= 2:
                clean_title = valid_lines[0].strip()
                album_artist = valid_lines[1].strip()
                print("Title file OK. Showing confirmation...")
                dlg = show_title_confirmation_dialog(
                    title_text=clean_title,
                    artist_text=album_artist,
                    initial_skip_audio=cli_skip_audio,
                    initial_skip_video=False,
                )
                if not dlg or not dlg.get("ok"):
                    sys.exit(0)

                clean_title = dlg.get("title", "").strip()
                album_artist = dlg.get("artist", "").strip()
                do_skip_audio = bool(dlg.get("skip_audio")) or cli_skip_audio
                do_skip_video = bool(dlg.get("skip_video"))

                clean_title = normalize_leading_date_in_title(clean_title)

                with TITLE_FILE_PATH.open("w", encoding="utf-8") as f:
                    f.write(clean_title + "\n")
                    f.write(album_artist + "\n")
                break

        print(f"Waiting for 2 lines in {TITLE_FILE_PATH}...")
        time.sleep(0.5)

    # --- Step 2: Select Video ---
    videos_folder = Path(os.environ.get("USERPROFILE", "")) / "Videos"
    if not videos_folder.exists():
        videos_folder = VIDEO_SOURCE_DIR

    selected_path = show_file_selector(videos_folder)
    if selected_path is None:
        sys.exit(0)
    latest_video = selected_path
    print(f"Selected: {latest_video.name}")

    # --- Step 3: Wait for OBS to Stop Recording ---
    print(f"Monitoring: {latest_video.name}")
    print("OBS is currently recording... Waiting.")
    while is_file_locked(latest_video):
        time.sleep(1.0)

    print("Recording finished! Starting processing...")

    # --- Step 4: Define Names and Paths ---
    safe_title = get_safe_filename(clean_title)
    title_starts_with_date = bool(re.match(r"^\s*(\d{8}|\d{4}[-/\s\\]?\d{2}[-/\s\\]?\d{2})", clean_title))
    if title_starts_with_date:
        base_name = safe_title
    else:
        base_name = f"{datetime.now():%Y%m%d} {safe_title}"

    audio_out_path = WORK_DIR / f"{base_name}.opus"
    video_out_path = WORK_DIR / f"{base_name}.mp4"

    folder_date_name = datetime.now().strftime("%Y-%m-%d")
    daily_folder = DEST_ORIG_VIDEO / folder_date_name
    final_orig_path = daily_folder / f"{base_name}{latest_video.suffix}"
    work_dir_orig_path = WORK_DIR / f"{base_name}{latest_video.suffix}"

    # --- Step 4b: Copy Original First ---
    if not debug_program:
        daily_folder.mkdir(parents=True, exist_ok=True)
        if not work_dir_orig_path.exists():
            print(f"  > Copying Original to work folder ({WORK_DIR})...")
            shutil.copy2(latest_video, work_dir_orig_path)
        else:
            print("  > Original already in work folder. Skipping.")

        if not final_orig_path.exists():
            print(f"  > Copying Original to {daily_folder}...")
            shutil.copy2(latest_video, final_orig_path)
        else:
            print("  > Original already in final folder. Skipping.")

    # --- Step 5: Process Files (NO OVERWRITE) ---

    # 5a. Extract Audio
    if do_skip_audio:
        print("  > Audio step skipped (user choice).")
    elif not audio_out_path.exists():
        print("  > Extracting Audio...")
        meta_title = "الشيخ محمد فواز النمر"
        audio_args = get_audio_encode_args(latest_video, audio_out_path, meta_title, album_artist)
        print("  > ffmpeg audio args:")
        print("    ffmpeg " + " ".join(audio_args))
        subprocess.run(["ffmpeg", *audio_args], check=False)
    else:
        print("  > Audio already exists. Skipping.")

    # 5b. Re-encode Video
    if do_skip_video:
        print("  > Video re-encode skipped (user choice).")
    elif not video_out_path.exists():
        print("  > Re-encoding Video (H.265 libx265 480p)...")
        meta_title = "الشيخ محمد فواز النمر"
        h265_args = get_h265480_encode_args(latest_video, video_out_path, meta_title)
        print("  > ffmpeg H.265 480p args:")
        print("    ffmpeg " + " ".join(h265_args))
        subprocess.run(["ffmpeg", *h265_args], check=False)
    else:
        print("  > Compressed video already exists. Skipping.")

    # --- Step 6: Copy to Destinations (NO OVERWRITE) ---
    if not debug_program:
        # 6a. Original already copied.

        # 6b. Copy Audio
        if do_skip_audio:
            print("  > Audio copy skipped (user choice).")
        else:
            final_audio_path = DEST_AUDIO / f"{base_name}.opus"
            if not final_audio_path.exists():
                print("  > Copying Audio...")
                shutil.copy2(audio_out_path, final_audio_path)
            else:
                print("  > Audio file already in destination. Skipping.")

        # 6c. Copy Compressed Video
        if do_skip_video:
            print("  > Compressed video copy skipped (user choice).")
        else:
            final_comp_path = DEST_COMP_VIDEO / f"{base_name}.mp4"
            if not final_comp_path.exists():
                print("  > Copying Compressed Video...")
                shutil.copy2(video_out_path, final_comp_path)
            else:
                print("  > Compressed video already in destination. Skipping.")
    else:
        print("Copy step skipped (debug mode).")

    print("All tasks completed successfully. Exiting.")
    time.sleep(5)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Video processing script (Python port of main.ps1).")
    parser.add_argument(
        "--debug-program",
        action="store_true",
        help="Skip final copy steps (equivalent to -debugProgram).",
    )
    parser.add_argument(
        "--skip-audio",
        action="store_true",
        help="Skip audio extraction/copy by default (can still override in dialog).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    run(debug_program=args.debug_program, cli_skip_audio=args.skip_audio)


if __name__ == "__main__":
    main()


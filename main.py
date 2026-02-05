from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import messagebox


def run_powershell(
    title: str,
    artist: str,
    skip_audio: bool,
    skip_video: bool,
    debug_program: bool,
) -> int:
    """
    Invoke main.ps1 via pwsh with the collected parameters.
    Returns the PowerShell process return code.
    """
    repo_root = Path(__file__).resolve().parent
    ps1_path = repo_root / "main.ps1"

    if not ps1_path.exists():
        messagebox.showerror("Error", f"Could not find PowerShell script:\n{ps1_path}")
        return 1

    cmd: list[str] = [
        "pwsh",
        "-NoLogo",
        "-NonInteractive",
        "-File",
        str(ps1_path),
        "-Title",
        title,
        "-Artist",
        artist,
    ]

    if skip_audio:
        cmd.append("-skipAudio")
    if skip_video:
        cmd.append("-skipVideo")
    if debug_program:
        cmd.append("-debugProgram")

    # Let PowerShell write directly to the same console.
    try:
        completed = subprocess.run(cmd, cwd=str(repo_root))
    except FileNotFoundError:
        messagebox.showerror(
            "Error",
            "Could not find 'pwsh' on PATH.\n"
            "Make sure PowerShell 7+ is installed and 'pwsh' is available.",
        )
        return 1

    return completed.returncode


def main(
    initial_title: str | None = None,
    initial_artist: str | None = None,
    initial_skip_audio: bool = False,
    initial_skip_video: bool = False,
    initial_debug: bool = False,
) -> None:
    root = tk.Tk()
    root.title("Video Processor")

    # Simple, compact layout
    root.resizable(False, False)

    # String / boolean variables (pre-filled from any CLI args)
    title_var = tk.StringVar(value=initial_title or "")
    artist_var = tk.StringVar(value=initial_artist or "")
    skip_audio_var = tk.BooleanVar(value=bool(initial_skip_audio))
    skip_video_var = tk.BooleanVar(value=bool(initial_skip_video))
    debug_var = tk.BooleanVar(value=bool(initial_debug))

    # Title
    tk.Label(root, text="Title:").grid(row=0, column=0, sticky="e", padx=8, pady=6)
    tk.Entry(root, textvariable=title_var, width=40).grid(
        row=0, column=1, columnspan=2, sticky="w", padx=8, pady=6
    )

    # Artist
    tk.Label(root, text="Artist:").grid(row=1, column=0, sticky="e", padx=8, pady=6)
    tk.Entry(root, textvariable=artist_var, width=40).grid(
        row=1, column=1, columnspan=2, sticky="w", padx=8, pady=6
    )

    # Checkboxes
    tk.Checkbutton(root, text="Skip audio", variable=skip_audio_var).grid(
        row=2, column=0, columnspan=3, sticky="w", padx=8
    )
    tk.Checkbutton(root, text="Skip video", variable=skip_video_var).grid(
        row=3, column=0, columnspan=3, sticky="w", padx=8
    )
    tk.Checkbutton(root, text="Debug mode", variable=debug_var).grid(
        row=4, column=0, columnspan=3, sticky="w", padx=8
    )

    def on_ok() -> None:
        title = title_var.get().strip()
        artist = artist_var.get().strip()

        if not title or not artist:
            messagebox.showwarning(
                "Missing data",
                "Both Title and Artist must be filled in.",
            )
            return

        root.withdraw()  # Hide window while PowerShell runs
        rc = run_powershell(
            title=title,
            artist=artist,
            skip_audio=skip_audio_var.get(),
            skip_video=skip_video_var.get(),
            debug_program=debug_var.get(),
        )
        if rc == 0:
            messagebox.showinfo("Done", "Processing completed successfully.")
        else:
            messagebox.showerror(
                "Error",
                f"PowerShell script exited with code {rc}. See console output for details.",
            )
        root.destroy()

    def on_cancel() -> None:
        root.destroy()

    # Buttons
    button_frame = tk.Frame(root)
    button_frame.grid(row=5, column=0, columnspan=3, pady=10)

    tk.Button(button_frame, text="OK", width=10, command=on_ok).pack(
        side="left", padx=5
    )
    tk.Button(button_frame, text="Cancel", width=10, command=on_cancel).pack(
        side="left", padx=5
    )

    # Make Enter = OK, Escape = Cancel
    root.bind("<Return>", lambda _event: on_ok())
    root.bind("<Escape>", lambda _event: on_cancel())

    root.mainloop()


if __name__ == "__main__":
    # Always parse CLI args, but use them only to pre-fill the GUI.
    parser = argparse.ArgumentParser(
        description="Python launcher + GUI for main.ps1 (Video Processor).",
        add_help=True,
    )
    parser.add_argument(
        "-Title",
        dest="Title",
        help="Initial title to show in the GUI (also passed to main.ps1 if unchanged).",
    )
    parser.add_argument(
        "-Artist",
        dest="Artist",
        help="Initial artist to show in the GUI (also passed to main.ps1 if unchanged).",
    )
    parser.add_argument(
        "-skipAudio",
        action="store_true",
        help="Pre-check 'Skip audio' in the GUI (maps to -skipAudio).",
    )
    parser.add_argument(
        "-skipVideo",
        action="store_true",
        help="Pre-check 'Skip video' in the GUI (maps to -skipVideo).",
    )
    parser.add_argument(
        "-debugProgram",
        action="store_true",
        help="Pre-check 'Debug mode' in the GUI (maps to -debugProgram).",
    )

    args = parser.parse_args()

    # Optionally pre-populate from a Desktop title.txt (first two non-empty lines).
    # Looks for ~/Desktop/title.txt (cross-platform home + Desktop).
    repo_root = Path(__file__).resolve().parent
    desktop = Path.home() / "Desktop"
    title_file = desktop / "title.txt"
    file_title: str | None = None
    file_artist: str | None = None
    if title_file.exists():
        try:
            lines = [line.strip() for line in title_file.read_text(encoding="utf-8").splitlines()]
            valid = [ln for ln in lines if ln]
            if len(valid) >= 2:
                file_title = valid[0]
                file_artist = valid[1]
        except Exception:
            # If anything goes wrong reading/parsing, just ignore and fall back to args only.
            pass

    # CLI args win over title.txt; title.txt is only used when args are missing.
    initial_title = args.Title or file_title
    initial_artist = args.Artist or file_artist

    try:
        main(
            initial_title=initial_title,
            initial_artist=initial_artist,
            initial_skip_audio=args.skipAudio,
            initial_skip_video=args.skipVideo,
            initial_debug=args.debugProgram,
        )
    except Exception as exc:  # pragma: no cover - last-resort error dialog
        messagebox.showerror("Unexpected error", str(exc))
        sys.exit(1)


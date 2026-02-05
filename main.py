from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QLineEdit,
    QCheckBox,
    QPushButton,
    QHBoxLayout,
    QVBoxLayout,
    QMessageBox,
)


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
        print(f"Error: Could not find PowerShell script: {ps1_path}", file=sys.stderr)
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
        print(
            "Error: Could not find 'pwsh' on PATH. "
            "Make sure PowerShell 7+ is installed and 'pwsh' is available.",
            file=sys.stderr,
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
    """
    Launch a Qt-based dialog to collect title/artist and skip flags,
    then invoke the PowerShell script.
    """

    app = QApplication.instance() or QApplication(sys.argv)

    class MainWindow(QWidget):
        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle("Video Processor")

            # Widgets
            self.title_edit = QLineEdit(self)
            self.artist_edit = QLineEdit(self)

            if initial_title:
                self.title_edit.setText(initial_title)
            if initial_artist:
                self.artist_edit.setText(initial_artist)

            self.skip_audio_cb = QCheckBox("Skip audio", self)
            self.skip_audio_cb.setChecked(bool(initial_skip_audio))

            self.skip_video_cb = QCheckBox("Skip video", self)
            self.skip_video_cb.setChecked(bool(initial_skip_video))

            self.debug_cb: QCheckBox | None = None
            if initial_debug:
                self.debug_cb = QCheckBox("Debug mode", self)
                self.debug_cb.setChecked(True)

            ok_btn = QPushButton("OK", self)
            cancel_btn = QPushButton("Cancel", self)

            ok_btn.clicked.connect(self.on_ok)  # type: ignore[arg-type]
            cancel_btn.clicked.connect(self.close)  # type: ignore[arg-type]

            # Layout
            form_layout = QVBoxLayout()

            title_row = QHBoxLayout()
            title_row.addWidget(QLabel("Title:", self))
            title_row.addWidget(self.title_edit)
            form_layout.addLayout(title_row)

            artist_row = QHBoxLayout()
            artist_row.addWidget(QLabel("Artist:", self))
            artist_row.addWidget(self.artist_edit)
            form_layout.addLayout(artist_row)

            form_layout.addWidget(self.skip_audio_cb)
            form_layout.addWidget(self.skip_video_cb)
            if self.debug_cb is not None:
                form_layout.addWidget(self.debug_cb)

            buttons_row = QHBoxLayout()
            buttons_row.addStretch(1)
            buttons_row.addWidget(ok_btn)
            buttons_row.addWidget(cancel_btn)

            main_layout = QVBoxLayout(self)
            main_layout.addLayout(form_layout)
            main_layout.addLayout(buttons_row)

            self.setLayout(main_layout)

        def on_ok(self) -> None:
            title = self.title_edit.text().strip()
            artist = self.artist_edit.text().strip()

            if not title or not artist:
                QMessageBox.warning(
                    self,
                    "Missing data",
                    "Both Title and Artist must be filled in.",
                    QMessageBox.Ok,
                )
                return

            skip_audio = self.skip_audio_cb.isChecked()
            skip_video = self.skip_video_cb.isChecked()
            debug_program = bool(self.debug_cb and self.debug_cb.isChecked())

            rc = run_powershell(
                title=title,
                artist=artist,
                skip_audio=skip_audio,
                skip_video=skip_video,
                debug_program=debug_program,
            )

            if rc == 0:
                QMessageBox.information(
                    self,
                    "Done",
                    "Processing completed successfully.",
                    QMessageBox.Ok,
                )
            else:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"PowerShell script exited with code {rc}. See console output for details.",
                    QMessageBox.Ok,
                )

            self.close()

        def keyPressEvent(self, event) -> None:  # type: ignore[override]
            if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                self.on_ok()
                return
            if event.key() == Qt.Key_Escape:
                self.close()
                return
            super().keyPressEvent(event)

    window = MainWindow()
    window.resize(480, 160)
    window.show()
    app.exec()


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


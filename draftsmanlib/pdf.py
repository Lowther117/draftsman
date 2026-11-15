"""A PDF copy of a finished document, made by whichever of Microsoft Word or
LibreOffice is on the machine. Draftsman cannot lay a Word file out itself -
only a word processor can do that faithfully - so without either there is no
PDF, and the .docx is still there."""
from __future__ import annotations

import os
import shutil
import subprocess
import sys

WORD_PS = r"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($env:DRAFTSMAN_IN, $false, $true)
  $doc.ExportAsFixedFormat($env:DRAFTSMAN_OUT, 17)
  $doc.Close($false)
} finally {
  $word.Quit()
}
"""


class NoConverter(RuntimeError):
    pass


def _soffice():
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    for candidate in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                      "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.isfile(candidate):
            return candidate
    return None


def _quiet():
    return {"creationflags": 0x08000000} if sys.platform == "win32" else {}     # CREATE_NO_WINDOW


def _by_word_windows(src, dst):
    env = dict(os.environ, DRAFTSMAN_IN=src, DRAFTSMAN_OUT=dst)
    subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", WORD_PS],
                   env=env, check=True, capture_output=True, timeout=180, **_quiet())


def _by_word_mac(src, dst):
    script = ['tell application "Microsoft Word"', f'open POSIX file "{src}"',
              f'save as active document file name "{dst}" file format format PDF',
              "close active document saving no", "end tell"]
    args = ["osascript"]
    for line in script:
        args += ["-e", line]
    subprocess.run(args, check=True, capture_output=True, timeout=180)


def _by_libreoffice(exe, src, dst):
    subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(dst), src],
                   check=True, capture_output=True, timeout=180, **_quiet())


def to_pdf(docx_path: str) -> str:
    """Write <name>.pdf beside the .docx and return its path. Raises
    NoConverter when neither Word nor LibreOffice could do it."""
    src = os.path.abspath(docx_path)
    dst = os.path.splitext(src)[0] + ".pdf"
    routes = []
    if sys.platform == "win32":
        routes.append(lambda: _by_word_windows(src, dst))
    exe = _soffice()
    if exe:
        routes.append(lambda: _by_libreoffice(exe, src, dst))
    if sys.platform == "darwin" and os.path.isdir("/Applications/Microsoft Word.app") and '"' not in src:
        routes.append(lambda: _by_word_mac(src, dst))
    for route in routes:
        try:
            if os.path.exists(dst):
                os.remove(dst)
            route()
            if os.path.isfile(dst) and os.path.getsize(dst) > 0:
                return dst
        except (OSError, subprocess.SubprocessError):
            continue
    raise NoConverter("A PDF copy needs Microsoft Word or LibreOffice on this computer, and neither could make one.")

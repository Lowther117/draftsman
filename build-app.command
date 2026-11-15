#!/bin/bash
# ---------------------------------------------------------------------------
#  Draftsman - macOS build. Double-click this file.
#  Installs Homebrew and Python if they are missing, installs every
#  dependency, and produces "dist/Draftsman.app".
# ---------------------------------------------------------------------------
cd "$(dirname "$0")" || exit 1
HERE="$(pwd)"
LOG="$HERE/build-mac-log.txt"
: > "$LOG"
exec > >(tee -a "$LOG") 2>&1

# A double-clicked .command starts with a bare PATH that has no Homebrew in it.
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
export HOMEBREW_NO_AUTO_UPDATE=1
export HOMEBREW_NO_INSTALL_CLEANUP=1

APP="$HERE/dist/Draftsman.app"
REPORT="$HERE/dist/draftsman-selftest.txt"

echo "==============================================="
echo "  Building Draftsman for macOS"
echo "  Full log: build-mac-log.txt"
echo "==============================================="
echo

die() {
  echo
  echo "==============================================="
  echo "  BUILD FAILED: $*"
  echo "==============================================="
  echo
  echo "Press return to close."
  read -r
  exit 1
}

ensure_brew() {
  if command -v brew >/dev/null 2>&1; then return 0; fi
  echo "Homebrew is not installed. Installing it now - it will ask for your password once."
  /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)" < /dev/tty \
    || die "Homebrew would not install."
  export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
  command -v brew >/dev/null 2>&1 || die "Homebrew installed but brew is still not on the PATH."
}

brew_install() {
  ensure_brew
  echo "  brew install $*"
  brew install "$@" < /dev/null || die "brew install $* failed."
}
# Apple's /usr/bin/python3 links against system Tk 8.5, which PyInstaller
# cannot bundle: the app builds and then never opens a window. Only accept an
# interpreter whose tkinter reports 8.6 or newer. Resolve candidates with
# readlink rather than running python3, because on a clean Mac running
# python3 pops up the Xcode command line tools installer.
#
# pick_python            prints the newest suitable interpreter
# pick_python ":a:b:"    the same, leaving out the interpreters listed between
#                        the colons
pick_python() {
  local skip="${1:-}" c real
  if [ -z "$skip" ] && [ -n "$DRAFTSMAN_BUILD_PYTHON" ] && [ -x "$DRAFTSMAN_BUILD_PYTHON" ]; then
    if "$DRAFTSMAN_BUILD_PYTHON" -c 'import tkinter,sys;sys.exit(0 if tkinter.TkVersion>=8.6 else 1)' 2>/dev/null; then
      echo "$DRAFTSMAN_BUILD_PYTHON"; return 0
    fi
  fi
  for c in /opt/homebrew/bin/python3.14 /opt/homebrew/bin/python3.13 /opt/homebrew/bin/python3.12 \
           /opt/homebrew/bin/python3.11 /opt/homebrew/bin/python3.10 /opt/homebrew/bin/python3.9 \
           /usr/local/bin/python3.14 /usr/local/bin/python3.13 /usr/local/bin/python3.12 \
           /usr/local/bin/python3.11 /usr/local/bin/python3.10 /usr/local/bin/python3.9 \
           /Library/Frameworks/Python.framework/Versions/3.14/bin/python3 \
           /Library/Frameworks/Python.framework/Versions/3.13/bin/python3 \
           /Library/Frameworks/Python.framework/Versions/3.12/bin/python3 \
           /Library/Frameworks/Python.framework/Versions/3.11/bin/python3 \
           /Library/Frameworks/Python.framework/Versions/3.10/bin/python3 \
           /Library/Frameworks/Python.framework/Versions/3.9/bin/python3; do
    [ -e "$c" ] || continue
    real="$(readlink -f "$c" 2>/dev/null || echo "$c")"
    [ -x "$real" ] || continue
    case "$real" in /usr/bin/python3) continue ;; esac
    case "$skip" in *":$real:"*) continue ;; esac
    if "$real" -c 'import tkinter,sys;sys.exit(0 if tkinter.TkVersion>=8.6 else 1)' 2>/dev/null; then
      echo "$real"; return 0
    fi
  done
  return 1
}

pyver() { "$1" -c 'import sys;print("Python %d.%d" % sys.version_info[:2])'; }

VENV="$HERE/.venv-build-mac"
VPY="$VENV/bin/python"

echo "[1/7] Finding a Python with a usable Tk..."
PY="$(pick_python)" || {
  echo "      None found. Installing Python and Tk through Homebrew..."
  brew_install python python-tk
  PY="$(pick_python)" || die "Even after installing python and python-tk, no interpreter reports Tk 8.6+."
}
echo "      using $PY ($(pyver "$PY"))"

echo "[2/7] Creating a clean build environment..."
rm -rf "$VENV"
"$PY" -m venv "$VENV" || die "Could not create the build environment."

echo "[3/7] Installing dependencies (wheels only, no compiling)..."
"$VPY" -m pip install --upgrade pip wheel >/dev/null || die "pip would not upgrade (is this Mac online?)."
"$VPY" -m pip install --only-binary :all: -r "$HERE/requirements.txt" || die "A dependency would not install (is this Mac online?)."
"$VPY" -m pip install --only-binary :all: "pyinstaller>=6.11,<7" || die "PyInstaller would not install."

echo "[4/7] Checking the code before packaging..."
"$VPY" -c "import sys, pkgutil, importlib; sys.path.insert(0,'.'); import draftsmanlib, docx, openpyxl; n=[importlib.import_module(m.name) for m in pkgutil.walk_packages(draftsmanlib.__path__,'draftsmanlib.')]; print('      imports ok:', len(n), 'modules')" \
  || die "The application does not import cleanly."

echo "[5/7] Packaging (this takes a minute or two)..."
rm -rf "$HERE/build" "$HERE/dist"
"$VPY" -m PyInstaller --noconfirm --clean --windowed \
  --name "Draftsman" \
  --osx-bundle-identifier "uk.lowther.draftsman" \
  --collect-submodules draftsmanlib \
  --collect-all docx \
  --hidden-import docx \
  --hidden-import openpyxl \
  --hidden-import tkinter \
  --hidden-import tkinter.ttk \
  --hidden-import tkinter.filedialog \
  --hidden-import tkinter.messagebox \
  --hidden-import sqlite3 \
  --exclude-module pytest \
  "$HERE/draftsman_app.py" || die "PyInstaller failed - see the log above."
[ -d "$APP" ] || die "dist/Draftsman.app was not produced."

echo "[6/7] Clearing quarantine and signing locally..."
# Without these two an app built on Apple silicon is killed the moment it opens.
xattr -cr "$APP" || true
codesign --force --deep --sign - "$APP" || die "Ad-hoc signing failed."

echo "[7/7] Running the self-test on the built app..."
echo "      It makes the starter library in a temporary folder and fills every template."
echo "      A Draftsman window will open for a few seconds - leave it alone"
echo "      until it closes by itself."
rm -f "$REPORT"
# The report file is the result; the same text on stdout would only show twice.
"$APP/Contents/MacOS/Draftsman" selftest > /dev/null || true
sleep 2
PROBLEM=""
if [ ! -f "$REPORT" ]; then
  PROBLEM="THE BUILT APP DID NOT RUN ITS SELF-TEST - it probably does not start. Look for dist/draftsman-crash.log."
else
  echo
  cat "$REPORT"
  grep -q "SELF-TEST PASSED" "$REPORT" || PROBLEM="See dist/draftsman-selftest.txt above."
fi
if [ -n "$PROBLEM" ]; then
  echo
  echo "==============================================="
  echo "  BUILT, BUT THE SELF-TEST REPORTED PROBLEMS."
  echo "  $PROBLEM"
  echo "==============================================="
  echo
  echo "Press return to close."
  read -r
  exit 2
fi

echo
echo "==============================================="
echo "  DONE.  dist/Draftsman.app is ready."
echo "  Drag it to your Applications folder."
echo "==============================================="
echo
echo "Press return to close."
read -r
exit 0

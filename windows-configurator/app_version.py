from pathlib import Path
import sys

# Default path next to this module (used during development / installed package)
_LOCAL_VERSION_FILE = Path(__file__).parent / "VERSION"

def _read_file(path: Path) -> str:
    try:
        return path.read_text(encoding='utf-8').strip()
    except Exception:
        return ""

def read_version() -> str:
    """Return version string.

    When running from a PyInstaller one-file bundle, resources are extracted to
    sys._MEIPASS; prefer reading VERSION from there if available. Otherwise
    read the source `VERSION` file next to this module. If both fail, return
    a safe default.
    """
    # If frozen by PyInstaller, check the extracted temp dir first
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        v = _read_file(Path(meipass) / "VERSION")
        if v:
            return v

    # Fallback to local repository file
    v = _read_file(_LOCAL_VERSION_FILE)
    return v or "0.0.0"

def bump_patch():
    v = read_version()
    parts = v.split('.')
    if len(parts) != 3:
        parts = ['0','0','0']
    try:
        parts[2] = str(int(parts[2]) + 1)
    except Exception:
        parts[2] = '0'
    new_v = '.'.join(parts)
    try:
        _LOCAL_VERSION_FILE.write_text(new_v + '\n', encoding='utf-8')
    except Exception:
        pass
    return new_v

def exe_name(base_name: str, version: str) -> str:
    safe = base_name.replace(' ', '_')
    return f"{safe}_v{version}.exe"

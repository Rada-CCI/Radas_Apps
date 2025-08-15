from pathlib import Path

VERSION_FILE = Path(__file__).parent / "VERSION"

def read_version():
    try:
        v = VERSION_FILE.read_text(encoding='utf-8').strip()
        return v
    except Exception:
        return "0.0.0"

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
    VERSION_FILE.write_text(new_v + '\n', encoding='utf-8')
    return new_v

def exe_name(base_name: str, version: str) -> str:
    safe = base_name.replace(' ', '_')
    return f"{safe}_v{version}.exe"

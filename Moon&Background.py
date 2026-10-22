import ctypes
import sys
import time
from datetime import datetime
from pathlib import Path

import win32com.client

INTERVAL_SECONDS = 10

WALLPAPER_STYLE = None

EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp"}

SHCNE_UPDATEITEM = 0x00002000
SHCNF_PATHW = 0x0005

SPI_SETDESKWALLPAPER = 0x0014
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDCHANGE = 0x02

_user32 = ctypes.windll.user32
_user32.SystemParametersInfoW.argtypes = [ctypes.c_uint, ctypes.c_uint, ctypes.c_wchar, ctypes.c_uint]
_user32.SystemParamerterInfoW.restype = ctypes.c_int

def set_wallpaper(path: Path) -> None:
    ok = _user32.SystemParametersInfoW(
        SPI_SETDESKWALLPAPER,
        0,
        str(path.resolve()),
        SPIF_UPDATEINIFILE | SPIF_SENDCHANGE,
    )
    if not ok:
        raise ctypes.WinError()

def aply_style(style: str) -> None:
    import winreg
    styles = {
        "center": ("0", "0"),
        "tile": ("0", "1"),
        "stretch": ("2", "0"),
        "fit": ("6", "0"),
        "fill": ("10", "0"),
        "span": ("22", "0"),
    }
    if style not in styles:
        raise ValueError(f"UNknown WALLPAPER_STYLE: {style!r}")

    wallpaper_style, tile = styles[style]
    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER, r"Control Panel\Desktop", 0, winreg.KEY_SET_VALUE
    ) as key:
        winreg.SetValueEx(key, "WallpaperStyle", 0, winreg.REG_SZ, wallpaper_style)
        winreg.SetValueEx(key, "TileWallpaper", 0, winreg.REG_SZ, tile)

def exit_with_message(message: str, code: int = 1) -> None:
    print(message)
    input("Press enter to exit")
    sys.exit(code)

def main() -> None:
    folder = Path(__file__).resolve().parent / "Wallpapers"
    images = []
    if folder.is_dir():
        images = sorted(
            (p for p in folder.iterdir() if p.is_file and p.suffix.lower() in EXTENSIONS)
        )
    if not images:
        exit_with_message(
            f"No images in {folder}"
        )
    index = 0
    print(f"Cycling {len(images)} wallpapers every {INTERVAL_SECONDS // 60} minutes. "
          "Press Ctrl+C to stop.")

    try:
        while True:
            image = images[index]

            try:
                set_wallpaper(image)
                now = datetime.now().strftime("H%:%M:%S")
                print(f"{now} - Set wallpaper to {image.name}")
            except Exception as exc:
                exit_with_message(f"Failed: {exc}")

            index = (index + 1) % len(images)
            time.sleep(INTERVAL_SECONDS)
    except KeyboardInterrupt:
        print("\nStopped")

if __name__ == "__main__":
    main()
import ctypes
import sys
import time
import argparse
from datetime import datetime
from pathlib import Path

import win32com.client

INTERVAL_SECONDS = 10

WALLPAPER_STYLES = ("fill", "fit", "stretch", "center", "tile", "span")

EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".ico"}

SHCNE_UPDATEITEM = 0x00002000
SHCNF_PATHW = 0x0005

SPI_SETDESKWALLPAPER = 0x0014
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDCHANGE = 0x02

_user32 = ctypes.windll.user32
_user32.SystemParametersInfoW.argtypes = [ctypes.c_uint, ctypes.c_uint, ctypes.c_wchar, ctypes.c_uint]
_user32.SystemParametersInfoW.restype = ctypes.c_int

_shell32 = ctypes.windll.shell32
_shell32.SHChangeNotify.argtypes = [  ctypes.c_long, ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_void_p]
_shell32.SHChangeNotify.restype = None

def set_wallpaper(path: Path) -> None:
    ok = _user32.SystemParametersInfoW(
        SPI_SETDESKWALLPAPER,
        0,
        str(path.resolve()),
        SPIF_UPDATEINIFILE | SPIF_SENDCHANGE,
    )
    if not ok:
        raise ctypes.WinError()

def apply_style(style: str) -> None:
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
        raise ValueError(f"Unknown WALLPAPER_STYLE: {style!r}")

    wallpaper_style, tile = styles[style]
    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER, r"Control Panel\Desktop", 0, winreg.KEY_SET_VALUE
    ) as key:
        winreg.SetValueEx(key, "WallpaperStyle", 0, winreg.REG_SZ, wallpaper_style)
        winreg.SetValueEx(key, "TileWallpaper", 0, winreg.REG_SZ, tile)

def get_images(folder):
    images = []
    if folder.is_dir():
        images = sorted(
            (p for p in folder.iterdir() if p.is_file and p.suffix.lower() in EXTENSIONS)
        )
    if not images:
        exit_with_message(
            f"No images in {folder}"
        )
        return
    return images

def exit_with_message(message: str, code: int = 1) -> None:
    print(message)
    input("Press enter to exit")
    sys.exit(code)


#applystyle?,
def main() -> None:

    parser = argparse.ArgumentParser()  
    parser.add_argument('--style', choices=WALLPAPER_STYLES, default=None)
    parser.add_argument('--symetric', choices={True, False}, default=False) 
    args = parser.parse_args()  
    
    bg_folder = Path(__file__).resolve().parent / "Wallpapers"
    ic_folder = Path(__file__).resolve().parent / "Icons"

    bg_images = get_images(bg_folder)
    ic_images = get_images(ic_folder)

    if args.symetric == True and (len(bg_images) != len(ic_images)):
        exit_with_message(f"The amount of icons:{len(ic_images)} don't mach the background:{len(bg_images)}")

    if args.style:
        try:
            apply_style(args.style)
        except Exception as exc:
            exit_with_message(f"Could not set wallpaper style: {exc}")


    ic_index = 0
    bg_index = 0
    print(f"Cycling {len(bg_images)} wallpapers and {len(ic_images)} icons every {INTERVAL_SECONDS // 60} minutes. "
          "Press Ctrl+C to stop.")

    try:
        while True:
            ic_images = ic_images[ic_index]
            bg_images = bg_images[bg_index]

            try:
                set_wallpaper(bg_images)
                now = datetime.now().strftime("H%:%M:%S")
                print(f"{now} - Set wallpaper to {bg_images.name}")
            except Exception as exc:
                exit_with_message(f"Failed: {exc}")

            index = (index + 1) % len(bg_images)
            time.sleep(INTERVAL_SECONDS)
    except KeyboardInterrupt:
        print("\nStopped")

if __name__ == "__main__":
    main()
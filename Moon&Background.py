import ctypes
import sys
import time
import argparse
import asyncio
from enum import Enum
from datetime import datetime
from pathlib import Path

import win32com.client

WALLPAPER_STYLES = ("fill", "fit", "stretch", "center", "tile", "span")

SHCNE_UPDATEITEM = 0x00002000
SHCNF_PATHW = 0x0005

SPI_SETDESKWALLPAPER = 0x0014
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDCHANGE = 0x02

EXTENSIONS = {
    "IMAGE" : (".jpg", ".jpeg", ".png", ".bmp"),
    "ICON" : (".ico"),
    "SHORTCUT" : (".lnk")
}

_user32 = ctypes.windll.user32
_user32.SystemParametersInfoW.argtypes = [ctypes.c_uint, ctypes.c_uint, ctypes.c_wchar, ctypes.c_uint]
_user32.SystemParametersInfoW.restype = ctypes.c_int

_shell32 = ctypes.windll.shell32
_shell32.SHChangeNotify.argtypes = [  ctypes.c_long, ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_void_p]
_shell32.SHChangeNotify.restype = None

def notify_item_changed(path: Path) -> None:
   ok = 1

def get_parameteres(item):
    interval = None
    while(interval == None):
        interval = int(input(f"How oftern to cycle {item} in seconds? "))
    return interval


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

def get_shortcuts():
    return

def get_files(folder, file_type):
    files = []
    if folder.is_dir():
        ext = EXTENSIONS[file_type]
        files = sorted(
            (p for p in folder.iterdir() if p.is_file() and (p.suffix.lower() in ext)), #Extensions.type contails p sufix 
            key=lambda p:p.name.lower(),
        )
    if not files:
        exit_with_message(
            f"No files in {folder}"
        )
        return
    return files

async def set_wallpaper(images, interval):
    index = 0
    while True:
        image = images[index]
        try:
            set_wallpaper(image)
            print(f"{datetime.now():%H:%M:%S} - Set wallpaper to {image.name}")
        except Exception as exc:
            print(f"Wallpaper failed: {exc}")
            return  # stop this job; the other keeps running

        index = (index + 1) % len(images)
        await asyncio.sleep(interval)

async def set_icon(shell, shortcut_path, icons, interval):
    index = 0
    while True:
        icon = icons[index]
        try:
            shortcut = shell.CreateShortcut(str(shortcut_path))
            shortcut.IconLocation = f"{icon},0"
            shortcut.Save()
            notify_item_changed(shortcut_path)
            print(f"{datetime.now():%H:%M:%S} - Set icon to {icon.name}")
        except Exception as exc:
            print(f"Icon failed: {exc}")
            return

        index = (index + 1) % len(icons)
        await asyncio.sleep(interval)


def exit_with_message(message: str, code: int = 1) -> None:
    print(message)
    input("Press enter to exit")
    sys.exit(code)


async def main() -> None:
    parser = argparse.ArgumentParser()  
    parser.add_argument('--style', choices=WALLPAPER_STYLES, default=None)
    parser.add_argument('--symetric', choices={True, False}, default=False) 
    args = parser.parse_args()  

    bg_folder = Path(__file__).resolve().parent / "Wallpapers"
    ic_folder = Path(__file__).resolve().parent / "Icons"
    shell = win32com.client.Dispatch("WScript.Shell")
    desktop = Path(shell.SpecialFolders("Desktop"))

    bg_images = get_files(bg_folder, "IMAGE")
    ic_images = get_files(ic_folder, "ICON")
    shortcuts = get_files(desktop, "SHORTCUT")

    print("Choose the shortcut to cycle:")
    for i, shortcut_file in enumerate(shortcuts, start = 1):
        print(f"{i}. {shortcut_file.stem}")
    try:
        selection = int(input("Enter its number").strip())
    except ValueError:
        selection = 0

    if not 1 <= selection <= len(shortcuts):
        exit_with_message("Invalid selection.")

    if args.symetric == True and (len(bg_images) != len(ic_images)):
        exit_with_message(f"The amount of icons:{len(ic_images)} don't mach the background:{len(bg_images)}")

    if args.style:
        try:
            apply_style(args.style)
        except Exception as exc:
            exit_with_message(f"Could not set wallpaper style: {exc}")

    if(not args.symetric):
        ic_cycle = get_parameteres("icons")
    bg_cycle = get_parameteres("backgrounds")

    shortcut_path = shortcuts[selection]

    if(args.symetric):
        print(f"Cycling {len(bg_images)} wallpapers and {len(ic_images)} icons every {bg_cycle // 60} minutes. "
            "Press Ctrl+C to stop.")
    else:
        print(f"Cycling {len(bg_images)} wallpapers every {bg_cycle} seconds. "
            "Press Ctrl+C to stop.")
        print(f"Cycling {len(ic_images)} icons every {ic_cycle} seconds. "
                    "Press Ctrl+C to stop.")

    try:
        await asyncio.gather(
            set_wallpaper(bg_images, bg_cycle),
            set_icon(shell, shortcut_path, ic_images, ic_cycle),
        )
    except KeyboardInterrupt:
        print("\nStopped")

if __name__ == "__main__":
    main()

asyncio.run(main())
import argparse
import asyncio
import ctypes
import sys
import time
from datetime import datetime
from pathlib import Path

import win32com.client

WALLPAPER_STYLES = ("fill", "fit", "stretch", "center", "tile", "span")

SHCNE_UPDATEITEM = 0x00002000
SHCNF_PATHW = 0x0005

SPI_SETDESKWALLPAPER = 0x0014
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDCHANGE = 0x02

WALLPAPER_DELAY_SECONDS = 0


EXTENSIONS = {
    "IMAGE": (".jpg", ".jpeg", ".png", ".bmp"),
    "ICON": (".ico",),
    "SHORTCUT": (".lnk",),                              #Trailing comma so whole word gets checked
}

_user32 = ctypes.windll.user32
_user32.SystemParametersInfoW.argtypes = [
    ctypes.c_uint, ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_uint
]
_user32.SystemParametersInfoW.restype = ctypes.c_int

_shell32 = ctypes.windll.shell32
_shell32.SHChangeNotify.argtypes = [
    ctypes.c_long, ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_void_p
]
_shell32.SHChangeNotify.restype = None


def set_wallpaper(path: Path) -> None:
    ok = _user32.SystemParametersInfoW(
        SPI_SETDESKWALLPAPER,                           #Returns true (except if error) 
        0,                                              #Not used
        str(path.resolve()),                            #Holds new value 
        SPIF_UPDATEINIFILE | SPIF_SENDCHANGE,           #Writes to user profile, broadcast message after user update
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
        winreg.SetValueEx(key, "TileWallpaper", 0, winreg.REG_SZ, tile)     #REG_SZ sets these to null term str


def notify_item_changed(path: Path) -> None:
    _shell32.SHChangeNotify(SHCNE_UPDATEITEM, SHCNF_PATHW, str(path), None)


def write_shortcut_icon(shell, shortcut_path: Path, icon: Path) -> None:
    shortcut = shell.CreateShortcut(str(shortcut_path))
    shortcut.IconLocation = f"{icon},0"
    shortcut.Save()


def exit_with_message(message: str, code: int = 1) -> None:
    print(message)
    input("Press Enter to exit")
    sys.exit(code)


def timed_ms(func, *args) -> float:
    start = time.perf_counter()
    func(*args)
    return (time.perf_counter() - start) * 1000


def get_interval(item: str) -> int:
    while True:
        raw = input(f"How often to cycle {item} in seconds? ").strip()
        try:
            value = int(raw)
        except ValueError:
            print("Please enter a whole number.")
            continue
        if value > 0:
            return value
        print("The interval must be greater than 0.")


def get_files(folder: Path, file_type: str) -> list[Path]:
    extensions = EXTENSIONS[file_type]
    files = []
    if folder.is_dir():
        files = sorted(
            (p for p in folder.iterdir()
             if p.is_file() and p.suffix.lower() in extensions),
            key=lambda p: p.name.lower(),
        )
    if not files:
        exit_with_message(f"No {file_type.lower()} files in {folder}")
    return files


async def wallpaper_loop(images, interval):
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


async def icon_loop(shell, shortcut_path, icons, interval):
    index = 0
    while True:
        icon = icons[index]
        try:
            write_shortcut_icon(shell, shortcut_path, icon)
            print(f"{datetime.now():%H:%M:%S} - Set icon to {icon.name}")
        except Exception as exc:
            print(f"Icon failed: {exc}")
            return

        index = (index + 1) % len(icons)
        await asyncio.sleep(interval)


async def notify_after(delay: float, path: Path) -> float:
    await asyncio.sleep(delay)
    return await asyncio.to_thread(timed_ms, notify_item_changed, path)


async def wallpaper_after(delay: float, image: Path) -> float:
    await asyncio.sleep(delay)
    return await asyncio.to_thread(timed_ms, set_wallpaper, image)


async def paired_loop(shell, shortcut_path, images, icons, interval):
    index = 0
    while True:
        image = images[index]
        icon = icons[index]
        try:
            # 1. Write the icon to the file first (silent, instant)
            write_shortcut_icon(shell, shortcut_path, icon)
            # 2. Change the wallpaper and tell Explorer to redraw the icon at the
            #    same moment, each in its own thread (the wallpaper call blocks
            #    while Windows broadcasts the change, so it can't share a thread)
            notify_ms, wallpaper_ms = await asyncio.gather(
                asyncio.to_thread(timed_ms, notify_item_changed, shortcut_path),
                wallpaper_after(WALLPAPER_DELAY_SECONDS, image),
                
            )
            print(
                f"{datetime.now():%H:%M:%S} - Set wallpaper to {image.name} "
                f"and icon to {icon.name} "
                f"(wallpaper call {wallpaper_ms:.0f} ms, icon notify {notify_ms:.0f} ms)"
            )
        except Exception as exc:
            print(f"Failed: {exc}")
            return

        index = (index + 1) % len(images)
        await asyncio.sleep(interval)


async def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--style", choices=WALLPAPER_STYLES, default=None)
    parser.add_argument(
        "--symetric",
        action="store_true",
        help="use one interval for both and require equal numbers of icons and wallpapers",
    )
    args = parser.parse_args()

    base = Path(__file__).resolve().parent
    bg_folder = base / "Wallpapers"
    ic_folder = base / "Icons"
    shell = win32com.client.Dispatch("WScript.Shell")
    desktop = Path(shell.SpecialFolders("Desktop"))

    bg_images = get_files(bg_folder, "IMAGE")
    ic_images = get_files(ic_folder, "ICON")
    shortcuts = get_files(desktop, "SHORTCUT")

    print("Choose the shortcut to cycle:")
    for i, shortcut_file in enumerate(shortcuts, start=1):
        print(f"{i}. {shortcut_file.stem}")
    try:
        selection = int(input("Enter its number: ").strip())
    except ValueError:
        selection = 0

    if not 1 <= selection <= len(shortcuts):
        exit_with_message("Invalid selection.")
    shortcut_path = shortcuts[selection - 1]

    if args.symetric and len(bg_images) != len(ic_images):
        exit_with_message(
            f"The number of icons ({len(ic_images)}) doesn't match "
            f"the number of wallpapers ({len(bg_images)})."
        )

    if args.style:
        try:
            apply_style(args.style)
        except Exception as exc:
            exit_with_message(f"Could not set wallpaper style: {exc}")

    if args.symetric:
        cycle = get_interval("wallpapers and icons")
        print(f"Cycling {len(bg_images)} wallpaper/icon pairs every {cycle} seconds.")
        print("Press Ctrl+C to stop.")
        await paired_loop(shell, shortcut_path, bg_images, ic_images, cycle)
        return

    bg_cycle = get_interval("wallpapers")
    ic_cycle = get_interval("icons")

    print(f"Cycling {len(bg_images)} wallpapers every {bg_cycle} seconds.")
    print(f"Cycling {len(ic_images)} icons every {ic_cycle} seconds.")
    print("Press Ctrl+C to stop.")

    await asyncio.gather(
        icon_loop(shell, shortcut_path, ic_images, ic_cycle),
        wallpaper_loop(bg_images, bg_cycle),
    )


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nStopped.")
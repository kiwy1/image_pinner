import os
import pathlib
import random
import string
import io
import requests
import colorama
from distutils.dir_util import copy_tree
from win32com.client import Dispatch
from PIL import Image

BASE_PATH = pathlib.Path(__file__).parent.resolve()
PROGRAMS_FOLDER = pathlib.Path.home() / "AppData" / "Roaming" / "Microsoft" / "Windows" / "Start Menu" / "Programs"


def create_shortcut(target_path: str, shortcut_path: str, icon_path: str | None = None):
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = target_path
    shortcut.WorkingDirectory = os.path.dirname(target_path)
    if icon_path:
        shortcut.IconLocation = icon_path
    shortcut.Save()


def generate_image(image_path: str, cols: int = 3, rows: int = 2):
    if image_path.startswith("https://"):
        image = Image.open(io.BytesIO(requests.get(image_path, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:135.0) Gecko/20100101 Firefox/135.0"}).content))
    else:
        image = Image.open(image_path)

    tile_path = "".join(random.choices(string.ascii_letters, k=10))
    image_pinner_folder = PROGRAMS_FOLDER / f"ImagePinner_{tile_path}"
    image_pinner_folder.mkdir()
    os.makedirs(BASE_PATH / tile_path)

    image.save(BASE_PATH / tile_path / "orig.png", format="PNG")

    tile_size = min(image.width // cols, image.height // rows)
    resized_image = image.resize((tile_size * cols, tile_size * rows), Image.LANCZOS)

    spaces = 1
    for i in range(cols):
        for j in range(rows):
            box = (i * tile_size, j * tile_size, (i + 1) * tile_size, (j + 1) * tile_size)
            tile = resized_image.crop(box)

            tile_folder = BASE_PATH / tile_path / f"{' ' * spaces}_"
            tile_folder.mkdir(parents=True)
            (tile_folder / "VisualElements").mkdir()
            (image_pinner_folder / f"{' ' * spaces}_").mkdir()

            tile.resize((300, 300)).save(tile_folder / "VisualElements" / f"MediumIcon{' ' * spaces}.png", format="PNG")
            tile.resize((256, 256)).save(tile_folder / "VisualElements" / f"ico{' ' * spaces}.ico", format="ICO")
            tile.resize((150, 150)).save(tile_folder / "VisualElements" / f"SmallIcon{' ' * spaces}.png", format="PNG")

            create_shortcut(str(BASE_PATH / tile_path / f"{' ' * spaces}_" / f"{' ' * spaces}.vbs"), str(image_pinner_folder / f"{' ' * spaces}_" / f"{' ' * spaces}.lnk"), str(tile_folder / "VisualElements" / f"ico{' ' * spaces}.ico"))

            with open(tile_folder / f"{' ' * spaces}.vbs", "w") as f:
                f.write(f'Set objShell = CreateObject("WScript.Shell")\nobjShell.Run """{BASE_PATH / tile_path / "orig.png"}"""')

            with open(tile_folder / f"{' ' * spaces}.VisualElementsManifest.xml", "w") as f:
                f.write(f'''<?xml version="1.0" encoding="utf-8"?>
<Application xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <VisualElements ShowNameOnSquare150x150Logo="on" Square150x150Logo="VisualElements/MediumIcon{' ' * spaces}.png" Square70x70Logo="VisualElements/SmallIcon{' ' * spaces}.png" ForegroundText="light" BackgroundColor="#000000" />
</Application>''')

            spaces += 1


def main():
    colorama.init()
    while True:
        os.system("cls")
        print(colorama.Fore.GREEN + "1. " + colorama.Fore.CYAN + "Open 'Start Menu' folder")
        print(colorama.Fore.GREEN + "2. " + colorama.Fore.CYAN + "Generate image")
        print(colorama.Fore.GREEN + "3. " + colorama.Fore.CYAN + "Exit")
        print("")
        action = input(colorama.Fore.GREEN + "Enter" + colorama.Fore.YELLOW + " > " + colorama.Fore.RESET)

        if action == "1":
            os.startfile(PROGRAMS_FOLDER)
        elif action == "2":
            os.system("cls")
            image_path = input(colorama.Fore.GREEN + "Enter image path OR image URL " + colorama.Fore.YELLOW + "> " + colorama.Fore.RESET)
            cols = input(colorama.Fore.GREEN + "Enter grid cols (DEFAULT: 3) " + colorama.Fore.YELLOW + "> " + colorama.Fore.RESET)
            rows = input(colorama.Fore.GREEN + "Enter grid rows (DEFAULT: 2) " + colorama.Fore.YELLOW + "> " + colorama.Fore.RESET)

            try:
                cols = max(1, int(cols))
            except ValueError:
                cols = 3

            try:
                rows = max(1, int(rows))
            except ValueError:
                rows = 2

            generate_image(image_path, cols, rows)
        elif action == "3":
            break


if __name__ == "__main__":
    main()

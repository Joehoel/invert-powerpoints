
import os
import subprocess
from glob import glob
from os import mkdir, path, rename, rmdir
from shutil import copy, move, rmtree
from tempfile import gettempdir
from zipfile import BadZipFile, ZipFile

from PIL import Image, ImageOps
from tqdm import tqdm


def open_file(filename: str):
    '''Open document with default application in Python.'''
    try:
        os.startfile(filename)  # type: ignore
    except AttributeError:
        subprocess.call(['open', filename])


def file_list(location: str):
    return glob(location, recursive=True)


def create_dir(path: str):
    try:
        mkdir(path)
    except:
        pass


def change_extension(location: str, ext: str):
    base = path.splitext(location)[0]
    file = f"{base}.{ext}"
    rename(location, file)
    return file


def invert_image(image):
    # image = Image.open(file)
    image = image.convert("RGB")

    r, g, b, a = image.split()
    rgb_image = Image.merge('RGB', (r, g, b))

    inverted_image = ImageOps.invert(rgb_image)
    # image = ImageOps.invert(image)

    r2, g2, b2 = inverted_image.split()

    image = Image.merge('RGBA', (r2, g2, b2, a))
    return image


for file in tqdm(file_list("input/**/*.png")):
    try:
        image = Image.open(file)

        image = invert_image(image)

        image.save(path.join("output", path.split(file)[1]), "PNG")

    except KeyboardInterrupt:
        exit()

    except Exception as e:
        print(e)
        continue

import os
import subprocess
from glob import glob
from os import mkdir, path, rename, rmdir
from shutil import copy, move, rmtree
from tempfile import gettempdir
from zipfile import BadZipFile, ZipFile

from PIL import Image, ImageOps
from pptx.api import Presentation  # type: ignore
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR  # type: ignore
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.presentation import Presentation as PPTX
from tqdm import tqdm

TEMP_DIR = gettempdir()


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


# FILE = path.join("input", "powerpoints",
    #  "'k Ben een koninklijk kind (EL 180).pptx")

files = file_list(path.join("input"))

# 0. Clear all previous operations
in_zips_path = path.join(TEMP_DIR, "in-zips")
extracted_path = path.join(TEMP_DIR, "extracted")
if path.exists(in_zips_path):
    rmtree(in_zips_path)
if path.exists(extracted_path):
    rmtree(extracted_path)

rmtree("output")
mkdir("output")

# 1. Convert powerpoint to zip


def convert_to_zip(file: str):
    create_dir(in_zips_path)
    new_file = copy(file, in_zips_path)
    return change_extension(new_file, "zip")


# convert_to_zip(FILE)

# 2. Extract media from zip


def extract_media(file: str):

    create_dir(extracted_path)

    # Create a directory for the all the media
    file_name = path.split(path.splitext(file)[0])[1]
    new_dir = path.join(extracted_path, file_name)
    create_dir(new_dir)

    # Read the zipfile
    with ZipFile(file, "r") as zip:
        # Get a list of all files
        for info in zip.infolist():
            # Check for a media directory
            if("media" in info.filename):
                # Extract all files to a new path
                new_path = zip.extract(
                    info.filename, extracted_path)
                # Move the files
                move(new_path, new_dir)

    rmtree(path.join(extracted_path, "ppt"))
    return new_dir


# extract_media(FILE)
# 3. Remove all white pixels from all media AND invert image


def remove_white(image):
    image = image.convert("RGBA")
    data = image.getdata()

    newData = []
    for item in data:
        if item[0] == 255 and item[1] == 255 and item[2] == 255:
            newData.append((255, 255, 255, 0))
        else:
            newData.append(item)

    image.putdata(newData)
    return image


def invert_image(image):
    # image = Image.open(file)
    image = image.convert("RGBA")

    r, g, b, a = image.split()
    rgb_image = Image.merge('RGB', (r, g, b))

    inverted_image = ImageOps.invert(rgb_image)
    # image = ImageOps.invert(image)

    r2, g2, b2 = inverted_image.split()

    image = Image.merge('RGBA', (r2, g2, b2, a))
    return image


def remove_transparency(im, bg_colour=(255, 255, 255)):

    # Only process if image has transparency (http://stackoverflow.com/a/1963146)
    if im.mode in ('RGBA', 'LA') or (im.mode == 'P' and 'transparency' in im.info):

        # Need to convert to RGBA if LA format due to a bug in PIL (http://stackoverflow.com/a/1963146)
        alpha = im.convert('RGBA').split()[-1]

        # Create a new background image of our matt color.
        # Must be RGBA because paste requires both images have the same format
        # (http://stackoverflow.com/a/8720632  and  http://stackoverflow.com/a/9459208)
        bg = Image.new("RGBA", im.size, bg_colour + (255,))
        bg.paste(im, mask=alpha)
        return bg

    else:
        print("Cant remove transparency")
        return im
# media = glob(path.join(TEMP_DIR, "extracted", "**/*"))

# for file in tqdm(media):
#     image = remove_white(file)
#     image = invert_image(file)
#     image.save(file)


# 5. Replace the new images with the old images in the zip


def replace_media(folder: str):
    file = f"{folder}.zip"

    input_path = path.join(TEMP_DIR, "in-zips", file)

    out_folder = path.join(TEMP_DIR, "in_zips")

    create_dir(out_folder)
    output_path = path.join(out_folder, file)

    with ZipFile(input_path, "r") as in_zip:
        with ZipFile(output_path, "w") as out_zip:
            out_zip.comment = in_zip.comment
            for info in in_zip.infolist():
                if("media" in info.filename):
                    location = path.split(info.filename)[1]
                    new_location = path.join(
                        TEMP_DIR, "extracted", folder, location)

                    out_zip.write(
                        new_location, info.filename)
                else:
                    out_zip.writestr(
                        info, in_zip.read(info.filename))
    return output_path


# FOLDER = path.splitext(path.split(FILE)[1])[0]

# out_zip = replace_media(FOLDER)

# 6. Convert the zip folders back to powerpoints


def convert_to_pptx(file: str):
    dir = path.join("output")
    create_dir(dir)
    new_file = copy(file, dir)
    return change_extension(new_file, "pptx")


# pptx_file = convert_to_pptx(out_zip)

# 7. Update the powerpoint colors


def modify_pptx(file: str):
    pptx: PPTX = Presentation(file)

    for slide in pptx.slides:  # type: ignore
        try:
            # Update background color
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.patterned()
            fill.back_color.rgb = RGBColor(0, 0, 0)

            # for shape in slide.placeholders:
            #     print('%d %s' % (shape.placeholder_format.idx, shape.name))

            for shape in slide.shapes:
                if shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:
                        for run in p.runs:
                            run.font.color.rgb = RGBColor(255, 255, 255)
                        # p.font.color.theme_color = MSO_THEME_COLOR.DARK_1
                # if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    # shape.element.fill.fore_color.rgb = RGBColor(255, 255, 255)
                    # shape.part.source.fill.fore_color.rgb = RGBColor(
                    #     255, 255, 255)

        except KeyError:
            continue

    name = path.split(file)[1]
    pptx.save(path.join("output", name))


# modify_pptx(pptx_file)

# print(TEMP_DIR)

for file in tqdm(file_list("input/**/*.pptx")):
    try:
        # print('\n', file)

        convert_to_zip(file)
        directory = extract_media(file)

        media = file_list(path.join(directory, "**/*.png"))

        for item in media:
            image = Image.open(item)

            image = remove_white(image)
            image = invert_image(image)
            image.save(item, "PNG")

        folder = path.splitext(path.split(file)[1])[0]

        out_zip = replace_media(folder)
        pptx_file = convert_to_pptx(out_zip)
        modify_pptx(pptx_file)

    except KeyboardInterrupt:
        exit()

    except Exception as e:
        print(e)
        continue

file = file_list(path.join("output", "**/*.pptx"))[0]

open_file(file)

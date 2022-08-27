import os
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor
import subprocess
import tempfile
from glob import glob
from multiprocessing import Pool, cpu_count
from os import mkdir, path, rename, rmdir
from shutil import copy, move, rmtree
from tempfile import gettempdir
from threading import Thread
from zipfile import BadZipFile, ZipFile

from alive_progress import alive_bar
from PIL import Image, ImageOps
from cv2 import Mat, bitwise_not, imread, imwrite
from pptx.api import Presentation  # type: ignore
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR  # type: ignore
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.presentation import Presentation as PPTX
import numpy as np
import cv2
import json
from matplotlib import pyplot as plt
# TEMP_DIR = gettempdir()


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

# files = file_list(path.join("input"))

# # 0. Clear all previous operations
# in_zips_path = path.join(TEMP_DIR, "in-zips")
# extracted_path = path.join(TEMP_DIR, "extracted")
# if path.exists(in_zips_path):
#     rmtree(in_zips_path)
# if path.exists(extracted_path):
#     rmtree(extracted_path)

def clear_output():
    rmtree("output")
    mkdir("output")

# 1. Convert powerpoint to zip


in_zips_path = tempfile.mkdtemp()
extracted_path = tempfile.mkdtemp()


def convert_to_zip(file: str):
    # create_dir(in_zips_path)
    new_file = copy(file, in_zips_path)
    return change_extension(new_file, "zip")


# convert_to_zip(FILE)

# 2. Extract media from zip


def extract_media(file: str):

    # create_dir(extracted_path)

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
    inverted_image = bitwise_not(image)
    return inverted_image
    # image = Image.open(file)
    # image = image.convert("L")

    # r, g, b, a = image.split()
    # rgb_image = Image.merge('RGB', (r, g, b))

    # inverted_image = ImageOps.invert(rgb_image)

    # r2, g2, b2 = inverted_image.split()

    # image = Image.merge('RGBA', (r2, g2, b2, a))
    # return image


def remove_transparency(image, bg_colour=(255, 255, 255)):
    alpha = image.convert('RGBA').split()[-1]
    bg = Image.new("RGBA", image.size, bg_colour + (255,))
    bg.paste(image, mask=alpha)
    return bg
    # trans_mask = image[:, :, 3] == 0

    # # replace areas of transparency with white and not transparent
    # image[trans_mask] = [255, 255, 255, 255]

    # # new image without alpha channel...
    # new_img = cv2.cvtColor(image, cv2.COLOR_BGRA2BGR)
    # return new_img
    # Only process if image has transparency (http://stackoverflow.com/a/1963146)
    # if image.mode in ('RGBA', 'LA') or (image.mode == 'P' and 'transparency' in image.info):
    #     print(image.info)
    #     # Need to convert to RGBA if LA format due to a bug in PIL (http://stackoverflow.com/a/1963146)
    #     alpha = image.convert('RGBA').split()[-1]

    #     # Create a new background image of our matt color.
    #     # Must be RGBA because paste requires both images have the same format
    #     # (http://stackoverflow.com/a/8720632  and  http://stackoverflow.com/a/9459208)
    #     bg = Image.new("RGBA", image.size, bg_colour + (255))
    #     bg.paste(image, mask=alpha)
    #     return bg

    # else:
    #     pass
    # return remove_transparency(image.convert("P"))
    # media = glob(path.join(TEMP_DIR, "extracted", "**/*"))

    # for file in tqdm(media):
    #     image = remove_white(file)
    #     image = invert_image(file)
    #     image.save(file)

    # 5. Replace the new images with the old images in the zip


def replace_media(folder: str):
    file = f"{folder}.zip"

    input_path = path.join(in_zips_path, file)

    # out_folder = path.join(extracted_path, "out_folder")

    # create_dir(out_folder)
    output_path = path.join(extracted_path, file)

    with ZipFile(input_path, "r") as in_zip:
        with ZipFile(output_path, "w") as out_zip:
            out_zip.comment = in_zip.comment
            for info in in_zip.infolist():
                if("media" in info.filename):
                    location = path.split(info.filename)[1]
                    new_location = path.join(
                        extracted_path, folder, location)

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
    return path.join("output", name)

# modify_pptx(pptx_file)

# print(TEMP_DIR)


def white_to_transparency(img):
    x = np.asarray(img.convert('RGBA')).copy()

    x[:, :, 3] = (0 * (x[:, :, :3] != 0).any(axis=2)).astype(np.uint8)

    return Image.fromarray(x)


def convert_image(file: str):
    image = Image.open(file)

    image = image.convert("RGB")

    image = ImageOps.invert(image)
    # pixels = image.load()
    # width, height = image.size

    # for y in range(height):
    #     for x in range(width):
    #         if pixels[x, y] == (0, 0, 0, 255):
    #             pixels[x, y] = (255, 255, 255, 255)
    #         elif pixels[x, y] == (255, 255, 255, 255):
    #             pixels[x, y] = (0, 0, 0, 0)
    image.save(file, "PNG")
    # image = remove_transparency(image)

    # image.convert("RGB")

    # image = ImageOps.invert(image)

    # image = imread(file, 0)

    # image = invert_image(image)

    # imwrite(file, image)
    # image.save(file)
    # image = invert_this(file, gray_scale=True)
    # image = remove_transparency(image)

    # cv2.imwrite(file, image)
    # image.save(file, "PNG")


def convert_pptx(file: str):

    convert_to_zip(file)
    directory = extract_media(file)

    media = file_list(path.join(directory, "**/*.png"))

    # with concurrent.futures.ProcessPoolExecutor() as executor:
    #     executor.map(convert_image, media)
    for item in media:
        convert_image(item)
        # image = Image.open(item)

        # image = invert_image(image)
        # image.save(item, "PNG")

    folder = path.splitext(path.split(file)[1])[0]

    out_zip = replace_media(folder)
    pptx_file = convert_to_pptx(out_zip)

    modify_pptx(pptx_file)

    return file


def read_image(image_file: str, gray_scale=False):
    image_src = cv2.imread(image_file)
    if gray_scale:
        image_src = cv2.cvtColor(image_src, cv2.COLOR_BGR2GRAY)
    else:
        image_src = cv2.cvtColor(image_src, cv2.COLOR_BGR2RGB)
    return image_src


def invert_this(image_file: str, gray_scale=False):
    image_src = read_image(image_file=image_file, gray_scale=gray_scale)
    # image_i = ~ image_src
    image_i = 255 - image_src

    return image_i


if __name__ == "__main__":
    try:
        clear_output()

        files = file_list("input/**/*.pptx")

        with Pool() as pool:
            with alive_bar(len(files), dual_line=True,  title="Converting powerpoints...") as bar:
                for file in pool.imap_unordered(convert_pptx, files):
                    bar.text = f"Converted '{file}'..."
                    bar()

        # threads = []

        # for file in files:
        #     threads.append(Thread(target=convert_pptx, args=[file, ]))

        # for thread in threads:
        #         thread.start()
        #         bar()
            # with ThreadPoolExecutor(cpu_count()) as executor:
            #     for _ in executor.map(convert_pptx, files):

        # with Pool(cpu_count()) as pool:
            # for _ in pool.imap_unordered(convert_pptx, files):
            # bar()

        file = file_list(path.join("output", "**/*.pptx"))[0]

        open_file(file)

        # with concurrent.futures.ProcessPoolExecutor() as executor:
        # results = executor.map(convert_pptx,
        #                        file_list("input/**/*.pptx"))
        # for result in results:
        #     if results != None:
        #         print(result)

        # for file in tqdm(file_list("input/**/*.pptx")):
        # convert_pptx(file)

        # with Pool() as pool:
        #     results = pool.imap(
        #         convert_pptx, file_list("input/**/*.pptx"))

        #     for result in tqdm(results):
        #         print(result)
    except Exception as e:
        print(e)

# for file in tqdm(file_list("input/**/*.pptx")):
#     try:
#         # print('\n', file)

#         with Pool() as pool:
#             results = pool.imap_unordered(
#                 convert_pptx, tqdm(file_list("input/**/*.pptx")))

#         # convert_pptx(file)

#     except KeyboardInterrupt:
#         exit()

#     except Exception as e:
#         print(e)
#         continue

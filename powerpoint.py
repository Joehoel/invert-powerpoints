from os import rename, path, rmdir, listdir, mkdir
from glob import glob
from shutil import copy, move
from zipfile import BadZipFile, ZipFile

from PIL import Image, ImageOps

from pptx.dml.color import RGBColor
from pptx import Presentation, PresentationPart

from tqdm import tqdm


def hex_to_rgb(input: str):
    h = input.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def change_extension(location: str, ext: str):
    # print(f"✏️ Changing extension from file: {path} to .{ext}")

    base = path.splitext(location)[0]
    rename(location, f"{base}.{ext}")


def generate_zips(location: str, ext: str):
    # print(f"🏭 Generating zips from {location} with {ext} extension")

    for file in tqdm(file_list(f"{location}/**/*.{ext}")):
        try:
            create_dir(f"{location}/../zips")
            new_file = copy(file, f"{location}/../zips")
            change_extension(new_file, "zip")
        except FileExistsError:
            pass
            # print(f"🟧 File: already exist")


def extract_media(location: str):
    # # print(f"👽 Extracting media from {location}...")

    try:
        file_name = path.split(path.splitext(location)[0])[1]

        new_dir = path.join("output", "images", file_name)
        mkdir(new_dir)
        # create_dir(new_dir)

        with ZipFile(location, "r") as zip:
            for info in zip.infolist():
                if("media" in info.filename):
                    new_path = zip.extract(
                        info.filename, path.join("output", "images"))
                    move(new_path, new_dir)

        rmdir(path.join("output", "images", "ppt", "media"))
        rmdir(path.join("output", "images", "ppt"))

    except Exception as e:
        # print(e)
        print("🚫 Something wen't wrong")


def file_list(location: str):
    return glob(location, recursive=True)


def remove_white(image: Image):
    image = image.convert("RGBA")
    data = image.getdata()

    newData = []
    for item in data:
        if item[0] == 255 and item[1] == 255 and item[2] == 255:
            newData.append((255, 255, 255, 0))
        else:
            newData.append(item)

    image.putdata(newData)

    return image.convert("RGB")


def remove_alpha(image: Image):
    # # print("Removing alpha from image...")

    image = image.convert("RGBA")
    r, g, b, a = image.split()
    rgb_image = Image.merge('RGB', (r, g, b))

    inverted_image = ImageOps.invert(rgb_image)

    r2, g2, b2 = inverted_image.split()

    image = Image.merge('RGBA', (r2, g2, b2, a))

    return image


def clear_output():
    for dir in file_list("./output/*"):
        rmdir(dir)


def invert_image(image: Image):
    # print("♻️ Inverting image...")

    # rgb_image = remove_alpha(image)

    # r, g, b, a = image.split()
    # rgb_image = Image.merge('RGB', (r, g, b))

    image = image.convert("RGBA")
    r, g, b, a = image.split()
    rgb_image = Image.merge('RGB', (r, g, b))

    inverted_image = ImageOps.invert(rgb_image)

    r2, g2, b2 = inverted_image.split()

    image = Image.merge('RGBA', (r2, g2, b2, a))

    # inverted_image = ImageOps.invert(remove_alpha(image))

    return image


def save_image(location: str, image: Image, out_dir: str):
    # new_name = path.join(out_dir, path.split(path)[-1])
    new_dir = path.join(out_dir, path.split(path.split(location)[0])[1])
    new_file = path.split(location)[-1]
    new_path = path.join(new_dir, new_file)

    try:
        mkdir(new_dir)
        # print(f"📂 Creating directory: {new_dir}")
    except:
        pass

    image.save(new_path)
    # print("💾 Saving Inverted", location, "as", new_path)


def create_dir(path: str):
    try:
        mkdir(path)
    except:
        pass
        # print(f"🟧 Directory: {path} already exist")


def overwrite_images_in_zip(file: str):
    # print(f"🖨️ Overwriting zip: {file} with converted images")
    with ZipFile(f"input/zips/{file}.zip", "r") as in_zip:
        with ZipFile(f"output/zips/{file}.zip", "w") as out_zip:
            out_zip.comment = in_zip.comment
            for info in in_zip.infolist():
                if("media" in info.filename):
                    location = path.split(info.filename)[1]
                    out_zip.write(
                        f"./output/converted/{file}/{location}", info.filename)
                else:
                    out_zip.writestr(
                        info, in_zip.read(info.filename))


def move_files():
    new_file = copy(
        f"./output/zips/{directory}.zip", "./output/powerpoints")
    change_extension(new_file, "pptx")


def modify_presentation(file: str):
    # print(f"✒️ Modifying: ./output/powerpoints/{file}.pptx")
    presentation = Presentation(
        f"./output/powerpoints/{file}.pptx")

    for slide in presentation.slides:
        try:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:
                        p.font.color.rgb = RGBColor(255, 255, 255)
                        # run = p.add_run()
                        # run.font.color.rgb = RGBColor(255, 255, 255)

            background = slide.background
            fill = background.fill
            fill.solid()
            fill.patterned()
            fill.back_color.rgb = RGBColor(0, 0, 0)

        except KeyError:
            # print(f"🗑️ Corrupt file: {file}.pptx")
            continue

    presentation.save(f"./output/powerpoints/{file}.pptx")


if __name__ == "__main__":
    try:
        create_dir("output/converted")
        create_dir("output/zips")
        create_dir("output/powerpoints")

        generate_zips("input/powerpoints", "pptx")

        for file in tqdm(file_list("input/zips/**/*.zip")):
            extract_media(file)

        for file in tqdm(file_list("output/images/**/*.png")):
            # Open image
            image = Image.open(file)
            # Invert colors
            image = invert_image(image)

            # Save image
            save_image(file, image, path.join("output", "converted"))

        directories = listdir("output/converted")

        for directory in tqdm(directories):
            overwrite_images_in_zip(directory)
            move_files()
            modify_presentation(directory)

    except BadZipFile:
        pass
        # print("🤐 Zip error")

    except KeyboardInterrupt:
        print("🚫 Cancelled")

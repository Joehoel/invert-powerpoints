import os
from glob import glob
from shutil import copy, move
from zipfile import BadZipFile, ZipFile

from PIL import Image, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE


def change_extension(path: str, ext: str):
    print(f"✏️ Changing extension from file: {path} to .{ext}")

    base = os.path.splitext(path)[0]
    os.rename(path, f"{base}.{ext}")


def generate_zips(path: str, ext: str):
    print(f"🏭 Generating zips from {path} with {ext} extension")

    for file in file_list(f"{path}\\**\\*.{ext}"):
        try:
            new_file = copy(file, f"{path}/../zips")
            change_extension(new_file, "zip")
        except FileExistsError:
            print(f"🟧 File: {new_file} already exist")


def extract_media(path: str):
    print(f"👽 Extracting media from {path}...")

    try:
        file_name = os.path.split(os.path.splitext(path)[0])[1]

        new_dir = os.path.join("output", "images", file_name)
        create_dir(new_dir)

        with ZipFile(path, "r") as zip:
            for info in zip.infolist():
                if("media" in info.filename):
                    new_path = zip.extract(
                        info.filename, os.path.join("output", "images"))
                    move(new_path, new_dir)

        os.rmdir(os.path.join("output", "images", "ppt", "media"))
        os.rmdir(os.path.join("output", "images", "ppt"))

    except Exception as e:
        print(e)
        print("🚫 Something wen't wrong")


def file_list(path: str):
    return glob(path, recursive=True)


def invert_image(image: Image):
    print("♻️ Inverting image...")

    image = image.convert("RGBA")
    r, g, b, a = image.split()
    rgb_image = Image.merge('RGB', (r, g, b))

    inverted_image = ImageOps.invert(rgb_image)

    r2, g2, b2 = inverted_image.split()

    image = Image.merge('RGBA', (r2, g2, b2, a))

    return image


def save_image(path: str, image: Image, out_dir: str):
    # new_name = os.path.join(out_dir, os.path.split(path)[-1])
    new_dir = os.path.join(out_dir, os.path.split(os.path.split(path)[0])[1])
    new_file = os.path.split(path)[-1]
    new_path = os.path.join(new_dir, new_file)

    try:
        os.mkdir(new_dir)
        print(f"📂 Creating directory: {new_dir}")
    except:
        pass

    image.save(new_path)
    print("💾 Saving Inverted", path, "as", new_path)


def create_dir(path: str):
    try:
        os.mkdir(path)
    except:
        print(f"🟧 Directory: {path} already exist")


def overwrite_images_in_zip(file: str):
    print(f"🖨️ Overwriting zip: {file} with converted images")
    with ZipFile(f"input\\zips\\{file}.zip", "r") as in_zip:
        with ZipFile(f"output\\zips\\{file}.zip", "w") as out_zip:
            out_zip.comment = in_zip.comment
            for info in in_zip.infolist():
                if("media" in info.filename):
                    path = os.path.split(info.filename)[1]
                    out_zip.write(
                        f"./output/converted/{file}/{path}", info.filename)
                else:
                    out_zip.writestr(
                        info, in_zip.read(info.filename))


def move_files():
    new_file = copy(
        f"./output/zips/{directory}.zip", "./output/powerpoints")
    change_extension(new_file, "pptx")


def modify_presentation(file: str):
    print(f"✒️ Modifying: ./output/powerpoints/{file}.pptx")
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
            print(f"🗑️ Corrupt file: {file}.pptx")
            continue

    presentation.save(f"./output/powerpoints/{file}.pptx")


if __name__ == "__main__":
    try:
        # generate_zips("input/powerpoints", "pptx")
        # for file in file_list("input\\zips\\**\\*.zip"):
        #     extract_media(file)

        for file in file_list("output\\images\\**\\*.png"):
            # Open image
            image = Image.open(file)
            # Invert colors
            image = invert_image(image)
            # Save image
            save_image(file, image, os.path.join("output", "converted"))

        # directories = os.listdir("output/converted")

        # create_dir("output\\zips")
        # create_dir("output\\powerpoints")

        # for directory in directories:
        #     overwrite_images_in_zip(directory)
        #     move_files()
        #     modify_presentation(directory)

    except BadZipFile:
        print("🤐 Zip error")

    except KeyboardInterrupt:
        print("🚫 Cancelled")

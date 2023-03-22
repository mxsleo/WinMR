#!/usr/bin/env python3


"""

Win Media Renamer v2

Renames images by datetime

TODO: Logging options
TODO: An option to remove empty directories
TODO: Video renamer
TODO: An option to rename images by size with naming options supported

"""


# ----------
# Imports
# ----------


from os import listdir, rename, mkdir

from os.path import abspath, exists, isdir, isfile, join, getmtime

from PIL.Image import open as image_open

from time import struct_time, localtime, strptime, strftime, gmtime, mktime


# ----------
# Constants
# ----------


# An option which controls if the program will enter subfolders
use_recursion: bool = True

# Timezone offset (hours)
datetime_gmt: int = 3

# Name and postfix formats
format_name: str = "%Y-%m-%d %H-%M-%S"
format_postfix: str = "_{}"

# If set to True
# in case if Win and EXIF datetimes are different
# will put both the values into a new filename
# else will prefer EXIF datetime
use_dual: bool = False

# Remove empty directoies
del_empty: bool = False

# Dictionary of supported extensions
# and target extension
# str.lower() should be used on a file extension
supported_extensions: dict = {
    "jpg": "jpg",
    "jpeg": "jpg",
    "png": "png"
}

# EXIF datetime tag:
# https://www.awaresystems.be/imaging/tiff/tifftags/privateifd/exif/datetimeoriginal.html
exif_tag_datetime: int = 0x9003

# Warning datetime limits
datetime_warning_low: struct_time = strptime(
    "2014-01-01 00-00-00", "%Y-%m-%d %H-%M-%S")
datetime_warning_high: struct_time = localtime()

# EXIF and Win Datetime maximal difference in seconds
eps: float = 10.0


# ----------
# Functions
# ----------


# Gets a folder path, checks it and turns into an absolute path
def get_folder_path(msg: str) -> str:

    path_folder_raw: str = input("Enter path to the {} folder: ".format(msg))

    path_folder_raw = path_folder_raw.strip('"')

    if path_folder_raw == "":
        return ""

    if not isdir(path_folder_raw):
        raise NotADirectoryError(path_folder_raw)

    return abspath(path_folder_raw)


# Returns datetime in format YYYY-MM-DD hh-mm-ss
def get_image_datetime(path_filename_extension: str) -> str:

    # Get Win datetime with timezone offset
    # Extract struct_time and printable string
    # Check the limits
    image_datetime_windows_float: float = getmtime(
        path_filename_extension) + datetime_gmt * 60 * 60

    image_datetime_windows: struct_time = gmtime(image_datetime_windows_float)
    image_datetime_windows_printable: str = strftime(
        format_name, image_datetime_windows)

    if (image_datetime_windows < datetime_warning_low or
            image_datetime_windows > datetime_warning_high):
        print("WDOR Warning: Win datetime out of range: \"{}\"".format(
            image_datetime_windows_printable))

    # Try to get EXIF datetime
    image_datetime_exif_raw: str = None
    with image_open(path_filename_extension) as image:
        image_exif = image._getexif()
        if image_exif is not None:
            image_datetime_exif_raw = image_exif.get(exif_tag_datetime)

    # Return Win datetime if EXIF is empty
    if image_datetime_exif_raw is None:
        print("E404 Warning: EXIF datetime is empty")
        return image_datetime_windows_printable

    # Get EXIF datetime with timezone offset
    # Extract struct_time and printable string
    # Check the limits
    image_datetime_exif_float: float = mktime(strptime(
        image_datetime_exif_raw, "%Y:%m:%d %H:%M:%S")) + datetime_gmt * 60 * 60

    image_datetime_exif: struct_time = gmtime(image_datetime_exif_float)
    image_datetime_exif_printable: str = strftime(
        format_name, image_datetime_exif)

    if (image_datetime_exif < datetime_warning_low or
            image_datetime_exif > datetime_warning_high):
        print("EDOR Warning: EXIF datetime out of range: \"{}\"".format(
            image_datetime_exif_printable))
        return image_datetime_windows_printable

    # Check if Win and EXIF datetimes are different
    if abs(image_datetime_windows_float - image_datetime_exif_float) > eps:
        print("DIFF Warning: EXIF datetime \"{}\""
              " differs too much from Win datetime \"{}\"".format(
                  image_datetime_exif_printable,
                  image_datetime_windows_printable))
        if use_dual:
            return "{} ({})".format(
                image_datetime_windows_printable,
                image_datetime_exif_printable)
        return image_datetime_exif_printable

    return image_datetime_exif_printable


# Returns a postfix in the specified format if the same file exists
def get_postfix(path_target: str, filename: str, extension: str) -> str:

    postfix: int = 0
    postfix_str: str = ""

    new_filename_extension: str = "{}.{}".format(filename, extension)
    new_path_filename_extension: str = join(
        path_target, new_filename_extension)

    while exists(new_path_filename_extension):

        postfix += 1
        postfix_str = format_postfix.format(postfix)

        new_filename_extension = "{}{}.{}".format(
            filename, postfix_str, extension)
        new_path_filename_extension = join(
            path_target, new_filename_extension)

    if postfix > 0:
        print("PFIX Warning: Postfix added: \"{}\"".format(postfix_str))

    return postfix_str


# Walks through the files in the source directory
# and moves them to the target directory renamed with datetime format
def rename_images_by_datetime(path_src: str, path_target: str) -> int:

    for filename_extension in listdir(path_src):

        print('-' * 100)

        path_filename_extension: str = join(path_src, filename_extension)

        if isfile(path_filename_extension):

            extension: str = filename_extension.split('.')[-1].lower()

            if extension in supported_extensions:

                new_extension: str = supported_extensions[extension]
                image_datetime: str = get_image_datetime(
                    path_filename_extension)
                postfix: str = get_postfix(
                    path_target, image_datetime, new_extension)

                new_filename_extension: str = "{}{}.{}".format(
                    image_datetime, postfix, new_extension)
                new_path_filename_extension: str = join(
                    path_target, new_filename_extension)

                print("\t\"{}\"->\"{}\"".format(filename_extension,
                      new_filename_extension))
                rename(path_filename_extension, new_path_filename_extension)

            else:
                print("Unsupported extension: \"{}\"".format(filename_extension))

        elif use_recursion and isdir(path_filename_extension):

            print("Entering directory: \"{}\"".format(filename_extension))

            new_path_filename_extension = join(path_target, filename_extension)

            if not exists(new_path_filename_extension):
                mkdir(new_path_filename_extension)

            rename_images_by_datetime(
                path_filename_extension, new_path_filename_extension)

        else:
            print("Not a file or directory: \"{}\"".format(filename_extension))

    return 0


# ----------
# Functions
# ----------


if __name__ == "__main__":
    # Get the images folder and the target folder paths
    # if the target folder path is empty, attempt to use the images folder one
    path_src: str = get_folder_path("images")
    path_target: str = get_folder_path("target")

    if path_target == "":
        path_target = path_src

    if path_src == path_target:
        print("Warning: the images folder is the same as the target folder")
        input("Press any key to continue")

    # Main
    rename_images_by_datetime(path_src, path_target)

    exit(0)

#!/usr/bin/env python3

"""
Win Media Renamer
Renames mediafiles by datetime

TODO:
- Logfile support
- Settingsfile support
- "Rename by resolution" support
- UNIX support
"""

# Imports

from os import name as osname, listdir, mkdir, rmdir, rename
from os.path import isdir, abspath, join, isfile, exists, getmtime
from time import struct_time, localtime, strftime, strptime, mktime
from PIL.Image import Image, Exif, open as image_open
try:
    from _win32typing import PyIPropertyStore  # type: ignore
except ModuleNotFoundError:
    print("Ignored: ModuleNotFoundError: No module named '_win32typing'")
from win32com.propsys import propsys, pscon  # type: ignore

# Constants

"""
WinMR Constants
Strictly recommended to be leaved untouched

exif_tag_dt - The EXIF datetime tag ID
ext_img - Supported image extensions with the target ones
ext_vid - Supported video extensions with the target ones
ext_sup - Supported extensions with the target ones
"""

exif_tag_dt: int = 0x9003

ext_img: dict = {
    "jpg": "jpg",
    "jpeg": "jpg",
    "png": "png"
}
ext_vid: dict = {
    "mp4": "mp4",
    "mov": "mov"
}
ext_sup: dict = ext_img | ext_vid

# Options

"""
WinMR options

recurse - If set to True, the program will work on subdirectories too
del_empty - If set to True, the program will delete empty directories
dual - If set to True, the program will put bth datetimes when high difference

fmt_pfx - Postfix format string
fmt_dt - Datetime format string

separate - If set to True, separators will be printed
splitter - A separator

w_pause - If set to True, the program will pause on every warning
w_stps - Controls if "the same directory" warning will be shown
w_apfx - Controls if "the postfix added" warning will be shown
w_wdor - Controls if "Win datetime out of range" warning will be shown
w_wdor_low - Lower datetime range warning limit
w_wdor_high - Higher datetime range warning limit
w_e404 - Controls if "EXIF/datetime not found" warning will be shown
w_edor - Controls if "EXIF datetime out of range" warning will be shown
w_diff - Controls if "Datetimes differ too much" warning will be shown
w_diff_eps - Max datetimes difference in seconds
"""

recurse: bool = True
del_empty: bool = True
dual: bool = False

fmt_pfx: str = "_{}"
fmt_dt: str = "%Y-%m-%d %H-%M-%S"

separate: bool = False
splitter: str = '-' * 100

w_pause: bool = False
w_stps: bool = True
w_apfx: bool = True
w_wdor: bool = True
w_wdor_low: struct_time = strptime("2015-01-01 00-00-00", "%Y-%m-%d %H-%M-%S")
w_wdor_high: struct_time = localtime()
w_e404: bool = True
w_edor: bool = True
w_diff: bool = True
w_diff_eps: float = 10.0

# Functions


def pause() -> None:
    """
    Pause function

    If allowed, pause the program until Enter is pressed
    """

    if w_pause:
        input("Press Enter to continue")


def splitout() -> None:
    """
    Separator printer

    If allowed, print the separator
    """

    if separate:
        print(splitter)


def get_path(msg: str) -> str | None:
    """
    Path reader and unifier

    Returns an absolute path or None if is is empty

    Input a directory path without special symbols
    If the entered path is empty return None
    Check if the path is actually a path to a directory
    Return absolute directory path

    msg - Folder description

    rpath - "Raw" directory path
    """

    rpath: str | None = None

    rpath = input("Enter path to the {} directory: ".format(msg)).strip('"'' ')

    if rpath == "":
        return None

    if not isdir(rpath):
        raise NotADirectoryError(rpath)

    return abspath(rpath)


def get_pfx(fn_ext: str, path: str) -> str:
    """
    Postfix generator

    Returns first available postfix
    for the specified filename in the specified path
    or an empty string if the same file does not exist at all

    Get a filename
    Get an extension
    Default target Filename + Postfix + Extension
    to original Filename + Extension
    Default Postfix to an empty string
    to return it if the same file does not exist at all
    While target Filename + Postfix + Extension has a clone:
    Increase postfix number
    Get new target Filename + Postfix + Extension

    Little bit complicated just for portability

    fn_ext - Filename + Extension
    path - Path for Filename + Extension postfix search

    fn - Filename
    ext - Extension
    t_fn_pfx_ext - Target Filename + Postfix + Extension
    pfx_num - Postfix (numeric)
    pfx - Postfix (printable)
    """

    fn: str | None = None
    ext: str | None = None

    t_fn_pfx_ext: str = fn_ext
    pfx_num: int = 0
    pfx: str = ""

    fn = ''.join(fn_ext.split('.')[:-1])
    ext = fn_ext.split('.')[-1].lower()

    while exists(join(path, t_fn_pfx_ext)):
        pfx_num += 1
        pfx = fmt_pfx.format(pfx_num)

        t_fn_pfx_ext = "{}{}.{}".format(fn, pfx, ext)

    if w_apfx and (pfx_num > 0):
        print("APFX Warning: postfix added ({})".format(pfx))
        pause()

    return pfx


def get_dt_img(path: str) -> str | None:
    """
    Image datetime reader

    Returns an image datetime in the specified format
    or None if the image does not have EXIF

    With an image opened
    Try to get an image EXIF
    If the EXIF is absent, warn user and return None
    Try to get an EXIF datetime
    If the EXIF datetimeis is absent, warn user and return None
    Convert it to struct_time
    Check the limits
    Return the datetime in printable format

    path - Path to the required file

    image - Image
    exif - Image exif
    dt_exif_r - "Raw" EXIF datetime
    dt_exif - EXIF datetime
    """

    image: Image | None = None
    exif: Exif | None = None
    dt_exif_r: str | None = None
    dt_exif: struct_time | None = None

    with image_open(path) as image:
        exif = image._getexif()  # type: ignore
    if exif is None:
        if w_e404:
            print("E404 Warning: Exif not found")
            pause()
        return None

    dt_exif_r = exif.get(exif_tag_dt)
    if dt_exif_r is None:
        if w_e404:
            print("E404 Warning: Exif datetime not found")
            pause()
        return None

    dt_exif = strptime(dt_exif_r, "%Y:%m:%d %H:%M:%S")

    if w_edor and (dt_exif < w_wdor_low or dt_exif > w_wdor_high):
        print("EDOR Warning: Win datetime out of range")
        pause()

    return strftime(fmt_dt, dt_exif)


def get_dt_vid(path: str) -> str | None:
    """
    Video datetime reader

    Returns a video datetime in the specified format
    or None if the video does not have props

    Fix the path for Win API
    Try to get video props
    If the props are absent, warn user and return None
    Try to get a video datetime prop
    If the video datetime prop is absent, warn user and return None
    Convert it to struct_time
    Check the limits
    Return the datetime in printable format

    path - Path to the required file

    prop - Video props
    dt_prop_r - "Raw" prop datetime
    dt_prop - Prop datetime
    path_win - Path to the required file in the Win format
    """

    props: PyIPropertyStore | None = None  # type: ignore
    # I have no idea value of what type messy Win
    # PyPROPVARIANT.GetValue() method returns
    # This is a highly weird random-acting proprietary darkest dark hole
    dt_prop_r = None
    dt_prop: struct_time | None = None

    path_win: str = path.replace('/', '\\')

    props = propsys.SHGetPropertyStoreFromParsingName(path_win)  # type: ignore
    if props is None:
        if w_e404:
            print("E404 Warning: Props not found")
            pause()
        return None

    dt_prop_r = props.GetValue(pscon.PKEY_Media_DateEncoded).GetValue()

    if dt_prop_r is None:
        if w_e404:
            print("E404 Warning: Video datetime prop not found")
            pause()
        return None

    dt_prop = localtime(dt_prop_r.timestamp())

    if w_edor and (dt_prop < w_wdor_low or dt_prop > w_wdor_high):
        print("EDOR Warning: Win datetime out of range")
        pause()

    return strftime(fmt_dt, dt_prop)


def get_dt_win(path: str) -> str:
    """
    Win datetime reader

    Returns a Win edit datetime in the specified format

    Get Win edit datetime (float) for the specified path
    Convert it to struct_time with timezone offset
    Check the limits
    Return the datetime in printable format

    path - Path to the required file

    dt_win_f - Win datetime (float)
    dt_win - Win datetime
    """

    dt_win_f: float | None = None
    dt_win: struct_time | None = None

    dt_win_f = getmtime(path)

    dt_win = localtime(dt_win_f)

    if w_wdor and (dt_win < w_wdor_low or dt_win > w_wdor_high):
        print("WDOR Warning: Win datetime out of range")
        pause()

    return strftime(fmt_dt, dt_win)


def get_t_fn_pfx_ext_dt(src: str, tgt: str, fn_ext: str) -> str:
    """
    Mediafile name constructor (datetime)

    Returns target mediafile Filename + Postfix + Extension

    Get an extension
    Get Path + Filename + Extension
    If the mediafile is an image, get it's image datetime
    If the mediafile is a video, get it's video datetime
    If got an empty "true" dt, get Win dt
    Else get Win dt and compare them,
    put both into filename when large difference
    Get a new mediafile extension
    Get a postfix
    Return a new Filename + Extension

    Little bit complicated just for portability

    src - Source directory path
    tgt - Target directory path
    fn_ext - Filename + Extension

    ext - Extension
    p_fn_ext - Path + Filename + Extension
    t_ext - Target extension
    pfx - Postfix
    dt - Datetime
    dt_bak - Win datetime (used for dt comparison)
    """

    ext: str | None = None
    p_fn_ext: str | None = None
    t_ext: str | None = None
    pfx: str | None = None
    dt: str | None = None
    dt_dual: str | None = None

    ext = fn_ext.split('.')[-1].lower()
    p_fn_ext = join(src, fn_ext)

    if ext in ext_img:
        dt = get_dt_img(p_fn_ext)
    elif ext in ext_vid:
        dt = get_dt_vid(p_fn_ext)

    if dt is None:
        dt = get_dt_win(p_fn_ext)

    else:
        dt_dual = get_dt_win(p_fn_ext)

        if abs(mktime(strptime(dt, fmt_dt)) -
                mktime(strptime(dt_dual, fmt_dt))) > w_diff_eps:

            if w_diff:
                print("DIFF Warning: Datetimes differ too much:")
                print("[{}] [{}]".format(dt, dt_dual))
                pause()

            if dual:
                dt = "{} ({})".format(dt, dt_dual)

    t_ext = ext_sup[ext]

    pfx = get_pfx("{}.{}".format(dt, t_ext), tgt)

    return "{}{}.{}".format(dt, pfx, t_ext)


def rename_media(src: str, tgt: str) -> int:
    """
    Media renamer

    Walk through the src directory elements:
    Print the separator
    Get Path + Filename + Extension
    If an element is not a file or directory, skip it
    If the element is a directory and the recursion is allowed:
    Get a target path
    If the target path does not exist create it
    Recurse
    If deletion is allowed and the element became empty, delete it
    If the element is a file:
    Get an extension
    If the extension is not supported, skip the file
    If the extension is supported:
    Get target Filename + Postfix + Extension
    Construct a target path
    Rename the mediafile

    This complicated structure is prepared for the "rename by size" option

    src - Source directory path
    tgt - Target directory path

    ext - Extension
    fn_ext - Filename + Extension
    p_fn_ex - Path + Filename + Extension
    t_fn_pfx_ext - Target Filename + Postfix + Extension
    t_path - Target path
    """

    ext: str | None = None
    fn_ext: str | None = None
    p_fn_ext: str | None = None
    t_fn_pfx_ext: str | None = None
    t_path: str | None = None

    for fn_ext in listdir(src):
        splitout()
        p_fn_ext = join(src, fn_ext)

        if isfile(p_fn_ext):
            ext = fn_ext.split('.')[-1].lower()

            if ext in ext_sup:
                t_fn_pfx_ext = get_t_fn_pfx_ext_dt(src, tgt, fn_ext)
                t_path = join(tgt, t_fn_pfx_ext)
                print("\t\"{}\" -> \"{}\"".format(fn_ext, t_fn_pfx_ext))
                rename(p_fn_ext, t_path)

            else:
                print("Unsupported extension: \"{}\"".format(fn_ext))

        elif isdir(p_fn_ext):
            if recurse:
                print("Entering directory: \"{}\"".format(fn_ext))

                t_path = join(tgt, fn_ext)
                if not exists(t_path):
                    mkdir(t_path)

                rename_media(p_fn_ext, t_path)

                if del_empty and (not listdir(p_fn_ext)):
                    rmdir(p_fn_ext)

        else:
            print("Not a file or directory: \"{}\"".format(fn_ext))

    return 0


if __name__ == "__main__":
    """
    Main function

    Check if OS is supported
    Get source and target mediafiles paths
    Make sure the source path is specified
    If the target path is empty, attemt to use the source one
    Warn user if the source and target paths are same
    Rename the mediafiles
    If deletion is allowed and the source directory became empty, delete it

    src - Source mediafiles path
    tgt - Target mediafiles path
    """

    if osname != "nt":
        raise Exception("OS is not supported")

    src: str | None = None
    tgt: str | None = None

    src = get_path("source")
    if src is None:
        raise Exception("Source directory path is not specified")

    tgt = get_path("target")
    if tgt is None:
        tgt = src

    if w_stps and (src == tgt):
        print("STPS Warning: the source and target paths are same")
        pause()

    rename_media(src, tgt)

    if del_empty and (not listdir(src)):
        rmdir(src)

    exit(0)

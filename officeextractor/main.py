import zipfile
from zipfile import ZipFile
from pathlib import Path
from typing import Union, Optional, List, Tuple

from officeextractor.exceptions import FileTypeError, NotAValidZipFileError

SUPPORTED_FILETYPES = [
    "docx",  # Microsoft Word
    "docm",
    "dotx",
    "dotm",
    "xlsx",  # Microsoft Excel
    "xlsm",
    "xltx",
    "xltm",
    "xlsb",
    "pptx",  # Microsoft PowerPoint
    "pptm",
    "potm",
    "potx",
    "ppsx",
    "ppsm",
    "odt",  # LibreOffice Writer
    "ott",
    "ods",  # LibreOffice Calc
    "ots",
    "odp",  # LibreOffice Impress
    "otp",
    "odg",
]

OFFICE_2003_FILETYPES = ["doc", "dot", "ppt", "pot", "xls", "xlt"]


def extract(
    src: Union[str, List[str], Tuple[str]], dest: str, log: Optional[bool] = True
) -> None:
    """Extract the media files (images, audio, video) from src and save them to
    subfolders.

    Parameters
    ----------
    src : str or List[str, ...] or Tuple[str, ...]
        Either a single file (string) or a list or tuple of files
        (list/tuple of strings).
    dest : str
        Output directory as relative or full path.
    log : bool
        Optional. Whether logging should be actived or not. If True, print a summary of
        the extraction. Default is True.

    Returns
    -------
    None

    Examples
    --------
    >>> extract(src=("File1.docx", "File2.xlsx"), dest="Output_Folder", log=True)

    4 media files found in File1.docx:
    - 2 jpeg
    - 1 gif
    - 1 png

    No media files found in File2.xlsx.
    """

    dest_path = Path(dest)  # Convert dest to Path object

    # Convert src to a list, in case it's just a string
    if isinstance(src, str):
        src = [src]

    for file_name in src:
        check_valid_file(file_name)  # Check if the file is valid
        # Create subfolder Path object
        output_folder = Path.joinpath(dest_path, file_name.split("/")[-1])
        # Do the actual extraction
        with ZipFile(file_name, "r") as zip_file:
            media_list = get_media_list(zip_file=zip_file)
            file_type_count = extract_media(
                media_list=media_list, zip_file=zip_file, output_folder=output_folder
            )

        if log:  # Print short summary, if log==True
            amount_files = sum(i[1] for i in file_type_count)
            if amount_files == 0:
                print(f"\nNo media files found in {file_name}.")
            elif amount_files == 1:
                print(f"\n1 media file found in {file_name}:")
            else:
                print(f"\n{amount_files} media files found in {file_name}:")

            # Print amount of file types
            for i in file_type_count:
                print(f"- {i[1]} {i[0]}")


def check_valid_file(file_name: str) -> None:
    """Check if file_name is a valid file.

    It is valid, if it is not a Office 2003 file, if it is in the list of supported file
    types and if it can be unziped.

    Parameters
    ----------
    file_name : str

    Returns
    -------
    None
    """

    file_type = file_name.split(".")[-1]

    if file_type in OFFICE_2003_FILETYPES:
        raise FileTypeError(
            f"Invalid file: {file_name}. Office 2003 files are not supported."
        )

    if file_type not in SUPPORTED_FILETYPES:
        raise FileTypeError(
            f"Invalid file: {file_name}. File type .{file_type} is not supported."
        )

    if not zipfile.is_zipfile(file_name):
        raise NotAValidZipFileError(
            f"Seems like file {file_name} is not a valid zip file. "
            f"Maybe the file is corrupted."
        )


def create_folder(folder_name: Path) -> None:
    """Create subfolder from folder_name.

    Parameters
    ----------
    folder_name : Path

    Returns
    -------
    None
    """

    if not Path.exists(folder_name):
        Path.mkdir(folder_name, parents=True)


def get_media_list(zip_file: ZipFile) -> List[str]:
    """Get a list of paths to all the media files (images, audio, video) in the zip
    file.

    Parameters
    ----------
    zip_file : ZipFile
        Containing zip file

    Returns
    -------
    media_list : list[str]
        The list of media files as strings, representing their location
        (e.g. ["word/media/image1.jpeg", "word/media/image2.jpeg"])
    """

    # Only used for testing - to see the structure of the zipped file
    # zip_file.extractall(path="FULL_ZIP")

    # Add the path of all media files to media_list
    media_list = []
    for media in zip_file.namelist():
        if "." in media:
            file_type = media.split(".")[-1]

            # Microsoft Office stores media files in a folder called "media",
            # LibreOffice in a folder called "Pictures"; .emf files are added to
            # Microsoft Office documents, if media objects are embedded - these files
            # are skipped.
            if ("media" in media or "Pictures" in media) and file_type not in "emf":
                media_list.append(media)

    return media_list


def extract_media(
    media_list: List[str], zip_file: ZipFile, output_folder: Path
) -> List[Tuple[str, int]]:
    """Read the media files, save them to the output_folder and return a list containing
    a summary.

    Parameters
    ----------
    media_list : list[str]
        List of paths to the media files in the `zip_file`
    zip_file : ZipFile
        The zip file to extract the media from
    output_folder : Path
        The folder that needs to be created. (e.g.: "Output/File1.docx")

    Returns
    ------
    file_tjype_count : list[tuple[str, int]]
        A summary of how many media files of which type have been extracted. Sorted by
        frequency (e.g.: [png: 4, jpg: 3, gif: 1]). Empty list, if there were no media
        files in `zip_file`.
    """

    file_type_count: dict = {}  # Dict for the found media files
    if media_list:
        create_folder(output_folder)  # Create subfolder

        for media in media_list:
            media_file = media.split("/")[-1]  # Get file name
            file_type = media_file.split(".")[-1]  # Get file type

            # Add file_type to dict, if it doesn't exist yet and increase count by 1
            file_type_count[file_type] = file_type_count.get(file_type, 0) + 1

            # Read the data from the media file
            media_data = zip_file.read(name=media)

            # Write media data back to a file in the output_folder
            with open(Path.joinpath(output_folder, media_file), "wb") as media_fh:
                media_fh.write(media_data)

    # Return the file_type_count as list, sorted first by values (frequency) and then by
    # key (file extension)
    return sorted(file_type_count.items(), key=lambda i: (-i[1], i[0]))


# Todo:
#  - CI (Travis)
#  - coveralls.io
#  - Add badges (pre-commit, black etc.)
#  - Add README.md
#  - Add support for password encryption
#  - Sphinx & Read the Docs

import unittest
from unittest import TestCase
from unittest.mock import patch, mock_open
from pathlib import Path
import zipfile

from officeextractor.main import (
    check_valid_file,
    create_folder,
    get_media_list,
    extract_media,
    extract,
)

from officeextractor.exceptions import FileTypeError, NotAValidFileError


class TestCheckValidFile(TestCase):
    def test_Office_2003_file_type(self):
        with self.assertRaises(FileTypeError):
            check_valid_file(file_name=Path("my/folder/Test.doc"))

    def test_unsupported_file_type(self):
        with self.assertRaises(FileTypeError):
            check_valid_file(file_name=Path("my/folder/Test.abcdefg"))

    @patch("officeextractor.main.zipfile.is_zipfile", return_value=False)
    def test_corrupted_zip_file(self, mock_zip_file):
        with self.assertRaises(NotAValidFileError):
            check_valid_file(file_name=Path("my/folder/Corrupt_File.docx"))

    @patch("officeextractor.main.zipfile.is_zipfile", return_value=True)
    def test_working_file(self, mock_zip_file):
        self.assertIsNone(check_valid_file(file_name=Path("my/folder/Valid_File.docx")))


@patch("officeextractor.main.Path.exists")
@patch("officeextractor.main.Path.mkdir")
class TestCreateFolder(TestCase):
    def test_folder_already_exists(self, mock_mkdir, mock_exists):
        mock_exists.return_value = True
        create_folder(folder_name=Path("Just/some/random/path"))

        self.assertEqual(1, mock_exists.call_count)
        self.assertEqual(0, mock_mkdir.call_count)

    def test_folder_does_not_exist(self, mock_mkdir, mock_exists):
        mock_exists.return_value = False
        create_folder(folder_name=Path("Just/some/random/path"))

        self.assertEqual(1, mock_exists.call_count)
        self.assertEqual(1, mock_mkdir.call_count)


@patch.object(zipfile.ZipFile, "namelist")
class TestGetMediaList(TestCase):
    def test_get_media_list(self, mock_zip_file):
        # Mock the content of the zip file
        mock_zip_file.namelist.return_value = [
            "no_file_extension",
            "file.txt",
            "folder/file.txt",
            "folder/no_file_extension",
            "word/media/image1.jpeg",
            "word/media/image2.gif",
            "word/media/image5.emf",
            "Pictures/image3.jpeg",
            "Pictures/image4.jpeg",
        ]

        self.assertEqual(
            [
                "word/media/image1.jpeg",
                "word/media/image2.gif",
                "Pictures/image3.jpeg",
                "Pictures/image4.jpeg",
            ],
            get_media_list(zip_file=mock_zip_file),
        )

    def test_get_media_list_empty(self, mock_zip_file):
        # Mock the content of the zip file
        mock_zip_file.namelist.return_value = [
            "no_file_extension",
            "file.txt",
            "folder/file.txt",
            "folder/no_file_extension",
        ]

        self.assertEqual([], get_media_list(zip_file=mock_zip_file))


@patch.object(zipfile.ZipFile, "read")
@patch("officeextractor.main.create_folder")
@patch("officeextractor.main.open", side_effect=mock_open())
class TestExtractMedia(TestCase):
    def test_extract_media(self, mock_open_file, mock_create_folder, mock_zip_file):
        media_list = [
            "word/media/image1.jpeg",
            "word/media/image2.gif",
            "Pictures/image3.jpeg",
            "Pictures/image4.jpeg",
            "Pictures/video5.mp4",
            "Pictures/image6.png",
            "Pictures/image7.png",
        ]
        output_folder = Path("AAAA/Test.docx")

        # Mock bytes content of media file
        mock_zip_file.read.return_value = b"abcdefg"

        file_type_count = extract_media(
            media_list=media_list, zip_file=mock_zip_file, output_folder=output_folder
        )

        mock_open_filename = mock_open_file.call_args_list[0][0][0].parts
        mock_open_mode = mock_open_file.call_args_list[0][0][1]

        self.assertEqual(1, mock_create_folder.call_count)
        self.assertEqual(7, mock_open_file.call_count)
        self.assertEqual(b"abcdefg", mock_open_file().write.call_args[0][0])

        self.assertEqual(("AAAA", "Test.docx", "image1.jpeg"), mock_open_filename)
        self.assertEqual("wb", mock_open_mode)
        self.assertEqual(
            [("jpeg", 3), ("png", 2), ("gif", 1), ("mp4", 1)], file_type_count
        )

    def test_extract_media_empty(
        self, mock_open_file, mock_create_folder, mock_zip_file
    ):
        media_list = []
        output_folder = Path("AAAA/Test.docx")

        file_type_count = extract_media(
            media_list=media_list, zip_file=mock_zip_file, output_folder=output_folder
        )

        self.assertFalse(mock_create_folder.called)
        self.assertFalse(mock_zip_file.called)
        self.assertFalse(mock_open_file.called)
        self.assertEqual([], file_type_count)


@patch("officeextractor.main.print")
@patch("officeextractor.main.ZipFile")
@patch("officeextractor.main.extract_media")
@patch("officeextractor.main.get_media_list")
@patch("officeextractor.main.check_valid_file")
class TestExtract(TestCase):
    def test_extract_log_true(
        self,
        mock_check_valid_file,
        mock_get_media_list,
        mock_extract_media,
        mock_zip_file,
        mock_print,
    ):
        src = [
            "some/folder/Test.docx",
            "some/other/folder/Test.xlsx",
            "Test.pptx",  # without folder
            r"some\other\folder2\Test2.xlsx",  # backward slash
            "some\\other\\folder3\\Test3.xlsx",  # double backward slash
        ]
        dest = "Output_folder"

        # Mock return values of extract_media()
        mock_extract_media.side_effect = [
            [],
            [("jpg", 1)],
            [("jpg", 3), ("gif", 2), ("mp4", 1)],
            [("png", 1)],
            [("jpg", 2)],
        ]

        extract(src=src, dest=dest, log=True)

        self.assertEqual(5, mock_check_valid_file.call_count)
        self.assertEqual(5, mock_get_media_list.call_count)
        self.assertEqual(5, mock_extract_media.call_count)
        self.assertEqual(5, mock_zip_file.call_count)
        self.assertEqual(Path(src[-1]), mock_zip_file.call_args[0][0])
        self.assertEqual("r", mock_zip_file.call_args[0][1])

        # Test print() to sys.stdout
        expected = (
            f"\nNo media files found in {Path(src[0])}.",
            f"\n1 media file extracted from {Path(src[1])}:",
            "- 1 jpg",
            "\n6 media files extracted from Test.pptx:",
            "- 3 jpg",
            "- 2 gif",
            "- 1 mp4",
            f"\n1 media file extracted from {Path(src[3])}:",
            "- 1 png",
            f"\n2 media files extracted from {Path(src[4])}:",
            "- 2 jpg",
        )

        for expected_calls, actual_calls in zip(expected, mock_print.call_args_list):
            self.assertEqual(expected_calls, actual_calls[0][0])

    def test_extract_log_false_single_string(
        self,
        mock_check_valid_file,
        mock_get_media_list,
        mock_extract_media,
        mock_zip_file,
        mock_print,
    ):
        src = "my/folder/Test.docx"
        dest = "Output_folder"

        extract(src=src, dest=dest, log=False)

        self.assertFalse(mock_print.called)


if __name__ == "__main__":
    unittest.main()

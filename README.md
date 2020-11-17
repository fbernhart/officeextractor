# officeextractor

|     |     |
| --- | --- |
| Test Status | [![Build Status](https://img.shields.io/travis/com/fbernhart/officeextractor/main.svg?style=flat-square&label=TravisCI&logo=Travis&logoColor=white)](https://travis-ci.com/fbernhart/officeextractor) [![Coverage Status](https://img.shields.io/coveralls/fbernhart/officeextractor/main.svg?style=flat-square&label=coverage&logo=coveralls&logoColor=white)](https://coveralls.io/github/fbernhart/officeextractor?branch=main) |
| Version Info | [![PyPI Version](https://img.shields.io/pypi/v/officeextrator?style=flat-square&label=PyPI&logo=PyPI&logoColor=white&color=blue)](https://pypi.org/project/officeextractor) [![PyPI Downloads](https://img.shields.io/pypi/dm/officeextrator.svg?style=flat-square&label=Downloads&logo=PyPI&logoColor=white)](https://pypi.org/project/officeextractor) |
| Compatibility | [![Python Versions](https://img.shields.io/pypi/pyversions/officeextrator?style=flat-square&label=Python&logo=Python&logoColor=white&color=blue)](https://pypi.org/project/officeextractor) |
| Style | [![Code Style: Black](https://img.shields.io/badge/code%20style-black-000000?style=flat-square&.svg)](https://github.com/psf/black) [![pre-commit](https://img.shields.io/badge/pre--commit-enabled-brightgreen?logo=pre-commit&logoColor=white&style=flat-square)](https://github.com/pre-commit/pre-commit) |

## About

`officeextractor` is a Python library to extract media files like images, audio and video from office documents (Microsoft Office & LibreOffice).

## Supported File Types

Supported | File Types | Supported Media Formats
--- | --- | ---
Microsoft Word | docx, docm, dotm, dotx | images 
Microsoft Excel | xlsx, xlsb, xlsm, xltm, xltx | images 
Microsoft PowerPoint | potx, ppsm, ppsx, pptm, pptx, potm | images, video & audio
LibreOffice Writer | odt, ott | images 
LibreOffice Calc | ods, ots | images 
LibreOffice Impress | odp, otp, odg | images 

> :warning: **NOTE:** There is no support for Microsoft Office 2003 files (doc, xls, ppt etc.)


## Installation

```
pip install officeextractor
```

## Usage

```
>>> import officeextractor

>>> officeextractor.extract(src=("File1.docx", "Folder/File2.xlsx"), dest="Path/To/Output/Folder")

4 media files extracted from File1.docx:
- 2 jpeg
- 1 gif
- 1 png

1 media file extracted from Folder/File2.xlsx:
- 1 png
```

#### Parameters
> officeextractor.extract(src, dest, log=True)

> src : str or List[str, ...] or Tuple[str, ...]
> 
> Either a single file (string) or a list or tuple of files (list/tuple of strings).

> dest : str
> 
>   Output directory as relative or full path.

> log : bool
> 
> Optional. Whether logging should be actived or not. If True, print a summary of the extraction. Default is True.

## Licence

[GNU General Public License v3.0](https://github.com/fbernhart/officeextractor/blob/main/LICENSE)

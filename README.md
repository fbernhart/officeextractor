# officeextractor

<table>
<tr>
    <td>Test Status</td>
    <td><a href="https://travis-ci.com/fbernhart/officeextractor"><img src="https://img.shields.io/travis/com/fbernhart/officeextractor/main?style=flat-square&label=TravisCI&logo=Travis&logoColor=white" alt="Build Status"></a> <a href="https://coveralls.io/github/fbernhart/officeextractor?branch=main"><img src="https://img.shields.io/coveralls/fbernhart/officeextractor/main?style=flat-square&label=coverage&logo=coveralls&logoColor=white" alt="Coverage Status"></a></td>
</tr>
<tr>
    <td>Version Info</td>
    <td><a href="https://pypi.org/project/officeextractor"><img src="https://img.shields.io/pypi/v/officeextractor?style=flat-square&label=PyPI&logo=PyPI&logoColor=white&color=blue" alt="PyPI Version"></a> <a href="https://pypi.org/project/officeextractor"><img src="https://img.shields.io/pypi/dm/officeextractor?style=flat-square&label=Downloads&logo=PyPI&logoColor=white" alt="PyPI Downloads"></a></td>
</tr>
<tr>
    <td>Compatibility</td>
    <td><a href="https://pypi.org/project/officeextractor"><img src="https://img.shields.io/pypi/pyversions/officeextractor?style=flat-square&label=Python&logo=Python&logoColor=white&color=blue" alt="Python Versions"></a></td>
</tr>
<tr>
    <td>Style</td>
    <td><a href="https://github.com/psf/black"><img src="https://img.shields.io/badge/code%20style-black-000000?style=flat-square" alt="Code Style: Black"></a> <a href="https://github.com/pre-commit/pre-commit"><img src="https://img.shields.io/badge/pre--commit-enabled-brightgreen?logo=pre-commit&logoColor=white&style=flat-square" alt="pre-commit"></a></td>
</tr>
</table>

<br>

## About

**officeextractor** is a Python library to extract media files like images, audio and video from office documents (Microsoft Office & LibreOffice).

<br>

## Supported File Types

Supported | File Types | Supported Media Formats
--- | --- | ---
Microsoft Word | docx, docm, dotm, dotx | images 
Microsoft Excel | xlsx, xlsb, xlsm, xltm, xltx | images 
Microsoft PowerPoint | potx, ppsm, ppsx, pptm, pptx, potm | images, video & audio
LibreOffice Writer | odt, ott | images 
LibreOffice Calc | ods, ots | images 
LibreOffice Impress | odp, otp, odg | images 

##### &#9888; **NOTE:** Microsoft Office 2003 files (doc, dot, xls, xlt, ppt, pot) are not supported.

<br>

## Installation

```
pip install officeextractor
```

<br>

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

##### Parameters
> **officeextractor.extract(src, dest, log=True)**

> **src :** ***str, list of str or tuple of str***
> 
> Either a single file (string) or several files (list/tuple of strings) as relative or full path.
> 
> **dest :** ***str***
> 
> Output directory as relative or full path. If the directory doesn't exist, it will be created.
> 
> **log :** ***bool, optional***
> 
> Whether logging should be actived or not. If True, print a summary of the extraction. Default is True.

<br>

## Licence

[GNU General Public License v3.0](https://github.com/fbernhart/officeextractor/blob/main/LICENSE)

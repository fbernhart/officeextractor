import pathlib
import setuptools

# The directory containing this file
HERE = pathlib.Path(__file__).parent

# The text of the README file
README = (HERE / "README.md").read_text()

# Extract the version from __init__.py
with open("officeextractor/__init__.py") as fh:
    for line in fh.readlines():
        if line.startswith("__version__"):
            delim = '"' if '"' in line else "'"
            __version__ = line.split(delim)[1]

setuptools.setup(
    name="officeextractor",
    version=__version__,
    description="officeextractor extracts media files (images, videos, music) from "
    "Microsoft Office and LibreOffice files.",
    long_description=README,
    long_description_content_type="text/markdown",
    keywords="office extractor media images audio video extract docx pptx xlsx "
    "libreoffice microsoft",
    url="https://github.com/fbernhart/officeextractor",
    author="Florian Bernhart",
    author_email="florian-bernhart@hotmail.com",
    license="GNU General Public License v3.0",
    project_urls={
        "Bug Tracker": "https://github.com/fbernhart/officeextractor/issues",
        "Source Code": "https://github.com/fbernhart/officeextractor",
    },
    classifiers=[
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Operating System :: OS Independent",
        "Topic :: Office/Business",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Utilities",
    ],
    packages=["officeextractor"],
    python_requires=">=3.6",
    # install_requires=["some_dependency", "other_dependency"],
)

# PrintMerge
Using free `PDFCreator` to print to `PNG` files.
This script resize 4 PNG file and merge it to 1 PNG file and print it using python

## Supported OS
Windows

## Installation
  - Install `PDFCreator`, configure to print to `PNG` files, run `checkmerge.bat` after PNG created
  - Create folder `PRINT`
  - Create these folders under folder PRINT:
    - `PNG` (PDFCreator create PNG files to this folder)
    - `PNGMERGED` (Merged file stored here)
    - `PNGARCHIVE` (Merged source file archived to this folder)
  - Install Python 2.7.14
  - Install pypiwin32: `pip install pypiwin32`
  - Install pillow: `pip install pillow`

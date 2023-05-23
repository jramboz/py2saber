# py2saber
`py2saber` is a Python library and command-line utility for working with [OpenCore](https://github.com/LamaDiLuce/polaris-opencore)-based lightsabers. It is a re-implementation of [Ian "Nuntis" Jukes](http://sabers.amazer.uk/) excellent [sendtosaber](https://github.com/Nuntis-Spayz/Send-To-Saber) program, and extends on Nuntis' work in several ways:

- Built-in support for multiple files, including wildcard support (e.g., `*.RAW`)
- Detailed debugging output available
- Reusable Python functions for easy incorporation into other applications

## Installation
Use with system Python >=3.6:
- Clone repository
- Install requirements: `pip install -r requirements.txt`
- Run `python py2saber.py` to display usage information

Alternately, you can download pre-built binaries (coming soon).

## Usage
```
usage: py2saber [-h] [-v] [-i] [-l] [-s | -w] [-c] [-D] [--erase-all] [files ...]

A utility for working with OpenCore-based sabers, based on "sendtosaber" by Nuntis

positional arguments:
  files                 one or more files to upload to saber (separated by spaces)

options:
  -h, --help            show this help message and exit
  -v, --version         display version and author information, then exit
  -i, --info            read and display saber firmware version and serial number
  -l, --list            list all files on saber
  -s, --silent          exit without waiting for keypress (default)
  -w, --wait            wait for keypress before exiting
  -c, --continue-on-file-not-found
                        if one or more specified files do not exist, continue processing the remaining files (otherwise program will exit)
  -D, --debug           Show debugging information
  --erase-all           erase all files on saber
```

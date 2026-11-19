#!/bin/sh
# exit when any command fails
set -e

# pyinstaller packs the ready wheels (lxml included) in seconds; nuitka's mingw backend ran out of
# memory compiling lxml
pipenv run pyinstaller --onefile --noconfirm --name oxml-filler --paths . \
    --hidden-import _docfill --hidden-import _str_decoder --hidden-import _version $@ index.py

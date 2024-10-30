#!/bin/sh

pipenv run nuitka --follow-imports --onefile $@ index.py 
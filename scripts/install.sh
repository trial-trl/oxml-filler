#!/bin/sh

pipenv install $@
pipenv run pip freeze > requirements.txt
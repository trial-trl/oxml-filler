#!/bin/sh
# exit when any command fails
set -e

env=".env"

if echo "$1" | grep -q "dev\|test\|prod\|docker\|example\|ci"; then
    env=".env.$1"
    shift
elif [ "$1" = "default" ]; then
    shift
elif echo "$1" | grep -Eq ^/\|./; then
    env="$1"
    shift
fi

cmd="$1"

if [ ! -x "$(command -v $cmd)" ]; then
    PIPENV_DOTENV_LOCATION=$env pipenv run $@
    exit;
fi

PIPENV_DOTENV_LOCATION=$env $@

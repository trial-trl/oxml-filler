# OXML Filler

This aims to provide basic functionality for **manipulate** [`Office Open XML`](http://officeopenxml.com)-compatible files,
primarily to **fill** `Content Control` **forms**

## Requirements

- [python >= 3](https://www.python.org/)

### Kickstart

- Install dependencies with `pipenv install`
- Create `.env` file, based on `.env.example`

## Available commands

`pipenv run env [variant]...` - Execute command in specific environment

`pipenv run install [package]...` - Install package in this project

`pipenv run build` - Compile project into executable

## Support for multiple environment files

In this project, you can run as many environments as desired!

You just have to create a variation of `.env` files, such as `.env.dev` or `.env.prod`, being `dev` and `prod` the **environment codes**

You can run commands without bothering with environments at all; the default `.env`s will be picked automatically:

```SH
pipenv run [command]...
```

Or you can explicit request an environment:

```SH
pipenv run env dev [command]...
```

Some examples:

```SH
pipenv run env dev build
pipenv run env prod build
```

## Guidelines

### Development

This project follows [Conventional Commits](https://www.conventionalcommits.org/) specification for commit messages

Follow [bitbucket](https://support.atlassian.com/bitbucket-cloud/docs/resolve-issues-automatically-when-users-push-code/), [github](https://docs.github.com/en/issues/tracking-your-work-with-issues/linking-a-pull-request-to-an-issue) or similar guidelines to manage issues with commit messages

### Automation scripts

You can create scripts to automate actions, like command shorthands, inside `scripts` folder

To run any script, just call `sh scripts/path/to/script.sh`

### Shell scripts

You **must** create scripts compatible with `sh` and `bash` shells **only**

Make the code as portable as you can, in compliance with POSIX

## Windows support

If developing on Windows machine, you **must** configure `git` to use LF line endings at system level:

```SH
git config --system core.autocrlf false
git config --system core.eol lf
```

or global level:

```SH
git config --global core.autocrlf false
git config --global core.eol lf
```

In order to run **shell scripts** in local or CI/CD server, you must follow these steps:

1. Install [Git for Windows](https://git-scm.com/download/win)
2. Add `C:\YOUR_GIT_INSTALATION_DIR\bin` to system PATH

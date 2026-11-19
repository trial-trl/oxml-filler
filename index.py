import os
import argparse
import platform
import subprocess
import sys
import json
from time import sleep

from _version import __version__

EXIT_OK = 0
EXIT_USAGE = 1
EXIT_TEMPLATE = 2


class UsageError(Exception):
    exit_code = EXIT_USAGE


class ArgumentParser(argparse.ArgumentParser):
    """ argparse exits with 2 on bad usage, which here means a template error """

    def error(self, message):
        raise UsageError(message)


def create_arg_parser():
    arg_parser = ArgumentParser(prog="oxml-filler")
    arg_parser.add_argument(
        "--input", "-i", help="path to template document", type=str)
    arg_parser.add_argument(
        "--data", "-d", help="data to fill template document", type=str)
    arg_parser.add_argument(
        "--data-file", help="path to a JSON file with the data to fill template document (UTF-8)", type=str)
    arg_parser.add_argument("--field_mapping", "--field-mapping", "-m", dest="field_mapping",
                            help="a map of form fields to be replaced in template document x data property to receive", type=str)
    arg_parser.add_argument(
        "--field-mapping-file", help="path to a JSON file with the field mapping (UTF-8)", type=str)
    arg_parser.add_argument("--output", "-o", help="path to output file", type=str)
    arg_parser.add_argument("--open",  help="automatically open the output file; default is True",
                            type=bool, default=True, action=argparse.BooleanOptionalAction)
    arg_parser.add_argument("--temp", help="mark output file as temporary in disk; default is True (only works if 'open' is True)",
                            type=bool, default=True, action=argparse.BooleanOptionalAction)
    arg_parser.add_argument("--describe", help="list the content controls of the template instead of filling it",
                            action="store_true")
    arg_parser.add_argument("--version", help="print the version of the filler", action="store_true")
    arg_parser.add_argument("--json", help="print the result as JSON in the last line of the output",
                            action="store_true")
    arg_parser.add_argument("--quiet", help="don't print progress messages", action="store_true")
    return arg_parser


def configure_output():
    """ accents must survive the pipe to the caller, whatever the console encoding is """
    for stream in (sys.stdout, sys.stderr):
        if (hasattr(stream, "reconfigure")):
            stream.reconfigure(encoding="utf-8")


def read_json(value: str, file: str, name: str):
    from _docfill import DataError, is_json

    content = value

    if (file):
        try:
            with open(file, encoding="utf-8-sig") as json_file:
                content = json_file.read()
        except OSError:
            raise DataError(f"Não foi possível ler o arquivo de {name}: {file}")

    if (content is None):
        raise UsageError(f"Informe '{name}' (--{name} ou --{name}-file)")

    if (is_json(content) is False):
        raise DataError(f"'{name}' must be a valid JSON")

    return json.loads(content)


def delete_output(path: str, quiet: bool):
    try:
        if (os.path.isfile(path)):
            os.remove(path)
    except OSError as e:
        """ keep trying to delete the output file """
        if (getattr(e, "winerror", None) == 32):
            log('Output file is still in use. Waiting to close...', quiet)
        else:
            log(str(e), quiet)
        sleep(5)
        delete_output(path, quiet)


def log(message: str, quiet: bool):
    if (not quiet):
        print(message)


def open_output(output_path: str, temp: bool, quiet: bool):
    log("Auto-opening the output file...", quiet)

    current_os = platform.system()

    if (current_os == "Windows"):
        subprocess.check_output(f'start "" "{output_path}"', shell=True)
    else:
        subprocess.check_output(f'"{output_path}"', shell=True)

    log("Output file opened.", quiet)

    if (temp):
        log("Output file was marked as temporary, so it will self-destroy after closed. Waiting to close...", quiet)

        sleep(10)

        delete_output(output_path, quiet)

        log("Output file closed.", quiet)

        log("Output file was destroyed.", quiet)


def describe(args) -> dict:
    from _docfill import describe_fields

    if (not args.input):
        raise UsageError("Informe o template com --input")

    return {"ok": True, "fields": describe_fields(args.input)}


def fill(args, quiet: bool) -> dict:
    from _docfill import DataError, fill_template, parse_field_mapping

    if (not args.input):
        raise UsageError("Informe o template com --input")

    data = read_json(args.data, args.data_file, "data")

    if (not isinstance(data, dict)):
        raise DataError("'data' must be a JSON object")

    field_mapping = parse_field_mapping(
        read_json(args.field_mapping, args.field_mapping_file, "field_mapping"), data)

    output_path, filled = fill_template(
        args.input, data, field_mapping, args.output, args.temp and args.open)

    log(f"Output file created in '{output_path}'", quiet)

    if (args.open):
        open_output(output_path, args.temp, quiet)

    return {"ok": True, "fields": filled, "output": output_path}


def run(args) -> dict:
    from _docfill import FillerError, TemplateError

    if (args.version):
        return {"ok": True, "version": __version__}

    quiet = args.quiet

    try:
        if (args.describe):
            return describe(args)
        return fill(args, quiet)
    except (FillerError, UsageError):
        raise
    except Exception as error:
        """ anything else broke while reading or writing the document """
        raise TemplateError(f"Não foi possível processar o template: {error}")


def print_result(args, result: dict):
    if (args.json):
        print(json.dumps(result, ensure_ascii=False))
    elif (args.version):
        print(result["version"])
    elif (args.describe):
        print(json.dumps(result["fields"], ensure_ascii=False, indent=2))


def main(argv: list[str] = None) -> int:
    configure_output()

    argv = sys.argv[1:] if argv is None else argv
    as_json = "--json" in argv

    try:
        args = create_arg_parser().parse_args(argv)
        result = run(args)
    except Exception as error:
        exit_code = getattr(error, "exit_code", EXIT_USAGE)
        if (as_json):
            print(json.dumps({"ok": False, "error": str(error)}, ensure_ascii=False))
        else:
            print(str(error), file=sys.stderr)
        return exit_code

    print_result(args, result)

    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())

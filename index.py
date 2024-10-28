import os
import argparse
import platform
import subprocess
import sys
import json
from time import sleep

from _docfill import FieldOptions, create_document

template_input_path: str = None
data = {}
field_mapping: dict[str, FieldOptions] = {}

arg_parser = argparse.ArgumentParser()
arg_parser.add_argument(
    "--input", "-i", help="path to template document", type=str)
arg_parser.add_argument(
    "--data", "-d", help="data to fill template document", type=str)
arg_parser.add_argument("--field_mapping", "-m",
                        help="a map of form fields to be replaced in template document x data property to receive", type=str)
arg_parser.add_argument("--output", "-o", help="path to output file", type=str)
arg_parser.add_argument("--open",  help="automatically open the output file; default is True",
                        type=bool, default=True, action=argparse.BooleanOptionalAction)
arg_parser.add_argument("--temp", help="mark output file as temporary in disk; default is True (only works if 'open' is True)",
                        type=bool, default=True, action=argparse.BooleanOptionalAction)
args = arg_parser.parse_args()


def is_json(string: str, allow_simple: bool = False) -> bool:
    """
    Return True if the given string is in valid JSON format, else return False.

    Example usage
    -------------
    >>> is_json("1")       # False because this is a simple data structure (int)
    False

    >>> is_json("1", True) # return True because this is a valid JSON string
    True

    :param string:
    :param allow_simple: bool; If true allows simple data structure
    :return: bool
    """
    try:
        valid_json = json.loads(string)
        if allow_simple is False and type(valid_json) in [int, str, float, bool]:
            return False
    except ValueError:
        return False
    return True


def parse_data():
    global data
    if (is_json(args.data) is False):
        sys.exit("'data' must be a valid JSON")
    data = json.loads(args.data)


def parse_field_mapping():
    global field_mapping
    if (is_json(args.field_mapping) is False):
        sys.exit("'field_mapping' must be a valid JSON")
    field_mapping_raw = json.loads(args.field_mapping)
    for field in field_mapping_raw:
        options = FieldOptions.from_json(field_mapping_raw[field]) if (
            isinstance(field_mapping_raw[field], str) is False or is_json(field_mapping_raw[field])) else FieldOptions.only_map(field_mapping_raw[field])
        if (options.map_to not in data):
            sys.exit(
                f"'{field}' maps to '{options.map_to}' which doesn't exist in 'data'")
        field_mapping[field] = options


def delete_output(path: str):
    try:
        if (os.path.isfile(path)):
            os.remove(path)
    except OSError as e:
        """ keep trying to delete the output file """
        if (e.winerror == 32):
            print('Output file is still in use. Waiting to close...')
        else:
            print(e)
        sleep(5)
        delete_output(path)


parse_data()
parse_field_mapping()

template_input_path = args.input
output_path = args.output
open = args.open
temp = args.temp

output_path = create_document(
    template_input_path, data, field_mapping, output_path, temp and open)

print(f"Output file created in '{output_path}'")

if (open):
    print(f"Auto-opening the output file...")

    current_os = platform.system()

    if (current_os == "Windows"):
        subprocess.check_output(f'start "" "{output_path}"', shell=True)
    else:
        subprocess.check_output(f'"{output_path}"', shell=True)

    print(f"Output file opened.")

    if (temp):
        print(f"Output file was marked as temporary, so it will self-destroy after closed. Waiting to close...")

        sleep(10)

        delete_output(output_path)

        print(f"Output file closed.")

        print(f"Output file was destroyed.")

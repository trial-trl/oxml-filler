""" Entry of the embedded runtime (serious_python): answers the request file of `OXML_FILLER_REQUEST` in the
response file of `OXML_FILLER_RESPONSE`, which the app polls """

import json
import os

REQUEST_VARIABLE = "OXML_FILLER_REQUEST"
RESPONSE_VARIABLE = "OXML_FILLER_RESPONSE"


def read_request(path: str):
    with open(path, encoding="utf-8-sig") as request_file:
        return json.load(request_file)


def write_response(path: str, response: dict):
    """ the app reads the file as soon as it exists: write it aside, then rename it in one step """
    temporary_path = f"{path}.tmp"

    with open(temporary_path, "w", encoding="utf-8") as response_file:
        json.dump(response, response_file, ensure_ascii=False)

    os.replace(temporary_path, path)


def run():
    response_path = os.environ.get(RESPONSE_VARIABLE)

    if (not response_path):
        return None

    try:
        from app_entry import handle

        response = handle(read_request(os.environ.get(REQUEST_VARIABLE)))
    except Exception as error:
        response = {"ok": False, "error": f"Pedido inválido: {error}"}

    write_response(response_path, response)

    return response


run()

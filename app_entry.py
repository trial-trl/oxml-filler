""" Requests of the embedded runtime (serious_python), answered like the `--json` of the command line """


def version(payload: dict) -> dict:
    """ Answered without loading the engine (lxml, python-docx), so the app can probe the runtime """
    from _version import describe_version

    return describe_version()


def fill(payload: dict) -> dict:
    from _docfill import DataError, fill_template, parse_field_mapping

    data = payload.get("data") or {}

    if (not isinstance(data, dict)):
        raise DataError("'data' must be a JSON object")

    field_mapping = parse_field_mapping(payload.get("mapping") or {}, data)

    output_path, filled = fill_template(
        payload.get("template"), data, field_mapping, payload.get("output"), False)

    return {"ok": True, "fields": filled, "output": output_path}


def describe(payload: dict) -> dict:
    from _docfill import describe_fields

    return {"ok": True, "fields": describe_fields(payload.get("template"))}


handlers = {
    "version": version,
    "fill": fill,
    "describe": describe,
}


def handle(payload: dict) -> dict:
    """ Answer a request (`mode`: `version`, `fill` or `describe`); errors come back as `{"ok": false, "error": ...}` """
    try:
        if (not isinstance(payload, dict)):
            return {"ok": False, "error": "O pedido precisa ser um objeto JSON"}

        mode = payload.get("mode", "fill")

        if (mode not in handlers):
            return {"ok": False, "error": f"Pedido desconhecido: '{mode}'"}

        return handlers[mode](payload)
    except Exception as error:
        return {"ok": False, "error": str(error) or error.__class__.__name__}

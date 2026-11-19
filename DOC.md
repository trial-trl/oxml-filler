# Command Documentation

## Description

The built program processes a printing document. It receives input data in JSON format, maps the fields as specified, and generates the final document.

## Command Syntax

```sh
<oxml_filler_program> --input "<file_path>" --data '<json_data>' --field_mapping '<json_mapping>'
```

## Arguments

### `--input`

- **Type:** String
- **Description:** Specifies the path to the file to be processed.
- **Example:** `"/path/to/document.docm"`

### `--data`

- **Type:** JSON
- **Description:** Contains dynamic data to be inserted into the document.
- **Example:**

  ```json
  {
    "Nome do Cliente": "",
    "Endereço de Entrega": "",
    "Canal de Venda": "PoS",
    "Número do Pedido": "2312",
    "Itens": [["1x", "Coca-Cola 200ml", "1,30"]],
    "Total": 1.3,
    "Forma de Pagamento": "credit_online"
  }
  ```

### `--field_mapping`

- **Type:** JSON
- **Description:** Defines the mapping of input JSON fields to document fields.
- **Example:**
  ```json
  {
    "clientname": "Nome do Cliente",
    "delivery_address": {"mapTo": "Endereço de Entrega", "default": "Retirada"},
    "sale_channel": "Canal de Venda",
    "order_id": {"mapTo": "Número do Pedido", "prefix": "#"},
    "order_items": "Itens",
    "total": {"mapTo": "Total", "format": "currency"},
    "payment_method": "Forma de Pagamento"
  }
  ```

### `--data-file` and `--field-mapping-file`

- **Type:** String
- **Description:** Path to a UTF-8 JSON file with the data or the field mapping, in place of `--data` and
  `--field_mapping`. Use them for large payloads, which don't fit in the command line on Windows.
- **Example:** `--data-file "C:\temp\dados.json" --field-mapping-file "C:\temp\mapeamento.json"`

### `--output`, `--open` and `--temp`

- **Description:** `--output` sets the path of the filled document. By default the document is opened after
  filling (`--no-open` to skip) and deleted after closed (`--no-temp` to keep it).

### `--describe`

- **Description:** Lists the content controls of `--input` instead of filling it: one entry per tag, in document
  order, with its `kind` (`text`, `dropdown` or `repeating`), the labels of a dropdown (`items`) and the cells of
  the first row of a repeating section (`columns`, with the placeholder text of each one).
- **Example:**

  ```json
  [
    {"tag": "clientname", "kind": "text", "items": [], "columns": []},
    {"tag": "payment_method", "kind": "dropdown", "items": ["Dinheiro", "Crédito via internet"], "columns": []},
    {"tag": "order_items", "kind": "repeating", "items": [], "columns": [
      {"index": 0, "placeholder": "0x"}, {"index": 1, "placeholder": "ITEM"}, {"index": 2, "placeholder": "0,00"}
    ]}
  ]
  ```

### `--version`

- **Description:** Prints the version of the filler. `--input` is accepted and ignored.

### `--json` and `--quiet`

- **Description:** `--json` prints the result as JSON in the **last line** of the output; `--quiet` hides the
  progress messages, so both together print only the JSON. Results:

  | Command | JSON |
  | --- | --- |
  | Fill | `{"ok": true, "fields": 15, "output": "<path>"}` (`fields`: how many fields were filled) |
  | `--describe` | `{"ok": true, "fields": [...]}` |
  | `--version` | `{"ok": true, "version": "1.1.0"}` |
  | Error | `{"ok": false, "error": "<readable message>"}` |

## Exit codes

| Code | Meaning |
| --- | --- |
| `0` | Success |
| `1` | Invalid usage: unknown option, missing argument, invalid JSON or a field mapped to missing data |
| `2` | Template missing, unreadable or not a Word document |

The output (stdout and stderr) is always UTF-8, whatever the console encoding is.

## Content controls

| Shape in the document | Filled with |
| --- | --- |
| `w:sdtContent` with `w:r` | the first run, keeping its formatting |
| `w:sdtContent` with `w:p` | the first run of the first paragraph (a run is created in an empty paragraph) |
| `w:dropDownList` / `w:comboBox` | the label (`w:displayText`) of the item whose value is given; unknown values as is |
| `w15:repeatingSection` | one item per row of the list, each cell with a column |
| Composite (e.g. `w:tbl` plus `w:p` without runs) | the first run found deep inside the control |

Tags in the mapping that don't exist in the template are skipped.

## How to Execute

1. Ensure `<oxml_filler_program>` is installed and accessible.
2. Navigate to the directory where the document is located.
3. Run the command as per the syntax provided, replacing values as needed.

**Example Execution:**

```sh
<oxml_filler_program> --input "/path/to/document.docm" \
  --data '{"Nome do Cliente": "", "Endereço de Entrega": "", "Canal de Venda": "PoS", "Número do Pedido": "2312", "Itens": [["1x", "Coca-Cola 200ml", "1,30"]], "Total": 1.3, "Forma de Pagamento": "credit_online"}' \
  --field_mapping '{"clientname": "Nome do Cliente", "delivery_address": {"mapTo": "Endereço de Entrega", "default": "Retirada"}, "sale_channel": "Canal de Venda", "order_id": {"mapTo": "Número do Pedido", "prefix": "#"}, "order_items": "Itens", "total": {"mapTo": "Total", "format": "currency"}, "payment_method": "Forma de Pagamento"}'
```

## Notes

- Ensure file paths are correct and use backslashes (`\`) on Windows or regular slashes (`/`) on Linux/Mac.
- JSON must be properly formatted and without line breaks within the command line.
- Use `\` to split long commands into multiple lines in the terminal.
- From another program, prefer `--data-file`, `--field-mapping-file`, `--quiet` and `--json`:

  ```sh
  <oxml_filler_program> --input "/path/to/document.docm" --data-file dados.json \
    --field-mapping-file mapeamento.json --output "/path/to/cupom.docm" --no-open --no-temp --quiet --json
  ```

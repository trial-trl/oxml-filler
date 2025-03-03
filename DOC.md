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

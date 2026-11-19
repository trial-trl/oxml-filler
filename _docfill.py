import os
import json
import zipfile
import tempfile
import time
import docx
from typing import SupportsIndex
from copy import deepcopy
from lxml.etree import ElementBase
from docx.document import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml.parser import OxmlElement
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R
from docx.oxml.table import CT_Row, CT_Tc
from docx.oxml.ns import nsmap, qn

from _str_decoder import decode_escapes

nsmap['w15'] = "http://schemas.microsoft.com/office/word/2012/wordml"


class FillerError(Exception):
    """ Base error of the filler; `exit_code` is the code the command line exits with """
    exit_code = 1


class DataError(FillerError):
    """ Invalid usage: data or field mapping that can't be used """
    exit_code = 1


class TemplateError(FillerError):
    """ Template missing, unreadable or not a Word document """
    exit_code = 2


class StringPadOptions:
    width: SupportsIndex
    fillchar: str

    def __init__(self, declaration: str):
        width_fillchar = str(declaration).split(",")
        self.width = int(width_fillchar[0])
        try:
            self.fillchar = str(width_fillchar[1])
        except IndexError:
            self.fillchar = " "

class FieldOptions:
    map_to: str
    format: str | None = None
    prefix: str | None = None
    suffix: str | None = None
    default: str | None = None
    padLeft: StringPadOptions | None = None
    padRight: StringPadOptions | None = None
    skipIfEmpty: bool | None = None

    def __init__(self, map_to: str):
        self.map_to = map_to

    @staticmethod
    def from_json(options):
        if (isinstance(options, str)):
            options = json.loads(options)
        fieldOptions = FieldOptions(options['mapTo'])
        if ('format' in options):
            fieldOptions.format = options['format']
        if ('prefix' in options):
            fieldOptions.prefix = options['prefix']
        if ('suffix' in options):
            fieldOptions.suffix = options['suffix']
        if ('skipIfEmpty' in options):
            fieldOptions.skipIfEmpty = options['skipIfEmpty']
        if ('default' in options):
            fieldOptions.default = options['default']
        if ('padLeft' in options):
            fieldOptions.padLeft = StringPadOptions(options['padLeft'])
        if ('padRight' in options):
            fieldOptions.padRight = StringPadOptions(options['padRight'])
        return fieldOptions

    @staticmethod
    def only_map(map_to: str):
        return FieldOptions(map_to)


def format_currency(value, fallback=None) -> str:
    """
    Money in pt-BR (`1.234,50`) without `locale.setlocale`, which only works where the locale is installed
    (it broke on Linux and would break on mobile); same output as the pt_BR locale on Windows
    """
    for candidate in (value, fallback):
        try:
            number = float(candidate)
        except (TypeError, ValueError):
            continue
        return f'{number:,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')
    return str(fallback if fallback is not None else value)


class PropValueParser:
    options: FieldOptions

    def __init__(self, options: FieldOptions):
        self.options = options

    def parse(self, value):
        _val = value

        if ((value is None or value == "" or value == 0) and self.options.default):
            _val = str(self.options.default)

        if (self.options.format == 'currency'):
            _val = format_currency(value, _val)

        if (self.options.padLeft):
            _val = str(_val).rjust(
                self.options.padLeft.width,
                self.options.padLeft.fillchar
            )

        if (self.options.padRight):
            _val = str(_val).ljust(
                self.options.padRight.width,
                self.options.padRight.fillchar
            )

        if (self.options.prefix):
            _val = str(self.options.prefix) + str(_val)

        if (self.options.suffix):
            _val = str(_val) + str(self.options.suffix)

        return decode_escapes(str(_val))

    def canFill(self, value):
        return not self.options.skipIfEmpty or not (value is None or value == "" or value == 0)


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
    except (TypeError, ValueError):
        return False
    return True


def parse_field_mapping(field_mapping_raw: dict, data: dict) -> dict[str, FieldOptions]:
    """ Turn the raw mapping (`field: "prop"` or `field: {"mapTo": "prop", ...}`) into field options """
    if (not isinstance(field_mapping_raw, dict)):
        raise DataError("O mapeamento de campos precisa ser um objeto JSON")

    field_mapping: dict[str, FieldOptions] = {}

    for field in field_mapping_raw:
        raw = field_mapping_raw[field]
        try:
            options = FieldOptions.from_json(raw) if (
                isinstance(raw, str) is False or is_json(raw)) else FieldOptions.only_map(raw)
        except (KeyError, TypeError, ValueError):
            raise DataError(f"O mapeamento do campo '{field}' é inválido")
        if (options.map_to not in data):
            raise DataError(
                f"O campo '{field}' aponta para '{options.map_to}', que não existe nos dados")
        field_mapping[field] = options

    return field_mapping


def open_template(template_path: str) -> Document:
    if (not template_path):
        raise TemplateError("Informe o caminho do template")
    if (not os.path.isfile(template_path)):
        raise TemplateError(f"Template não encontrado: {template_path}")
    try:
        return docx.Document(template_path)
    except (PackageNotFoundError, zipfile.BadZipFile, KeyError, ValueError, OSError) as error:
        raise TemplateError(f"Template inválido ou ilegível: {template_path} ({error})")


def find_dropdown_pr(element):
    dropdown = element.xpath('./w:dropDownList | ./w:comboBox', namespaces=nsmap)
    return dropdown[0] if len(dropdown) else None


def find_text_pr(element):
    text = element.xpath('./w:text', namespaces=nsmap)
    return text[0] if len(text) else None


def find_repeating_section_pr(element):
    repeatingSection = element.xpath(
        './w15:repeatingSection', namespaces=nsmap)
    return repeatingSection[0] if len(repeatingSection) else None


def find_wt(element, direct=True):
    text = element.xpath('./w:t' if direct else './/w:t', namespaces=nsmap)
    return text[0] if len(text) else None


def find_sdt_content(sdt: ElementBase):
    return sdt.find(qn('w:sdtContent')) if sdt is not None else None


def map_dropdown(dropdown_element: list[ElementBase]):
    possible_values = {}
    for listItem in dropdown_element.findall(qn('w:listItem')):
        value = listItem.get(qn('w:value'))
        possible_values[value] = listItem.get(qn('w:displayText')) or value
    return possible_values


def get_first_deep_run(element):
    """ First `w:r` anywhere below the element (inside tables and nested controls too), or None """
    runs = element.xpath('.//w:r', namespaces=nsmap)
    return runs[0] if len(runs) else None


def get_first_run_and_remove_others(element):
    all_runs: list[CT_R] = element.findall(qn('w:r'))

    if (len(all_runs) == 0):
        """ no direct runs: use the first deep one, or create one in an empty paragraph """
        run = get_first_deep_run(element)
        if (run is None):
            run = OxmlElement('w:r')
            element.append(run)
        return run

    run: CT_R = all_runs[0]

    remove_other_than_first(all_runs, element)

    return run


def get_first_paragraph_run_and_remove_others(element):
    p: CT_P = element.find(qn('w:p'))

    if (p is None):
        return get_first_deep_run(element)

    return get_first_run_and_remove_others(p)


def remove_other_than_first(children, element):
    i = 0
    for child in children:
        if (i > 0):
            element.remove(child)
        i += 1


def map_row_to_std(row, std: ElementBase, value_parser: PropValueParser):
    is_list = isinstance(row, list)

    sdtContent: ElementBase = find_sdt_content(std)

    if (sdtContent is None or len(sdtContent) == 0):
        return

    target: ElementBase = None

    if (len(sdtContent) > 1):
        target = list(sdtContent)
    else:
        target = sdtContent[0]

    if (isinstance(target, CT_Row)):
        cell: CT_Tc
        for index, cell in enumerate(target.findall(qn('w:tc'))):
            run = get_first_paragraph_run_and_remove_others(cell)

            if (is_list and run is not None and index < len(row)):
                run.text = value_parser.parse(row[index])


def is_inside_repeating_section(sdt: ElementBase) -> bool:
    return len(sdt.xpath(
        'ancestor::w:sdt[w:sdtPr/w15:repeatingSection or w:sdtPr/w15:repeatingSectionItem]',
        namespaces=nsmap)) > 0


def get_text(element: ElementBase) -> str:
    return ''.join(element.xpath('.//w:t/text()', namespaces=nsmap))


def describe_repeating_columns(sdtContent: ElementBase) -> list[dict]:
    """ Cells of the first row of the section (not the layers of the `w:sdt` around it) """
    if (sdtContent is None):
        return []

    rows = sdtContent.xpath('.//w:tr', namespaces=nsmap)

    if (len(rows) == 0):
        return []

    return [{"index": index, "placeholder": get_text(cell)}
            for index, cell in enumerate(rows[0].findall(qn('w:tc')))]


def describe_document(document: Document) -> list[dict]:
    fields = []
    seen = set()

    for tag in document.element.xpath('//w:sdt/w:sdtPr/w:tag'):
        name = tag.get(qn('w:val'))
        if (not name or name in seen):
            continue

        form_field_pr: ElementBase = tag.getparent()
        sdt: ElementBase = form_field_pr.getparent()

        if (is_inside_repeating_section(sdt)):
            continue

        seen.add(name)

        dropdown_pr = find_dropdown_pr(form_field_pr)
        repeating_section_pr = find_repeating_section_pr(form_field_pr)

        field = {"tag": name, "kind": "text", "items": [], "columns": []}

        if (repeating_section_pr is not None):
            field["kind"] = "repeating"
            field["columns"] = describe_repeating_columns(find_sdt_content(sdt))
        elif (dropdown_pr is not None):
            field["kind"] = "dropdown"
            field["items"] = list(map_dropdown(dropdown_pr).values())

        fields.append(field)

    return fields


def describe_fields(template_path: str) -> list[dict]:
    """
    Describe the content controls of the template: one entry per tag, in document order

    Each entry has `tag`, `kind` (`text`, `dropdown` or `repeating`), `items` (labels of a dropdown) and
    `columns` (cells of the first row of a repeating section, with the placeholder text of each one)
    """
    return describe_document(open_template(template_path))


def fill_repeating_section(form_field_sdtContent: ElementBase, prop_value, value_parser: PropValueParser):
    if (len(form_field_sdtContent) == 0):
        return

    base_section = list(form_field_sdtContent)[0]
    clone_base_section = deepcopy(base_section)
    form_field_sdtContent.clear()

    for row in (prop_value or []):
        clone_section = deepcopy(clone_base_section)
        """ Word gives new items new ids: repeated ids break the document """
        for id in clone_section.xpath('.//w:sdtPr/w:id', namespaces=nsmap):
            id.getparent().remove(id)
        map_row_to_std(row, clone_section, value_parser)
        form_field_sdtContent.append(clone_section)


def fill_dropdown(form_field_sdtContent: ElementBase, dropdown_pr: ElementBase, prop_value):
    """ The dropdown shows the human label of the item; an unknown value is shown as is """
    possible_values = map_dropdown(dropdown_pr)
    text = possible_values.get(str(prop_value), str(prop_value))
    dropdown_t = find_wt(form_field_sdtContent, False)

    if (dropdown_t is not None):
        dropdown_t.text = text
        return

    run = get_first_deep_run(form_field_sdtContent)

    if (run is not None):
        run.text = text


def fill_text(field: str, form_field_sdtContent: ElementBase, prop_value, value_parser: PropValueParser):
    run: CT_R = None
    holder = None

    for child in form_field_sdtContent:
        if (isinstance(child, CT_R)):
            run = child
            holder = child
            break
        if (isinstance(child, CT_P) and len(child.xpath('.//w:r', namespaces=nsmap))):
            run = get_first_run_and_remove_others(child)
            holder = child
            break

    if (run is None):
        """ composite controls (a table plus a paragraph without runs): fill the first deep run """
        run = get_first_deep_run(form_field_sdtContent)

        if (run is None):
            paragraph = form_field_sdtContent.find(qn('w:p'))
            if (paragraph is None):
                raise TemplateError(f"O controle '{field}' não tem parágrafo nem texto para preencher")
            run = get_first_run_and_remove_others(paragraph)

        run.text = value_parser.parse(prop_value)
        return

    remove_other_than_first(
        [holder] + [child for child in form_field_sdtContent if child is not holder], form_field_sdtContent)

    run.text = value_parser.parse(prop_value)


def fill_document(document: Document, data={}, field_mapping: dict[str, FieldOptions] = {}) -> int:
    """ Fill the content controls of the document; return how many fields were filled """
    filled = 0

    for field in field_mapping:
        field_options = field_mapping[field]
        value_parser = PropValueParser(field_options)
        prop_value = data.get(field_options.map_to)

        if not value_parser.canFill(prop_value):
            continue

        form_field_tags = document.element.xpath(f'//w:tag[@w:val="{field}"]')

        if (len(form_field_tags) == 0):
            continue

        form_field_pr: ElementBase = form_field_tags[0].getparent()
        form_field_sdtContent: ElementBase = find_sdt_content(form_field_pr.getparent())

        if (form_field_sdtContent is None):
            continue

        dropdown_pr = find_dropdown_pr(form_field_pr)
        repeating_section_pr = find_repeating_section_pr(form_field_pr)

        if (repeating_section_pr is not None):
            fill_repeating_section(form_field_sdtContent, prop_value, value_parser)
        elif (dropdown_pr is not None):
            fill_dropdown(form_field_sdtContent, dropdown_pr, prop_value)
        else:
            """ Unidentified content controls are handled as text by default """
            fill_text(field, form_field_sdtContent, prop_value, value_parser)

        filled += 1

    return filled


def resolve_output_path(template_path: str, output_path: str = None, temp=True):
    file_suffix = f'{str(time.time()).replace(".", "_")}{os.path.splitext(template_path)[1]}'

    if (output_path is None and temp):
        output_path = os.path.join(tempfile.mkdtemp(), f'tmp_{file_suffix}')

    if (output_path is None):
        output_path = os.path.join(os.path.dirname(
            template_path), f'output_{file_suffix}')

    return output_path


def fill_template(template_path: str, data={}, field_mapping: dict[str, FieldOptions] = {}, output_path: str = None, temp=True) -> tuple[str, int]:
    """ Fill the template and save it; return the output path and how many fields were filled """
    document = open_template(template_path)

    filled = fill_document(document, data, field_mapping)

    output_path = resolve_output_path(template_path, output_path, temp)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    document.save(output_path)

    return output_path, filled


def create_document(template_path: str, data={}, field_mapping: dict[str, FieldOptions] = {}, output_path: str = None, temp=True):
    output_path, _ = fill_template(template_path, data, field_mapping, output_path, temp)

    return output_path

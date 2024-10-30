import os
import tempfile
import time
import docx
import locale
from copy import deepcopy
from lxml.etree import ElementBase
from docx.document import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R
from docx.oxml.table import CT_Row, CT_Tc
from docx.oxml.ns import nsmap, qn

from _str_decoder import decode_escapes

nsmap['w15'] = "http://schemas.microsoft.com/office/word/2012/wordml"


class FieldOptions:
    map_to: str
    format: str | None = None
    prefix: str | None = None
    suffix: str | None = None

    def __init__(self, map_to: str):
        self.map_to = map_to

    @staticmethod
    def from_json(json):
        if (isinstance(json, str)):
            json = json.loads(json)
        fieldOptions = FieldOptions(json['mapTo'])
        if ('format' in json):
            fieldOptions.format = json['format']
        if ('prefix' in json):
            fieldOptions.prefix = json['prefix']
        if ('suffix' in json):
            fieldOptions.prefix = json['suffix']
        return fieldOptions

    @staticmethod
    def only_map(map_to: str):
        return FieldOptions(map_to)


class PropValueParser:
    options: FieldOptions

    def __init__(self, options: FieldOptions):
        self.options = options

    def parse(self, value):
        _val = value

        if (self.options.format == 'currency'):
            locale.setlocale(locale.LC_ALL, 'pt_BR')
            _val = locale.format_string('%.2f', value, True)

        if (self.options.prefix):
            _val = str(self.options.prefix) + str(_val)

        if (self.options.suffix):
            _val = str(_val) + str(self.options.suffix)

        return decode_escapes(str(_val))


def find_dropdown_pr(element):
    dropdown = element.xpath('./w:dropDownList', namespaces=nsmap)
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


def map_dropdown(dropdown_element: list[ElementBase]):
    possible_values = {}
    for listItem in dropdown_element:
        possible_values[listItem.get(qn('w:value'))] = listItem.get(
            qn('w:displayText'))
    return possible_values


def map_row_to_std(row, std: ElementBase, value_parser: PropValueParser):
    is_list = isinstance(row, list)

    pr: ElementBase = list(std)[0]
    sdtContent: ElementBase = list(std)[1]

    target: ElementBase = None

    if (len(sdtContent) > 1):
        target = list(sdtContent)
    else:
        target = sdtContent[0]

    if (isinstance(target, CT_Row)):
        cell: CT_Tc
        for index, cell in enumerate(list(target)):
            run = get_first_paragraph_run_and_remove_others(cell)

            if (is_list):
                run.text = value_parser.parse(row[index])


def get_first_run_and_remove_others(element):
    all_runs: CT_R = element.findall(qn('w:r'))
    run: CT_R = all_runs[0]

    remove_other_than_first(all_runs, element)

    return run


def get_first_paragraph_run_and_remove_others(element):
    p: CT_P = element.find(qn('w:p'))

    return get_first_run_and_remove_others(p)


def remove_other_than_first(children, element):
    i = 0
    for child in children:
        if (i > 0):
            element.remove(child)
        i += 1


def create_document(template_path: str, data={}, field_mapping: dict[str, FieldOptions] = {}, output_path: str = None, temp=True):
    def get_field_options(field: str):
        return field_mapping[field]

    def get_prop_value_for_field(field: str):
        prop = field_mapping[field].map_to
        return data[prop]

    document: Document = docx.Document(template_path)

    for field in field_mapping:
        field_options = get_field_options(field)
        value_parser = PropValueParser(field_options)

        form_field_tag = document.element.xpath(
            f'//w:tag[@w:val="{field}"]')[0]
        form_field_pr: ElementBase = form_field_tag.getparent()
        form_field_sdt: ElementBase = form_field_pr.getparent()
        form_field_sdtContent: ElementBase = form_field_pr.getnext()

        dropdown_pr = find_dropdown_pr(form_field_pr)
        text_pr = find_text_pr(form_field_pr)
        repeating_section_pr = find_repeating_section_pr(form_field_pr)

        is_text = text_pr is not None
        is_dropdown = dropdown_pr is not None
        is_repeating_section = repeating_section_pr is not None

        if (is_repeating_section):
            prop_value: list = get_prop_value_for_field(field)
            base_section = list(form_field_sdtContent)[0]
            clone_base_section = deepcopy(base_section)
            form_field_sdtContent.clear()
            for row in prop_value:
                clone_section = deepcopy(clone_base_section)
                id = clone_section.xpath('//w:sdtPr/w:id', namespaces=nsmap)[0]
                id.getparent().remove(id)
                map_row_to_std(row, clone_section, value_parser)
                form_field_sdtContent.append(clone_section)

            continue

        if (is_text):
            prop_value = get_prop_value_for_field(field)
            run: CT_R = None

            for child in form_field_sdtContent:
                if (isinstance(child, CT_R)):
                    run = child
                    break
                if (isinstance(child, CT_P)):
                    run = get_first_run_and_remove_others(child)
                    break

            remove_other_than_first(
                list(form_field_sdtContent), form_field_sdtContent)

            if (run is None):
                str.exit(f"'{field}' not have 'w:p' or 'w:r' in 'sdtContent'")

            run.text = value_parser.parse(prop_value)
            continue

        if (is_dropdown):
            prop_value = str(get_prop_value_for_field(field))
            possible_values = map_dropdown(dropdown_pr)
            dropdown_t = find_wt(form_field_sdtContent, False)
            dropdown_t.text = possible_values[prop_value]
            continue

    file_suffix = f'{str(time.time()).replace(".", "_")}{os.path.splitext(template_path)[1]}'

    if (output_path is None and temp):
        output_path = os.path.join(tempfile.mkdtemp(), f'tmp_{file_suffix}')

    if (output_path is None):
        output_path = os.path.join(os.path.dirname(
            template_path), f'output_{file_suffix}')

    document.save(output_path)

    return output_path

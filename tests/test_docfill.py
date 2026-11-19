import pytest
import docx

from _docfill import (DataError, TemplateError, describe_fields, fill_template,
                      parse_field_mapping)
from conftest import DATA, MAPPING, output_path, tag_text
from template import TAGS


def fill(template, tmp_path, data=None, mapping=None):
    data = DATA if data is None else data
    field_mapping = parse_field_mapping(MAPPING if mapping is None else mapping, data)
    return fill_template(template, data, field_mapping, output_path(tmp_path), False)


class TestDescribe:
    def test_lists_every_tag_in_document_order(self, template):
        fields = describe_fields(template)

        assert [field['tag'] for field in fields] == TAGS

    def test_skips_controls_inside_the_repeating_section(self, template):
        tags = [field['tag'] for field in describe_fields(template)]

        assert 'item_price' not in tags

    def test_gives_the_kind_of_each_control(self, template):
        kinds = {field['tag']: field['kind'] for field in describe_fields(template)}

        assert kinds['payment_method'] == 'dropdown'
        assert kinds['order_items'] == 'repeating'
        assert kinds['clientname'] == 'text'
        assert kinds['return_change'] == 'text'

    def test_gives_the_labels_of_the_dropdown(self, template):
        dropdown = next(f for f in describe_fields(template) if f['tag'] == 'payment_method')

        assert dropdown['items'] == ['Escolha um item.', 'Crédito via internet', 'Dinheiro']

    def test_counts_the_cells_of_the_repeating_row(self, template):
        section = next(f for f in describe_fields(template) if f['tag'] == 'order_items')

        assert section['columns'] == [
            {'index': 0, 'placeholder': '0x'},
            {'index': 1, 'placeholder': 'ITEM'},
            {'index': 2, 'placeholder': '0,00'},
        ]

    def test_fails_with_template_error_for_a_missing_file(self, tmp_path):
        with pytest.raises(TemplateError):
            describe_fields(str(tmp_path / 'nada.docx'))

    def test_fails_with_template_error_for_a_file_that_is_not_word(self, tmp_path):
        path = tmp_path / 'texto.docx'
        path.write_text('não é um documento', encoding='utf-8')

        with pytest.raises(TemplateError):
            describe_fields(str(path))


class TestFill:
    def test_fills_every_mapped_field(self, template, tmp_path):
        _, filled = fill(template, tmp_path)

        assert filled == 6

    def test_keeps_a_single_run_in_text_controls(self, template, tmp_path):
        output, _ = fill(template, tmp_path)

        assert tag_text(output, 'clientname') == 'João da Silva'
        assert tag_text(output, 'total') == 'R$ 95'

    def test_shows_the_label_of_the_dropdown_item(self, template, tmp_path):
        output, _ = fill(template, tmp_path)

        assert tag_text(output, 'payment_method') == 'Crédito via internet'

    def test_shows_an_unknown_dropdown_value_as_is(self, template, tmp_path):
        data = {**DATA, 'Forma de Pagamento': 'voucher'}
        output, _ = fill(template, tmp_path, data)

        assert tag_text(output, 'payment_method') == 'voucher'

    @pytest.mark.parametrize('count', [0, 1, 3, 40])
    def test_repeats_the_section_for_each_item(self, template, tmp_path, count):
        items = [[f'{n}x', f'Item {n}', f'{n},00'] for n in range(1, count + 1)]
        output, _ = fill(template, tmp_path, {**DATA, 'Itens': items})

        document = docx.Document(output)
        rows = document.element.body.xpath('//w:sdt[w:sdtPr/w:tag/@w:val="order_items"]//w:tr')
        assert len(rows) == count
        assert [''.join(row.xpath('.//w:t/text()')) for row in rows] == [''.join(item) for item in items]

    def test_drops_the_ids_of_the_repeated_items(self, template, tmp_path):
        output, _ = fill(template, tmp_path)

        document = docx.Document(output)
        section = document.element.body.xpath('//w:sdt[w:sdtPr/w:tag/@w:val="order_items"]')[0]
        assert section.xpath('./w:sdtContent//w:sdtPr/w:id') == []

    def test_fills_the_first_deep_run_of_a_composite_control(self, template, tmp_path):
        output, _ = fill(template, tmp_path)

        assert tag_text(output, 'return_change') == '5,00'

    def test_creates_a_run_in_an_empty_paragraph(self, template, tmp_path):
        output, _ = fill(template, tmp_path)

        assert tag_text(output, 'notes') == 'sem cebola'

    def test_skips_empty_values_when_asked(self, template, tmp_path):
        mapping = {**MAPPING, 'notes': {'mapTo': 'Observações', 'skipIfEmpty': True}}
        output, filled = fill(template, tmp_path, {**DATA, 'Observações': ''}, mapping)

        assert filled == 5
        assert tag_text(output, 'notes') == ''

    def test_ignores_tags_missing_in_the_template(self, template, tmp_path):
        _, filled = fill(template, tmp_path, {**DATA, 'x': 1}, {**MAPPING, 'not_in_template': 'x'})

        assert filled == 6

    def test_applies_default_and_padding_options(self, template, tmp_path):
        mapping = {**MAPPING, 'clientname': {'mapTo': 'Nome do Cliente', 'default': 'Retirada', 'padLeft': '10,*'}}
        output, _ = fill(template, tmp_path, {**DATA, 'Nome do Cliente': ''}, mapping)

        assert tag_text(output, 'clientname') == '**Retirada'

    def test_accepts_options_as_a_json_string(self, template, tmp_path):
        mapping = {**MAPPING, 'total': '{"mapTo": "Total", "suffix": " reais"}'}
        output, _ = fill(template, tmp_path, mapping=mapping)

        assert tag_text(output, 'total') == '95 reais'


class TestMapping:
    def test_rejects_a_field_mapped_to_missing_data(self):
        with pytest.raises(DataError):
            parse_field_mapping({'clientname': 'Inexistente'}, DATA)

    def test_rejects_a_mapping_that_is_not_an_object(self):
        with pytest.raises(DataError):
            parse_field_mapping(['clientname'], DATA)

    def test_rejects_options_without_map_to(self):
        with pytest.raises(DataError):
            parse_field_mapping({'clientname': {'prefix': '#'}}, DATA)


class TestMacroEnabled:
    def test_describes_and_fills_a_docm_keeping_its_type(self, tmp_path):
        import zipfile
        from template import MACRO_ENABLED_TYPE, build_template

        template = build_template(str(tmp_path / 'cupom.docm'))
        assert len(describe_fields(template)) == 6

        field_mapping = parse_field_mapping(MAPPING, DATA)
        output, filled = fill_template(template, DATA, field_mapping, output_path(tmp_path, 'saida.docm'), False)

        assert filled == 6
        with zipfile.ZipFile(output) as package:
            assert MACRO_ENABLED_TYPE in package.read('[Content_Types].xml').decode()
        assert tag_text(output, 'clientname') == 'João da Silva'

import locale

import pytest

from _docfill import FieldOptions, PropValueParser, format_currency


@pytest.mark.parametrize('value, expected', [
    (95, '95,00'),
    (0, '0,00'),
    (1.3, '1,30'),
    (1234.5, '1.234,50'),
    (-1234.5, '-1.234,50'),
    (1234567.891, '1.234.567,89'),
    ('95.5', '95,50'),
])
def test_formats_money_in_pt_br(value, expected):
    assert format_currency(value) == expected


def test_does_not_touch_the_system_locale(monkeypatch):
    def fail(*args):
        raise locale.Error('unsupported locale setting')

    monkeypatch.setattr(locale, 'setlocale', fail)
    options = FieldOptions.from_json({'mapTo': 'Total', 'format': 'currency'})

    assert PropValueParser(options).parse(95) == '95,00'


def test_formats_the_default_of_an_empty_value():
    options = FieldOptions.from_json({'mapTo': 'Desconto', 'format': 'currency', 'default': '0'})

    assert PropValueParser(options).parse(None) == '0,00'


def test_keeps_a_value_that_is_not_a_number():
    assert format_currency('a combinar') == 'a combinar'


def test_combines_with_prefix_and_padding():
    options = FieldOptions.from_json({'mapTo': 'Total', 'format': 'currency', 'prefix': 'R$ ', 'padLeft': '8'})

    assert PropValueParser(options).parse(95) == 'R$    95,00'

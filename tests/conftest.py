import json
import os

import pytest

from template import build_template

DATA = {
    "Nome do Cliente": "João da Silva",
    "Total": 95,
    "Forma de Pagamento": "credit_online",
    "Itens": [["1x", "Coca-Cola 200ml", "1,30"], ["2x", "Pão de queijo", "8,00"], ["1x", "Café", "4,50"]],
    "Troco": "5,00",
    "Observações": "sem cebola",
}

MAPPING = {
    "clientname": "Nome do Cliente",
    "total": {"mapTo": "Total", "prefix": "R$ "},
    "payment_method": "Forma de Pagamento",
    "order_items": "Itens",
    "return_change": "Troco",
    "notes": "Observações",
}


@pytest.fixture
def template(tmp_path):
    return build_template(str(tmp_path / 'cupom.docx'))


@pytest.fixture
def data_file(tmp_path):
    path = tmp_path / 'dados.json'
    path.write_text(json.dumps(DATA, ensure_ascii=False), encoding='utf-8')
    return str(path)


@pytest.fixture
def mapping_file(tmp_path):
    path = tmp_path / 'mapeamento.json'
    path.write_text(json.dumps(MAPPING, ensure_ascii=False), encoding='utf-8')
    return str(path)


def read_text(path: str) -> str:
    import docx
    document = docx.Document(path)
    return ''.join(document.element.body.xpath('.//w:t/text()'))


def tag_text(path: str, tag: str) -> str:
    import docx
    document = docx.Document(path)
    sdt = document.element.body.xpath(f'//w:sdt[w:sdtPr/w:tag/@w:val="{tag}"]')[0]
    return ''.join(sdt.xpath('./w:sdtContent//w:t/text()'))


def output_path(tmp_path, name='saida.docx') -> str:
    return os.path.join(str(tmp_path), name)

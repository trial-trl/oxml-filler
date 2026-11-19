import json
import os
import subprocess
import sys

from app_entry import handle
from conftest import DATA, MAPPING, output_path, tag_text

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class TestHandle:
    def test_fills_like_the_command_line(self, template, tmp_path):
        output = output_path(tmp_path)
        result = handle({'mode': 'fill', 'template': template, 'data': DATA, 'mapping': MAPPING,
                         'output': output})

        assert result == {'ok': True, 'fields': 6, 'output': output}
        assert tag_text(output, 'payment_method') == 'Crédito via internet'

    def test_describes_the_template(self, template):
        result = handle({'mode': 'describe', 'template': template})

        assert result['ok'] is True
        assert len(result['fields']) == 6

    def test_answers_errors_as_json(self, tmp_path):
        result = handle({'mode': 'describe', 'template': str(tmp_path / 'nada.docx')})

        assert result['ok'] is False
        assert 'Template não encontrado' in result['error']

    def test_rejects_an_unknown_mode(self):
        assert handle({'mode': 'print'}) == {'ok': False, 'error': "Pedido desconhecido: 'print'"}

    def test_rejects_a_payload_that_is_not_an_object(self):
        assert handle([]) == {'ok': False, 'error': 'O pedido precisa ser um objeto JSON'}


def run_main(tmp_path, payload):
    request = tmp_path / 'pedido.json'
    response = tmp_path / 'resposta.json'
    request.write_text(payload if isinstance(payload, str) else json.dumps(payload), encoding='utf-8')
    env = {**os.environ, 'OXML_FILLER_REQUEST': str(request), 'OXML_FILLER_RESPONSE': str(response)}
    process = subprocess.run([sys.executable, os.path.join(ROOT, 'main.py')], env=env, cwd=ROOT,
                             capture_output=True)
    return process, response


class TestMain:
    def test_answers_the_request_in_the_response_file(self, template, tmp_path):
        output = output_path(tmp_path)
        process, response = run_main(tmp_path, {'mode': 'fill', 'template': template, 'data': DATA,
                                                'mapping': MAPPING, 'output': output})

        assert process.returncode == 0
        assert json.loads(response.read_text(encoding='utf-8')) == {'ok': True, 'fields': 6, 'output': output}
        assert not os.path.exists(f'{response}.tmp')

    def test_answers_an_invalid_request_with_an_error(self, tmp_path):
        process, response = run_main(tmp_path, '{nada')

        assert process.returncode == 0
        result = json.loads(response.read_text(encoding='utf-8'))
        assert result['ok'] is False
        assert result['error'].startswith('Pedido inválido')

    def test_does_nothing_without_the_response_variable(self):
        env = {key: value for key, value in os.environ.items() if not key.startswith('OXML_FILLER_')}
        process = subprocess.run([sys.executable, os.path.join(ROOT, 'main.py')], env=env, cwd=ROOT,
                                 capture_output=True)

        assert process.returncode == 0
        assert process.stdout == b''

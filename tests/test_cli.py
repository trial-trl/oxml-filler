import json
import os
import subprocess
import sys

import pytest

from conftest import output_path, tag_text

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def cli(*args):
    # a console without UTF-8 must still get the accents right
    env = {**os.environ, 'PYTHONIOENCODING': 'ascii', 'LC_ALL': 'C'}
    return subprocess.run([sys.executable, os.path.join(ROOT, 'index.py'), *args],
                          capture_output=True, env=env, cwd=ROOT)


def last_json(process):
    lines = [line for line in process.stdout.decode('utf-8').splitlines() if line.strip()]
    return json.loads(lines[-1])


def fill_args(template, data_file, mapping_file, output):
    return ['--input', template, '--data-file', data_file, '--field-mapping-file', mapping_file,
            '--output', output, '--no-open', '--no-temp', '--quiet', '--json']


class TestFill:
    def test_fills_from_files_and_answers_json(self, template, data_file, mapping_file, tmp_path):
        output = output_path(tmp_path)
        process = cli(*fill_args(template, data_file, mapping_file, output))

        assert process.returncode == 0
        assert last_json(process) == {'ok': True, 'fields': 6, 'output': output}
        assert tag_text(output, 'clientname') == 'João da Silva'

    def test_quiet_prints_only_the_json(self, template, data_file, mapping_file, tmp_path):
        process = cli(*fill_args(template, data_file, mapping_file, output_path(tmp_path)))

        assert len(process.stdout.decode('utf-8').strip().splitlines()) == 1

    def test_keeps_the_inline_arguments(self, template, tmp_path):
        output = output_path(tmp_path)
        process = cli('--input', template, '--data', json.dumps({'Nome': 'Ana'}),
                      '--field_mapping', json.dumps({'clientname': 'Nome'}),
                      '--output', output, '--no-open')

        assert process.returncode == 0
        assert f"Output file created in '{output}'" in process.stdout.decode('utf-8')
        assert tag_text(output, 'clientname') == 'Ana'

    def test_exits_1_for_invalid_json_data(self, template, mapping_file, tmp_path):
        data_file = tmp_path / 'ruim.json'
        data_file.write_text('{nada', encoding='utf-8')
        process = cli(*fill_args(template, str(data_file), mapping_file, output_path(tmp_path)))

        assert process.returncode == 1
        assert last_json(process) == {'ok': False, 'error': "'data' must be a valid JSON"}

    def test_exits_1_for_a_field_mapped_to_missing_data(self, template, data_file, tmp_path):
        mapping_file = tmp_path / 'mapa.json'
        mapping_file.write_text(json.dumps({'clientname': 'Inexistente'}), encoding='utf-8')
        process = cli(*fill_args(template, data_file, str(mapping_file), output_path(tmp_path)))

        assert process.returncode == 1
        assert 'Inexistente' in last_json(process)['error']

    def test_exits_1_for_an_unknown_option(self):
        process = cli('--nope', '--json')

        assert process.returncode == 1
        assert last_json(process)['ok'] is False

    def test_exits_2_for_a_missing_template(self, data_file, mapping_file, tmp_path):
        process = cli(*fill_args(str(tmp_path / 'nada.docx'), data_file, mapping_file,
                                 output_path(tmp_path)))

        assert process.returncode == 2
        assert 'Template não encontrado' in last_json(process)['error']

    def test_writes_errors_to_stderr_without_json(self, tmp_path):
        process = cli('--input', str(tmp_path / 'nada.docx'), '--describe')

        assert process.returncode == 2
        assert 'Template não encontrado' in process.stderr.decode('utf-8')


class TestDescribe:
    def test_answers_the_fields_as_json(self, template):
        process = cli('--input', template, '--describe', '--json')

        assert process.returncode == 0
        result = last_json(process)
        assert result['ok'] is True
        assert [field['tag'] for field in result['fields']][:3] == ['clientname', 'total', 'payment_method']
        assert 'Crédito via internet' in result['fields'][2]['items']

    def test_exits_2_for_a_file_that_is_not_word(self, tmp_path):
        path = tmp_path / 'texto.docx'
        path.write_text('texto', encoding='utf-8')
        process = cli('--input', str(path), '--describe', '--json')

        assert process.returncode == 2


class TestVersion:
    @pytest.mark.parametrize('args', [['--version', '--json'], ['--input', 'qualquer.docx', '--version', '--json']])
    def test_answers_the_version_as_json(self, args):
        process = cli(*args)

        assert process.returncode == 0
        assert last_json(process) == {'ok': True, 'version': '1.1.0'}

    def test_prints_the_version(self):
        assert cli('--version').stdout.decode('utf-8').strip() == '1.1.0'

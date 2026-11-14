"""Integratietests in tijdelijke mappen; de echte collectie wordt niet gewijzigd."""
import base64
import hashlib
import http.client
import json
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch
import zipfile
from xml.sax.saxutils import escape

import bijwerken as b
import server

PNG = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jrZkAAAAASUVORK5CYII=')


def workbook(path, brand='Ford', location=None):
    headers = ''.join(f'<c r="{chr(65+i)}1" t="inlineStr"><is><t>{escape(name)}</t></is></c>' for i, name in enumerate(b.FIELDS + ['opmerking']))
    cells = '<c r="A2"><v>1</v></c>'
    for column, value in [('B', brand), ('C', 'Mustang'), ('H', '1:18')]:
        cells += f'<c r="{column}2" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
    if location is not None:
        cells += f'<c r="J2" t="inlineStr"><is><t>{escape(location)}</t></is></c>'
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Inventaris" r:id="r1"/></sheets></workbook>')
        archive.writestr('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>')
        archive.writestr('xl/worksheets/sheet1.xml', f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">{headers}</row><row r="2">{cells}</row><row r="3"><c r="A3"><v>2</v></c></row></sheetData></worksheet>')


class InventoryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.source = self.directory / 'Inventaris.xlsx'
        workbook(self.source)
        (self.directory / 'inventaris.template.html').write_text((server.CODE_DIR / 'inventaris.template.html').read_text(encoding='utf-8'), encoding='utf-8')
        self.source_patch = patch.object(b, 'SOURCE', self.source)
        self.source_patch.start()
        self.addCleanup(self.source_patch.stop)
        b.main(self.directory)
        self.http = server.create_server(self.directory, port=0)
        self.thread = threading.Thread(target=self.http.serve_forever, daemon=True)
        self.thread.start()
        self.addCleanup(self.stop_server)
        self.origin = f'http://127.0.0.1:{self.http.server_port}'

    def stop_server(self):
        self.http.shutdown()
        self.http.server_close()
        self.thread.join()

    def request(self, method, route, body=None, headers=None):
        client = http.client.HTTPConnection('127.0.0.1', self.http.server_port, timeout=10)
        client.request(method, route, body=body, headers=headers or {})
        response = client.getresponse()
        raw = response.read()
        status = response.status
        client.close()
        return status, raw

    def metadata(self):
        values = [''] * 15
        values[1:4] = ['Citroën', 'DS <Special>', '1967']
        values[7] = '1:18'
        values[11] = '75,50'
        return {'id': '', 'requestId': 'test-request', 'values': values, 'note': 'Voorzichtig', 'ad': 'Tekst\n€ 75,50 <script>geen HTML</script>', 'photos': [{'upload': 'photo_0'}]}

    def post(self, metadata, uploads=None, revision=None, origin=None):
        boundary = 'test-inventaris-boundary'
        body = bytearray()
        fields = {'metadata': json.dumps(metadata).encode(), **(uploads if uploads is not None else {'photo_0': PNG})}
        for name, raw in fields.items():
            body.extend(f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"\r\n\r\n'.encode())
            body.extend(raw)
            body.extend(b'\r\n')
        body.extend(f'--{boundary}--\r\n'.encode())
        headers = {'Content-Type': 'multipart/form-data; boundary=' + boundary, 'Origin': origin or self.origin, 'X-Inventaris-Token': self.http.token, 'X-Inventaris-Revision': revision or server.revision(self.directory)}
        status, raw = self.request('POST', '/api/models', body, headers)
        return status, json.loads(raw)

    def test_create_edit_refresh_and_restart(self):
        before = hashlib.sha256(self.source.read_bytes()).hexdigest()
        metadata = self.metadata()
        status, response = self.post(metadata)
        self.assertEqual(status, 200, response)
        key = response['id']
        data = b.snapshot(self.directory)
        record = next(r for r in data['records'] if r['values'][0] == key)
        self.assertEqual(record['values'][11], 75.5)
        self.assertEqual(record['values'][2], 'DS <Special>')
        self.assertEqual(record['extra']['ad'], metadata['ad'])
        photo = record['extra']['photos'][0]
        self.assertEqual((self.directory / photo).read_bytes(), PNG)
        self.assertEqual(self.request('GET', '/' + photo), (200, PNG))
        # Een herhaald verzoek mag geen dubbel model maken.
        self.assertEqual(self.post(metadata)[1]['id'], key)
        self.assertEqual(len(b.load_local(self.directory)['models']), 1)
        edit = {**metadata, 'id': key, 'identity': record['values'][1:6], 'photos': [{'path': photo}]}
        edit['values'] = metadata['values'][:]
        edit['values'][11] = '80'
        self.assertEqual(self.post(edit, {})[0], 200)
        b.main(self.directory)
        server.rebuild(self.directory)
        self.assertEqual(len(b.snapshot(self.directory)['records']), 2)
        self.assertEqual(b.load_local(self.directory)['models'][0]['values'][11], 80)
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), before)
        import pyarrow.parquet as pq
        self.assertEqual(pq.read_table(self.directory / 'inventaris.parquet').num_rows, 2)

    def test_excel_media_survives_refresh_and_identity_change(self):
        record = b.snapshot(self.directory)['records'][0]
        metadata = self.metadata()
        metadata.update(id='1', identity=record['values'][1:6], values=[str(v) for v in record['values']], note=record['note'])
        self.assertEqual(self.post(metadata)[0], 200)
        b.main(self.directory)
        self.assertEqual(b.snapshot(self.directory)['records'][0]['extra']['ad'], metadata['ad'])
        workbook(self.source, brand='Buick')
        b.main(self.directory)
        data = b.snapshot(self.directory)
        self.assertNotIn('extra', data['records'][0])
        self.assertEqual(len(data['localWarnings']), 1)
        self.assertIn('1', b.load_local(self.directory)['excel_media'])
        # De oude modelidentiteit mag niet stilzwijgend aan het nieuwe model worden gekoppeld.
        self.assertEqual(self.post(metadata)[0], 400)

    def state_change(self, record, *, restore=False):
        metadata = {'id': str(record['values'][0]), 'identity': record['values'][1:6], 'origin': record['origin']}
        headers = {'Content-Type': 'application/json', 'Origin': self.origin, 'X-Inventaris-Token': self.http.token, 'X-Inventaris-Revision': server.revision(self.directory)}
        status, raw = self.request('POST', '/api/models/' + ('restore' if restore else 'delete'), json.dumps(metadata).encode(), headers)
        return status, json.loads(raw)

    def test_edit_excel_fields_preserves_other_excel_updates(self):
        before = hashlib.sha256(self.source.read_bytes()).hexdigest()
        record = b.snapshot(self.directory)['records'][0]
        metadata = {'id': '1', 'identity': record['values'][1:6], 'values': [str(v) for v in record['values']], 'note': 'Nieuwe opmerking', 'photos': [], 'ad': 'Zoekertje'}
        metadata['values'][2] = 'Mustang GT'
        metadata['values'][11] = '82,50'
        self.assertEqual(self.post(metadata, {})[0], 200)
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), before)
        for _ in range(2):
            server.rebuild(self.directory)
        workbook(self.source, location='Nieuwe kast')
        b.main(self.directory)
        record = b.snapshot(self.directory)['records'][0]
        self.assertEqual(record['values'][2], 'Mustang GT')
        self.assertEqual(record['values'][9], 'Nieuwe kast')
        self.assertEqual(record['values'][11], 82.5)
        self.assertEqual(record['note'], 'Nieuwe opmerking')
        self.assertTrue(record['locallyEdited'])
        # Een tweede bewerking gebruikt nog steeds de oorspronkelijke Excel-identiteit.
        metadata.update(identity=record['values'][1:6], values=[str(v) for v in record['values']])
        metadata['values'][11] = '90'
        self.assertEqual(self.post(metadata, {})[0], 200)
        b.main(self.directory)
        self.assertEqual(b.snapshot(self.directory)['records'][0]['values'][11], 90)
        self.assertEqual(b.load_local(self.directory)['excel_edits']['1']['identity'][1], 'Mustang')

    def test_delete_restore_local_and_excel_models(self):
        before = hashlib.sha256(self.source.read_bytes()).hexdigest()
        self.assertEqual(self.post(self.metadata())[0], 200)
        records = b.snapshot(self.directory)['records']
        for record in records:
            self.assertEqual(self.state_change(record)[0], 200)
        data = b.snapshot(self.directory)
        self.assertEqual(data['records'], [])
        self.assertEqual(len(data['trash']), 2)
        b.main(self.directory)
        server.rebuild(self.directory)
        self.assertEqual(b.snapshot(self.directory)['records'], [])
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), before)
        for record in records:
            self.assertEqual(self.state_change(record, restore=True)[0], 200)
        data = b.snapshot(self.directory)
        self.assertEqual(len(data['records']), 2)
        self.assertEqual(data['trash'], [])
        local_record = next(r for r in data['records'] if r['origin'] == 'local')
        self.assertEqual(local_record['extra']['ad'], self.metadata()['ad'])
        self.assertTrue((self.directory / local_record['extra']['photos'][0]).exists())

    def test_deleted_edited_excel_model_can_be_restored_without_refresh(self):
        record = b.snapshot(self.directory)['records'][0]
        metadata = {'id': '1', 'identity': record['values'][1:6], 'values': [str(v) for v in record['values']], 'note': '', 'photos': [], 'ad': 'Bewaarde tekst'}
        metadata['values'][2] = 'Gewijzigd model'
        self.assertEqual(self.post(metadata, {})[0], 200)
        record = b.snapshot(self.directory)['records'][0]
        self.assertEqual(self.state_change(record)[0], 200)
        server.rebuild(self.directory)
        self.assertEqual(b.snapshot(self.directory)['records'], [])
        self.assertEqual(self.state_change(record, restore=True)[0], 200)
        restored = b.snapshot(self.directory)['records'][0]
        self.assertEqual(restored['values'][2], 'Gewijzigd model')
        self.assertEqual(restored['extra']['ad'], 'Bewaarde tekst')

    def test_deleted_excel_number_reused_for_other_model_is_not_hidden(self):
        record = b.snapshot(self.directory)['records'][0]
        self.assertEqual(self.state_change(record)[0], 200)
        workbook(self.source, brand='Buick')
        b.main(self.directory)
        data = b.snapshot(self.directory)
        self.assertEqual(len(data['records']), 1)
        self.assertEqual(data['records'][0]['values'][1], 'Buick')
        self.assertEqual(len(data['localWarnings']), 1)
        self.assertEqual(len(data['trash']), 1)

    def test_invalid_input_does_not_change_collection(self):
        before = (self.directory / 'index.html').read_bytes()
        bad = self.metadata()
        bad['values'][1] = ''
        self.assertEqual(self.post(bad)[0], 400)
        bad = self.metadata()
        bad['photos'] = [{'path': '../Inventaris.xlsx'}]
        self.assertEqual(self.post(bad)[0], 400)
        self.assertEqual(self.post(self.metadata(), {'photo_0': b'<svg/>'})[0], 400)
        self.assertEqual((self.directory / 'index.html').read_bytes(), before)
        self.assertFalse((self.directory / b.LOCAL_FILE).exists())

    def test_security_and_stale_changes(self):
        self.assertEqual(self.post(self.metadata(), origin='https://example.com')[0], 403)
        self.assertEqual(self.post(self.metadata(), revision='outdated')[0], 409)
        self.assertEqual(self.request('GET', '/lokale_modellen.json')[0], 404)
        self.assertEqual(self.request('GET', '/fotos/../Inventaris.xlsx')[0], 403)
        self.assertEqual(self.request('GET', '/api/config', headers={'Host': 'example.com'})[0], 403)

    def test_busy_port_is_exclusive_and_automatic_start_uses_free_port(self):
        busy_port = self.http.server_port
        with self.assertRaises(OSError):
            server.create_server(self.directory, busy_port)
        with patch.object(server, 'DEFAULT_PORT', busy_port):
            other = server.start_server(self.directory)
        try:
            self.assertNotEqual(other.server_port, busy_port)
            self.assertGreater(other.server_port, 0)
            # De oorspronkelijke app blijft bereikbaar op haar eigen poort.
            self.assertEqual(self.request('GET', '/')[0], 200)
        finally:
            other.server_close()

    def test_save_without_readable_excel_and_photo_order(self):
        metadata = self.metadata()
        metadata['photos'] = [{'upload': 'second'}, {'upload': 'first'}]
        with patch.object(b, 'SOURCE', self.directory / 'missing.xlsx'):
            self.assertEqual(self.post(metadata, {'first': PNG, 'second': PNG + b'second'})[0], 200)
        record = b.load_local(self.directory)['models'][0]
        self.assertEqual((self.directory / record['extra']['photos'][0]).read_bytes(), PNG + b'second')
        self.assertEqual((self.directory / record['extra']['photos'][1]).read_bytes(), PNG)


if __name__ == '__main__':
    unittest.main()

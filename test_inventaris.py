"""Integratietests in tijdelijke mappen; de echte collectie wordt niet gewijzigd."""
import base64
import hashlib
import io
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
import collectie
import beheer

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

    def appointment(self, record, item):
        metadata = {'id': str(record['values'][0]), 'identity': record['values'][1:6], 'appointment': item}
        headers = {'Content-Type': 'application/json', 'Origin': self.origin, 'X-Inventaris-Token': self.http.token, 'X-Inventaris-Revision': server.revision(self.directory)}
        status, raw = self.request('POST', '/api/appointments', json.dumps(metadata).encode(), headers)
        return status, json.loads(raw)

    def test_appointments_save_survive_refresh_and_delete(self):
        self.assertEqual(self.post(self.metadata())[0], 200)
        records = b.snapshot(self.directory)['records']
        excel = next(r for r in records if r['origin'] == 'excel')
        local = next(r for r in records if r['origin'] == 'local')
        item = {'start': '2026-11-20T14:30', 'duration': 30, 'buyer': 'Jan <06>', 'note': 'Contant'}
        for record in (excel, local):
            self.assertEqual(self.appointment(record, item)[0], 200)
        # Ongeldige invoer verandert niets.
        self.assertEqual(self.appointment(excel, {**item, 'start': '20-11-2026'})[0], 400)
        self.assertEqual(self.appointment(excel, {**item, 'duration': 0})[0], 400)
        b.main(self.directory)
        server.rebuild(self.directory)
        data = b.snapshot(self.directory)
        self.assertEqual([r.get('appointment') for r in data['records']], [item, item])
        self.assertEqual(data['localWarnings'], [])
        # Bewerken van een model behoudt de afspraak.
        record = next(r for r in data['records'] if r['origin'] == 'excel')
        edit = {'id': '1', 'identity': record['values'][1:6], 'values': [str(v) for v in record['values']], 'note': '', 'photos': [], 'ad': ''}
        edit['values'][14] = 'Gereserveerd'
        self.assertEqual(self.post(edit, {})[0], 200)
        self.assertEqual(b.snapshot(self.directory)['records'][0]['appointment'], item)
        # Verwijderen van de afspraak.
        for record in b.snapshot(self.directory)['records']:
            self.assertEqual(self.appointment(record, None)[0], 200)
        self.assertTrue(all('appointment' not in r for r in b.snapshot(self.directory)['records']))
        self.assertNotIn('appointments', b.load_local(self.directory))

    def test_chat_url_saved_validated_and_kept(self):
        record = b.snapshot(self.directory)['records'][0]
        metadata = {'id': '1', 'identity': record['values'][1:6], 'values': [str(v) for v in record['values']], 'note': '', 'photos': [], 'ad': '', 'chatUrl': 'https://chatgpt.com/c/abc-123'}
        self.assertEqual(self.post(metadata, {})[0], 200)
        self.assertEqual(self.post({**metadata, 'chatUrl': 'javascript:alert(1)'}, {})[0], 400)
        # Een statuswijziging zonder chatUrl-veld behoudt de URL.
        quick = {k: v for k, v in metadata.items() if k != 'chatUrl'}
        quick['values'] = metadata['values'][:]
        quick['values'][14] = 'Online'
        self.assertEqual(self.post(quick, {})[0], 200)
        b.main(self.directory)
        server.rebuild(self.directory)
        record = b.snapshot(self.directory)['records'][0]
        self.assertEqual(record['extra']['chatUrl'], 'https://chatgpt.com/c/abc-123')
        self.assertEqual(record['values'][14], 'Online')
        # Leegmaken verwijdert de URL.
        self.assertEqual(self.post({**metadata, 'identity': record['values'][1:6], 'chatUrl': ''}, {})[0], 200)
        self.assertNotIn('chatUrl', b.snapshot(self.directory)['records'][0]['extra'])

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

    def json_post(self, route, metadata):
        headers = {'Content-Type': 'application/json', 'Origin': self.origin, 'X-Inventaris-Token': self.http.token, 'X-Inventaris-Revision': server.revision(self.directory)}
        status, raw = self.request('POST', route, json.dumps(metadata).encode(), headers)
        return status, json.loads(raw)

    def test_full_export_contains_excel_local_media_sales_and_appointments(self):
        self.assertEqual(self.post(self.metadata())[0], 200)
        record = next(r for r in b.snapshot(self.directory)['records'] if r['origin'] == 'local')
        self.assertEqual(self.appointment(record, {'start': '2026-11-20T14:30', 'duration': 30, 'buyer': 'Jan', 'note': ''})[0], 200)
        status = {'id': record['values'][0], 'identity': record['values'][1:6], 'status': 'Verkocht', 'sale': {'price': '70,25', 'date': '2026-11-17'}}
        self.assertEqual(self.json_post('/api/status', status)[0], 200)
        import pyarrow.parquet as pq
        code, raw = self.request('GET', '/api/export/collectie.parquet')
        self.assertEqual(code, 200)
        table = pq.read_table(io.BytesIO(raw))
        self.assertEqual(table.num_rows, 2)
        rows = table.to_pylist()
        exported = next(r for r in rows if r['Bron'] == 'local')
        self.assertEqual(exported['Verkoopprijs'], 70.25)
        self.assertEqual(exported['Verkoopdatum'], '2026-11-17')
        self.assertEqual(json.loads(exported['Afspraak'])['buyer'], 'Jan')
        self.assertEqual(exported['Zoekertjestekst'], self.metadata()['ad'])
        code, raw = self.request('GET', '/api/export/collectie.zip')
        self.assertEqual(code, 200)
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            self.assertEqual(archive.read(exported['Foto’s'][0]), PNG)
            self.assertEqual(json.loads(archive.read('export-info.json'))['models'], 2)
            self.assertEqual(len(json.loads(archive.read('collectie.json'))['records']), 2)
            self.assertIn('lokale_modellen.json', archive.namelist())

    def test_automatic_backups_restore_data_and_photo_bytes(self):
        source_hash = hashlib.sha256(self.source.read_bytes()).hexdigest()
        self.assertEqual(self.post(self.metadata())[0], 200)
        record = next(r for r in b.snapshot(self.directory)['records'] if r['origin'] == 'local')
        original_photo = record['extra']['photos'][0]
        edit = {**self.metadata(), 'id': record['values'][0], 'identity': record['values'][1:6], 'values': [str(v) for v in record['values']], 'photos': [{'path': original_photo}]}
        edit['values'][11] = '99'
        self.assertEqual(self.post(edit, {})[0], 200)
        backups = collectie.list_backups(self.directory)
        self.assertEqual(len(backups), 2)
        self.assertEqual(len(list((self.directory / 'backups' / 'fotobestanden').iterdir())), 1)
        (self.directory / original_photo).unlink()
        self.assertEqual(self.json_post('/api/backups/restore', {'id': backups[0]['id']})[0], 200)
        restored = next(r for r in b.snapshot(self.directory)['records'] if r['origin'] == 'local')
        self.assertEqual(restored['values'][11], 75.5)
        self.assertEqual((self.directory / original_photo).read_bytes(), PNG)
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), source_hash)
        # Ook terug naar vóór het eerste paginamodel werkt; herstellen maakt een terugweg.
        self.assertEqual(self.json_post('/api/backups/restore', {'id': backups[-1]['id']})[0], 200)
        self.assertEqual(len(b.snapshot(self.directory)['records']), 1)
        self.assertGreaterEqual(len(collectie.list_backups(self.directory)), 4)

    def test_corrupt_backup_photo_rejected_before_data_is_changed(self):
        self.assertEqual(self.post(self.metadata())[0], 200)
        self.assertEqual(self.json_post('/api/backups/create', {})[0], 200)
        backup = collectie.list_backups(self.directory)[0]
        pool = self.directory / 'backups' / 'fotobestanden'
        next(pool.iterdir()).write_bytes(b'damaged')
        before = (self.directory / b.LOCAL_FILE).read_bytes()
        self.assertEqual(self.json_post('/api/backups/restore', {'id': backup['id']})[0], 400)
        self.assertEqual((self.directory / b.LOCAL_FILE).read_bytes(), before)
        self.assertEqual(self.json_post('/api/backups/restore', {'id': '../index.html'})[0], 400)

    def test_duplicate_warning_requires_explicit_confirmation(self):
        self.assertEqual(self.post(self.metadata())[0], 200)
        duplicate = self.metadata()
        duplicate['requestId'] = 'another-copy'
        before = (self.directory / b.LOCAL_FILE).read_bytes()
        code, result = self.post(duplicate)
        self.assertEqual(code, 409)
        self.assertEqual(len(result['duplicates']), 1)
        self.assertEqual((self.directory / b.LOCAL_FILE).read_bytes(), before)
        duplicate['confirmDuplicate'] = True
        self.assertEqual(self.post(duplicate)[0], 200)
        data = b.snapshot(self.directory)
        copies = [r for r in data['records'] if r['origin'] == 'local']
        self.assertEqual(len(copies), 2)
        self.assertTrue(all(len(r['quality']['duplicates']) == 1 for r in copies))
        # Snel een status kiezen mag ook bij een bevestigd tweede exemplaar.
        record = copies[0]
        self.assertEqual(self.json_post('/api/status', {'id': record['values'][0], 'identity': record['values'][1:6], 'status': 'Online'})[0], 200)

    def test_publication_checks_prices_and_missing_fields(self):
        record = {'values': ['', 'Ford', 'Mustang', 1967, 'Maker', '123', 'Goed', '1:18', 'Onbekend', 'Kast', 'Nee', 75.5, 'Onbekend', 'Onbekend', 'Online'], 'extra': {'photos': ['fotos/01.jpg'], 'ad': '**Vraagprijs: € 75,50**'}}
        self.assertEqual(collectie.checks(record), [])
        record['extra']['ad'] = 'Bouwjaar 1967, artikel 123. Vraagprijs: € 85.'
        self.assertEqual([i['code'] for i in collectie.checks(record)], ['price_mismatch'])
        record['extra'] = {}
        record['values'][11] = 'Onbekend'
        self.assertEqual({i['code'] for i in collectie.checks(record)}, {'photos', 'ad', 'price'})
        self.assertEqual(collectie.ad_prices('Vraagprijs: € 1.250,50'), [1250.5])
        self.assertEqual(collectie.ad_prices('Model 123 uit 1967'), [])

    def test_quick_status_records_real_sale_and_preserves_excel_link(self):
        record = b.snapshot(self.directory)['records'][0]
        metadata = {'id': '1', 'identity': record['values'][1:6], 'status': 'Verkocht', 'sale': {'price': '0', 'date': '2026-11-17'}}
        self.assertEqual(self.json_post('/api/status', metadata)[0], 200)
        b.main(self.directory)
        server.rebuild(self.directory)
        saved = b.snapshot(self.directory)['records'][0]
        self.assertEqual(saved['values'][14], 'Verkocht')
        self.assertEqual(saved['sale'], {'price': 0, 'date': '2026-11-17'})
        for bad in ({'price': '-1', 'date': '2026-11-17'}, {'price': '10', 'date': 'not-a-date'}):
            self.assertEqual(self.json_post('/api/status', {**metadata, 'sale': bad})[0], 400)

    def make_bundle(self, records, **changes):
        metadata = {'requestId': 'bundle-test', 'name': 'Klassiekers', 'buyer': 'Jan', 'contact': 'Testcontact', 'total': '120,50', 'status': 'Gereserveerd', 'members': [{'id': str(r['values'][0]), 'identity': r['values'][1:6]} for r in records], 'appointment': {'start': '2026-11-20T14:30', 'duration': 30}, 'note': 'Samen afhalen'}
        metadata.update(changes)
        return self.json_post('/api/bundles', metadata)

    def test_bundle_reservation_sale_refresh_and_cancellation(self):
        self.assertEqual(self.post(self.metadata())[0], 200)
        records = b.snapshot(self.directory)['records']
        individual = {'start': '2026-11-21T09:00', 'duration': 30, 'buyer': 'Eerdere afspraak', 'note': ''}
        self.assertEqual(self.appointment(records[0], individual)[0], 200)
        code, result = self.make_bundle(records)
        self.assertEqual(code, 200, result)
        key = result['id']
        # Een herhaald verzoek maakt geen dubbele bundel.
        self.assertEqual(self.make_bundle(records)[1]['id'], key)
        data = b.snapshot(self.directory)
        self.assertEqual(len(data['bundles']), 1)
        self.assertTrue(all(r['values'][14] == 'Gereserveerd' and r['bundle']['id'] == key for r in data['records']))
        self.assertTrue(all(r['appointment']['start'] == '2026-11-20T14:30' and r['appointmentScope'] == 'bundle' for r in data['records']))
        self.assertEqual(self.make_bundle(data['records'], id=key, status='Verkocht', date='2026-11-17')[0], 200)
        b.main(self.directory)
        server.rebuild(self.directory)
        sold = b.snapshot(self.directory)
        self.assertEqual(sold['bundles'][0]['total'], 120.5)
        self.assertTrue(all(r['values'][14] == 'Verkocht' and r['sale']['price'] is None for r in sold['records']))
        self.assertEqual(self.json_post('/api/bundles', {'id': key, 'action': 'cancel'})[0], 200)
        restored = b.snapshot(self.directory)
        self.assertTrue(all('bundle' not in r for r in restored['records']))
        self.assertEqual(restored['records'][0]['values'][14], records[0]['values'][14])
        self.assertEqual(restored['records'][0]['appointment'], individual)
        self.assertTrue(all(any('bundel' in e['action'] for e in b.load_local(self.directory)['history'][str(r['values'][0])]) for r in restored['records']))

    def test_bundle_conflicts_and_individual_changes_are_rejected_atomically(self):
        self.assertEqual(self.post(self.metadata())[0], 200)
        records = b.snapshot(self.directory)['records']
        key = self.make_bundle(records)[1]['id']
        bundled = b.snapshot(self.directory)['records']
        before = (self.directory / b.LOCAL_FILE).read_bytes()
        self.assertEqual(self.make_bundle(bundled, requestId='other-bundle')[0], 400)
        self.assertEqual(self.state_change(bundled[0])[0], 400)
        self.assertEqual(self.appointment(bundled[0], None)[0], 400)
        self.assertEqual(self.json_post('/api/status', {'id': str(bundled[0]['values'][0]), 'identity': bundled[0]['values'][1:6], 'status': 'Online'})[0], 400)
        self.assertEqual(self.make_bundle(bundled, id=key, total='-1')[0], 400)
        self.assertEqual((self.directory / b.LOCAL_FILE).read_bytes(), before)
        # Gewone modelgegevens blijven bewerkbaar zonder bundelstatus of afspraak te wijzigen.
        record = bundled[0]
        edit = {'id': str(record['values'][0]), 'identity': record['values'][1:6], 'values': [str(v) for v in record['values']], 'note': 'Nieuwe locatie-informatie', 'photos': [], 'ad': ''}
        self.assertEqual(self.post(edit, {})[0], 200)
        self.assertEqual(b.snapshot(self.directory)['records'][0]['bundle']['id'], key)

    def test_bundle_member_replaced_and_excel_identity_guard(self):
        self.assertEqual(self.post(self.metadata())[0], 200)
        third = self.metadata()
        third['requestId'] = 'third-model'
        third['values'][2] = 'Ander model'
        self.assertEqual(self.post(third)[0], 200)
        records = b.snapshot(self.directory)['records']
        key = self.make_bundle(records[:2])[1]['id']
        data = b.snapshot(self.directory)
        self.assertEqual(self.make_bundle([data['records'][0], data['records'][2]], id=key)[0], 200)
        changed = b.snapshot(self.directory)
        self.assertNotIn('bundle', changed['records'][1])
        self.assertEqual(changed['records'][1]['values'][14], records[1]['values'][14])
        workbook(self.source, brand='Buick')
        b.main(self.directory)
        changed = b.snapshot(self.directory)
        excel = changed['records'][0]
        self.assertEqual(excel['values'][1], 'Buick')
        self.assertNotIn('bundle', excel)
        self.assertTrue(changed['localWarnings'])

    def edit_metadata(self, record, **updates):
        metadata = {'id': str(record['values'][0]), 'identity': record['values'][1:6], 'values': [str(v) for v in record['values']], 'note': record.get('note', ''), 'ad': record.get('extra', {}).get('ad', ''), 'photos': [{'path': p} for p in record.get('extra', {}).get('photos', [])]}
        metadata.update(updates)
        return metadata

    def test_history_selective_undo_preserves_later_unrelated_fields(self):
        self.assertEqual(self.post(self.metadata())[0], 200)
        record = next(r for r in b.snapshot(self.directory)['records'] if r['origin'] == 'local')
        key = str(record['values'][0])
        edit = self.edit_metadata(record)
        edit['values'][11] = '90'
        self.assertEqual(self.post(edit, {})[0], 200)
        event = b.load_local(self.directory)['history'][key][-1]
        record = next(r for r in b.snapshot(self.directory)['records'] if str(r['values'][0]) == key)
        self.assertEqual(self.post(self.edit_metadata(record, note='Latere opmerking'), {})[0], 200)
        self.assertEqual(self.json_post('/api/history/undo', {'id': key, 'eventId': event['id']})[0], 200)
        record = next(r for r in b.snapshot(self.directory)['records'] if str(r['values'][0]) == key)
        self.assertEqual(record['values'][11], 75.5)
        self.assertEqual(record['note'], 'Latere opmerking')
        self.assertEqual(b.load_local(self.directory)['history'][key][-1]['undoOf'], event['id'])
        code, raw = self.request('GET', '/api/history?id=' + key)
        self.assertEqual(code, 200)
        self.assertGreaterEqual(len(json.loads(raw)['events']), 4)

    def test_history_conflicting_undo_and_bundle_undo_are_blocked(self):
        self.assertEqual(self.post(self.metadata())[0], 200)
        record = next(r for r in b.snapshot(self.directory)['records'] if r['origin'] == 'local')
        key = str(record['values'][0])
        edit = self.edit_metadata(record)
        edit['values'][11] = '90'
        self.assertEqual(self.post(edit, {})[0], 200)
        event = b.load_local(self.directory)['history'][key][-1]
        record = next(r for r in b.snapshot(self.directory)['records'] if str(r['values'][0]) == key)
        edit = self.edit_metadata(record)
        edit['values'][11] = '100'
        self.assertEqual(self.post(edit, {})[0], 200)
        before = (self.directory / b.LOCAL_FILE).read_bytes()
        self.assertEqual(self.json_post('/api/history/undo', {'id': key, 'eventId': event['id']})[0], 400)
        self.assertEqual((self.directory / b.LOCAL_FILE).read_bytes(), before)
        self.assertEqual(self.make_bundle(b.snapshot(self.directory)['records'])[0], 200)
        event = b.load_local(self.directory)['history'][key][-1]
        self.assertFalse(event['undoable'])
        self.assertEqual(self.json_post('/api/history/undo', {'id': key, 'eventId': event['id']})[0], 400)

    def test_excel_import_history_and_undo_identity_guard(self):
        workbook(self.source, location='Kast 2')
        b.main(self.directory)
        events = b.load_local(self.directory)['history']['1']
        self.assertTrue(any(c['field'] == 'value:9' for c in events[-1]['changes']))
        event = events[-1]
        self.assertEqual(self.json_post('/api/history/undo', {'id': '1', 'eventId': event['id']})[0], 200)
        self.assertEqual(b.snapshot(self.directory)['records'][0]['values'][9], 'Onbekend')
        workbook(self.source, brand='Buick')
        b.main(self.directory)
        self.assertEqual(self.json_post('/api/history/undo', {'id': '1', 'eventId': event['id']})[0], 400)

    def test_history_restores_photo_text_and_appointment(self):
        self.assertEqual(self.post(self.metadata())[0], 200)
        record = next(r for r in b.snapshot(self.directory)['records'] if r['origin'] == 'local')
        key = str(record['values'][0])
        old_photos = record['extra']['photos'][:]
        self.assertEqual(self.post(self.edit_metadata(record, photos=[], ad='Nieuwe tekst'), {})[0], 200)
        event = b.load_local(self.directory)['history'][key][-1]
        self.assertEqual(self.json_post('/api/history/undo', {'id': key, 'eventId': event['id']})[0], 200)
        record = next(r for r in b.snapshot(self.directory)['records'] if str(r['values'][0]) == key)
        self.assertEqual(record['extra']['photos'], old_photos)
        self.assertEqual(record['extra']['ad'], self.metadata()['ad'])
        appointment = {'start': '2026-11-20T15:00', 'duration': 45, 'buyer': 'Jan', 'note': 'Test'}
        self.assertEqual(self.appointment(record, appointment)[0], 200)
        event = b.load_local(self.directory)['history'][key][-1]
        self.assertEqual(self.json_post('/api/history/undo', {'id': key, 'eventId': event['id']})[0], 200)
        record = next(r for r in b.snapshot(self.directory)['records'] if str(r['values'][0]) == key)
        self.assertNotIn('appointment', record)

    def test_bundle_and_history_are_in_full_export(self):
        self.assertEqual(self.post(self.metadata())[0], 200)
        records = b.snapshot(self.directory)['records']
        key = self.make_bundle(records, status='Verkocht', date='2026-11-17')[1]['id']
        code, raw = self.request('GET', '/api/export/collectie.zip')
        self.assertEqual(code, 200)
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            bundles = json.loads(archive.read('verkoopbundels.json'))
            history = json.loads(archive.read('wijzigingsgeschiedenis.json'))
            self.assertEqual(bundles[key]['total'], 120.5)
            self.assertTrue(all(str(r['values'][0]) in history for r in records))
            import pyarrow.parquet as pq
            rows = pq.read_table(io.BytesIO(archive.read('collectie.parquet'))).to_pylist()
            self.assertTrue(all(r['Bundel ID'] == key and r['Verkoopprijs'] is None for r in rows))
            self.assertTrue(all(json.loads(r['Wijzigingsgeschiedenis']) for r in rows))


if __name__ == '__main__':
    unittest.main()

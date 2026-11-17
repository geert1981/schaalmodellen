"""Lokale inventarisapp: Excel blijft leesbron, browsertoevoegingen staan apart."""
import argparse
import copy
import datetime
from email import policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import math
import re
from pathlib import Path
import secrets
import socket
import sys
import threading
from urllib.parse import unquote, urlsplit
import uuid
import webbrowser

import bijwerken
import miniaturen
import collectie
import beheer
from zoneinfo import ZoneInfo

CODE_DIR = Path(__file__).resolve().parent
# Gegevens (index.html, json, parquet, fotos) staan in data/: een junction naar OneDrive.
DIRECTORY = CODE_DIR / 'data'
MAX_BODY = 80 * 1024 * 1024
LOCK = threading.RLock()
THUMB_LOCK = threading.Lock()
DEFAULT_PORT = 18765


def revision(directory):
    return bijwerken.local_revision(bijwerken.load_local(directory))


def atomic_write(path, content):
    if path.name == bijwerken.LOCAL_FILE:
        collectie.create_backup(path.parent, 'Vóór collectie wijzigen')
    temp = path.with_name(path.name + '.tmp')
    try:
        temp.write_text(content, encoding='utf-8')
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


def rebuild(directory):
    data = bijwerken.merge_local(bijwerken.snapshot(directory), bijwerken.load_local(directory))
    atomic_write(directory / 'index.html', bijwerken.render_html(directory, data))


def commit_local(directory, local, before_data, before_local, action, undo_of=None):
    merged = bijwerken.merge_local(copy.deepcopy(before_data), copy.deepcopy(local))
    beheer.log_changes(local, before_data, before_local, merged, action, undo_of)
    merged = bijwerken.merge_local(copy.deepcopy(before_data), copy.deepcopy(local))
    html = bijwerken.render_html(directory, merged)
    atomic_write(directory / bijwerken.LOCAL_FILE, json.dumps(local, ensure_ascii=False, indent=2) + '\n')
    atomic_write(directory / 'index.html', html)


def image_extension(raw):
    if raw.startswith(b'\xff\xd8\xff'):
        return '.jpg'
    if raw.startswith(b'\x89PNG\r\n\x1a\n'):
        return '.png'
    if raw.startswith((b'GIF87a', b'GIF89a')):
        return '.gif'
    if raw[:4] == b'RIFF' and raw[8:12] == b'WEBP':
        return '.webp'
    raise ValueError('Gebruik JPG-, PNG-, WebP- of GIF-foto’s.')


def text(value, maximum=10000):
    if not isinstance(value, str) or len(value) > maximum:
        raise ValueError('Ongeldige of te lange tekst.')
    return value.strip()


def normalize_values(values, key, *, require_identity=True):
    if not isinstance(values, list) or len(values) != 15:
        raise ValueError('Alle 15 inventarisvelden zijn nodig.')
    result = [key] + [text(v) or 'Onbekend' for v in values[1:]]
    if require_identity and any(result[i] == 'Onbekend' for i in (1, 2)):
        raise ValueError('Vul minstens automerk en model in.')
    for i in (3, 8):
        if result[i].isdigit():
            result[i] = int(result[i])
    for i in (11, 12, 13):
        if result[i] != 'Onbekend':
            try:
                value = float(result[i].replace(',', '.'))
            except ValueError as exc:
                raise ValueError('Vul een geldige prijs in, bijvoorbeeld 75,50.') from exc
            if not math.isfinite(value) or value < 0:
                raise ValueError('Prijzen moeten nul of hoger zijn.')
            result[i] = value
    return result


class DuplicateModelError(ValueError):
    def __init__(self, matches):
        super().__init__('Er zijn mogelijk dubbele modellen. Controleer de bestaande exemplaren en bevestig indien je toch wilt opslaan.')
        self.matches = matches


def save_sale(local, key, record_values, origin, metadata):
    if str(record_values[14]).casefold() != 'verkocht':
        return
    current = local.get('sales', {}).get(key)
    if 'sale' not in metadata and current:
        return
    item = metadata.get('sale') or {}
    if not isinstance(item, dict):
        raise ValueError('Ongeldige verkoopgegevens.')
    date = text(item.get('date', datetime.datetime.now(ZoneInfo('Europe/Brussels')).date().isoformat()), 10)
    try:
        datetime.date.fromisoformat(date)
    except ValueError as exc:
        raise ValueError('Kies een geldige verkoopdatum.') from exc
    price = item.get('price')
    if price in ('', None):
        price = None
    else:
        price = float(str(price).replace(',', '.'))
        if not math.isfinite(price) or price < 0:
            raise ValueError('De verkoopprijs moet nul of hoger zijn.')
    local.setdefault('sales', {})[key] = {'identity': None if origin == 'local' else record_values[1:6], 'price': price, 'date': date}


def save_model(directory, metadata, uploads):
    """Roep aan onder LOCK. Bewaar lokale modellen en media voor Excel apart."""
    local = bijwerken.load_local(directory)
    before_local = copy.deepcopy(local)
    data = bijwerken.merge_local(bijwerken.snapshot(directory), copy.deepcopy(local))
    key = text(metadata.get('id', ''), 100)
    existing = next((r for r in data['records'] if str(r['values'][0]) == key), None)
    if key and existing is None:
        raise ValueError('Dit model bestaat niet meer. Vernieuw de pagina.')
    if existing and metadata.get('identity') != existing['values'][1:6]:
        raise ValueError('De modelgegevens zijn intussen gewijzigd. Vernieuw de pagina voordat je opslaat.')
    request_id = text(metadata.get('requestId', ''), 100)
    if not key and request_id:
        previous = next((r for r in local['models'] if r.get('requestId') == request_id), None)
        if previous:
            return str(previous['values'][0])
    if not key:
        key = 'W-' + uuid.uuid4().hex[:12]
    origin = existing.get('origin', 'excel') if existing else 'local'
    values = normalize_values(metadata.get('values'), key, require_identity=origin == 'local')
    note = text(metadata.get('note', ''))
    if existing:
        values[0] = existing['values'][0]
        # Een ongewijzigd invoerveld behoudt het originele Excel-type en de tekst.
        for index in range(1, 15):
            if metadata['values'][index] == str(existing['values'][index]):
                values[index] = existing['values'][index]
        if existing.get('bundle') and (values[14] != existing['values'][14] or 'sale' in metadata):
            raise ValueError('Beheer status en verkoopgegevens van dit model via de verkoopbundel.')
    matches = collectie.duplicates(data['records'], values, key)
    if matches and metadata.get('confirmDuplicate') is not True:
        raise DuplicateModelError(matches)
    ad = text(metadata.get('ad', ''), 100000)
    photos = metadata.get('photos', [])
    if not isinstance(photos, list) or len(photos) > 20:
        raise ValueError('Je kunt maximaal 20 foto’s per model bewaren.')
    allowed = set((existing or {}).get('extra', {}).get('photos', []))
    prepared = []
    for photo in photos:
        if not isinstance(photo, dict):
            raise ValueError('Ongeldige fotolijst.')
        if 'path' in photo:
            path = photo['path']
            if not isinstance(path, str) or path not in allowed:
                raise ValueError('Een bestaande foto hoort niet bij dit model.')
            prepared.append((path, None))
        else:
            upload = uploads.get(photo.get('upload'))
            if not upload or len(upload) > 15 * 1024 * 1024:
                raise ValueError('Een foto ontbreekt of is groter dan 15 MB.')
            extension = image_extension(upload)
            prepared.append(('fotos/web/' + uuid.uuid4().hex + extension, upload))
    extra = {'photos': [p for p, _ in prepared], 'ad': ad}
    # ChatGPT URL: ontbreekt het veld in het verzoek, dan blijft de bestaande waarde behouden.
    if 'chatUrl' in metadata:
        chat_url = text(metadata.get('chatUrl'), 2000)
        if chat_url and not re.fullmatch(r'https://\S+', chat_url):
            raise ValueError('Vul een geldige ChatGPT URL in die begint met https://')
    else:
        chat_url = (existing or {}).get('extra', {}).get('chatUrl', '')
    if chat_url:
        extra['chatUrl'] = chat_url
    if origin == 'local':
        record = {'origin': 'local', 'values': values, 'note': note, 'extra': extra, 'requestId': (existing or {}).get('requestId', request_id)}
        local['models'] = [r for r in local['models'] if str(r['values'][0]) != key] + [record]
    else:
        baseline = existing.get('originalValues', existing['values'])
        identity = baseline[1:6]
        local['excel_media'][key] = {'identity': identity, 'extra': extra}
        changes = {str(i): values[i] for i in range(1, 15) if values[i] != baseline[i]}
        edits = {'identity': identity, 'values': changes}
        if note != existing.get('originalNote', existing.get('note', '')):
            edits['note'] = note
        local.setdefault('excel_edits', {}).pop(key, None)
        if changes or 'note' in edits:
            local['excel_edits'][key] = edits
    baseline = existing.get('originalValues', existing['values']) if existing else values
    sale_identity = baseline[:]
    sale_identity[14] = values[14]
    save_sale(local, key, sale_identity, origin, metadata)
    written = []
    committed = False
    try:
        for path, raw in prepared:
            if raw is not None:
                target = directory / path
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(raw)
                written.append(target)
        commit_local(directory, local, data, before_local, 'Model bewerkt' if existing else 'Model toegevoegd')
        committed = True
    finally:
        # Na JSON-opslag is de informatie duurzaam; herstart herbouwt eventueel de HTML.
        durable = (directory / bijwerken.LOCAL_FILE).exists() and bijwerken.load_local(directory) == local
        if not committed and not durable:
            for path in written:
                path.unlink(missing_ok=True)
    return key


def change_model_state(directory, metadata, *, restore=False):
    """Verplaats naar de lokale prullenbak of herstel, zonder bronbestanden te wissen."""
    key = text(metadata.get('id', ''), 100)
    local = bijwerken.load_local(directory)
    before_local = copy.deepcopy(local)
    data = bijwerken.merge_local(bijwerken.snapshot(directory), copy.deepcopy(local))
    if restore:
        if metadata.get('origin') == 'local':
            record = next((r for r in local.get('archived_models', []) if str(r['values'][0]) == key), None)
            if not record:
                raise ValueError('Dit model staat niet meer in de prullenbak.')
            local['archived_models'] = [r for r in local['archived_models'] if str(r['values'][0]) != key]
            local['models'].append(record)
        elif metadata.get('origin') == 'excel':
            if key not in local.get('hidden_excel', {}):
                raise ValueError('Dit model staat niet meer in de prullenbak.')
            del local['hidden_excel'][key]
        else:
            raise ValueError('Ongeldige modelbron.')
    else:
        record = next((r for r in data['records'] if str(r['values'][0]) == key), None)
        if not record or metadata.get('identity') != record['values'][1:6]:
            raise ValueError('Het model is intussen gewijzigd of verwijderd. Vernieuw de pagina.')
        if record.get('bundle'):
            raise ValueError('Haal dit model eerst uit de verkoopbundel voordat je het verwijdert.')
        if record.get('origin') == 'local':
            local['models'] = [r for r in local['models'] if str(r['values'][0]) != key]
            record = copy.deepcopy(record)
            record.pop('appointment', None)
            record.pop('sale', None)
            local.setdefault('archived_models', []).append(record)
        else:
            baseline = record.get('originalValues', record['values'])
            local.setdefault('hidden_excel', {})[key] = {'identity': baseline[1:6], 'title': str(record['values'][1]) + ' ' + str(record['values'][2])}
    commit_local(directory, local, data, before_local, 'Hersteld uit prullenbak' if restore else 'Naar prullenbak verplaatst')
    return key


def save_appointment(directory, metadata):
    """Bewaar of verwijder de afspraak met een koper voor één model."""
    key = text(metadata.get('id', ''), 100)
    local = bijwerken.load_local(directory)
    before_local = copy.deepcopy(local)
    data = bijwerken.merge_local(bijwerken.snapshot(directory), copy.deepcopy(local))
    record = next((r for r in data['records'] if str(r['values'][0]) == key), None)
    if not record or metadata.get('identity') != record['values'][1:6]:
        raise ValueError('Het model is intussen gewijzigd of verwijderd. Vernieuw de pagina.')
    if record.get('bundle'):
        raise ValueError('Beheer de afhaalafspraak via de verkoopbundel.')
    appointments = local.setdefault('appointments', {})
    item = metadata.get('appointment')
    if item is None:
        appointments.pop(key, None)
    else:
        if not isinstance(item, dict):
            raise ValueError('Ongeldige afspraak.')
        start = text(item.get('start', ''), 20)
        try:
            datetime.datetime.strptime(start, '%Y-%m-%dT%H:%M')
        except ValueError as exc:
            raise ValueError('Kies een geldige datum en tijd.') from exc
        duration = item.get('duration', 30)
        if not isinstance(duration, int) or isinstance(duration, bool) or not 5 <= duration <= 600:
            raise ValueError('Kies een duur tussen 5 minuten en 10 uur.')
        identity = None if record.get('origin') == 'local' else record.get('originalValues', record['values'])[1:6]
        appointments[key] = {'identity': identity, 'start': start, 'duration': duration, 'buyer': text(item.get('buyer', ''), 500), 'note': text(item.get('note', ''), 2000)}
    if not appointments:
        local.pop('appointments')
    commit_local(directory, local, data, before_local, 'Afspraak gewijzigd')
    return key


def save_status(directory, metadata):
    data = bijwerken.merge_local(bijwerken.snapshot(directory), bijwerken.load_local(directory))
    record = next((r for r in data['records'] if str(r['values'][0]) == metadata.get('id')), None)
    if not record:
        raise ValueError('Dit model bestaat niet meer.')
    choices = {v.casefold(): v for v in ('Nog te plaatsen', 'Online', 'Gereserveerd', 'Verkocht', 'Onbekend')}
    value = text(metadata.get('status', ''), 100) or 'Onbekend'
    if value.casefold() not in choices:
        raise ValueError('Kies een geldige verkoopstatus.')
    values = [str(v) for v in record['values']]
    values[14] = choices[value.casefold()]
    request = {'id': metadata['id'], 'identity': metadata.get('identity'), 'values': values, 'note': record.get('note', ''), 'ad': record.get('extra', {}).get('ad', ''), 'photos': [{'path': p} for p in record.get('extra', {}).get('photos', [])], 'confirmDuplicate': True}
    if 'sale' in metadata:
        request['sale'] = metadata['sale']
    return save_model(directory, request, {})


def save_bundle(directory, metadata):
    local = bijwerken.load_local(directory)
    before_local = copy.deepcopy(local)
    data = bijwerken.merge_local(bijwerken.snapshot(directory), copy.deepcopy(local))
    bundles = local.setdefault('bundles', {})
    key = text(metadata.get('id', ''), 100)
    old = bundles.get(key)
    if key and not old:
        raise ValueError('Deze verkoopbundel bestaat niet meer.')
    if not key:
        request_id = text(metadata.get('requestId', ''), 100)
        previous = next((b for b in bundles.values() if request_id and b.get('requestId') == request_id), None)
        if previous:
            return previous['id']
        key = 'B-' + uuid.uuid4().hex[:12]
    records = {str(r['values'][0]): r for r in data['records']}
    cancel = metadata.get('action') == 'cancel'
    if cancel and not old:
        raise ValueError('De verkoopbundel ontbreekt.')
    if cancel and old['status'] == 'Geannuleerd':
        return key
    selected = [] if cancel else metadata.get('members', [])
    if not isinstance(selected, list) or (not cancel and len(selected) < 2):
        raise ValueError('Selecteer minstens twee modellen voor de bundel.')
    chosen = []
    seen = set()
    for item in selected:
        model_id = text(item.get('id', ''), 100)
        record = records.get(model_id)
        if model_id in seen or not record or item.get('identity') != record['values'][1:6]:
            raise ValueError('Een geselecteerd model is dubbel, ontbreekt of is intussen gewijzigd. Vernieuw de pagina.')
        seen.add(model_id)
        if record.get('bundle', {}).get('id') not in (None, key):
            raise ValueError('Model ' + model_id + ' hoort al bij een andere verkoopbundel.')
        if str(record['values'][14]).casefold() == 'verkocht' and not (record.get('bundle', {}).get('id') == key):
            raise ValueError('Model ' + model_id + ' is al afzonderlijk verkocht.')
        chosen.append(record)
    # Bestaande leden die verdwijnen, krijgen hun eerdere status terug.
    prior = copy.deepcopy((old or {}).get('prior', {})) if (old or {}).get('status') != 'Geannuleerd' else {}
    for member in (old or {}).get('members', []) if (old or {}).get('status') != 'Geannuleerd' else []:
        record = records.get(member['id'])
        if member['id'] not in seen and record and beheer.identity(record) == member['identity']:
            values = record['values'][:]
            values[14] = prior.get(member['id'], {}).get('status', 'Onbekend')
            beheer.put_fields(local, record, values)
    if cancel:
        bundle = {**old, 'status': 'Geannuleerd'}
    else:
        buyer = text(metadata.get('buyer', ''), 500)
        if not buyer:
            raise ValueError('Vul de naam van de koper in.')
        status = text(metadata.get('status', 'Gereserveerd'), 50)
        if status not in ('Gereserveerd', 'Verkocht'):
            raise ValueError('Kies Gereserveerd of Verkocht voor de bundel.')
        total = metadata.get('total')
        if total in ('', None):
            total = None
        else:
            total = float(str(total).replace(',', '.'))
            if not math.isfinite(total) or total < 0:
                raise ValueError('Vul een geldig bundeltotaal in.')
        date = text(metadata.get('date') or datetime.datetime.now(ZoneInfo('Europe/Brussels')).date().isoformat(), 10)
        datetime.date.fromisoformat(date)
        appointment = metadata.get('appointment')
        if appointment is not None:
            if not isinstance(appointment, dict):
                raise ValueError('Ongeldige afhaalafspraak.')
            start = text(appointment.get('start', ''), 20)
            datetime.datetime.strptime(start, '%Y-%m-%dT%H:%M')
            duration = appointment.get('duration', 30)
            if type(duration) is not int or not 5 <= duration <= 600:
                raise ValueError('De afspraakduur moet tussen 5 en 600 minuten liggen.')
            appointment = {'start': start, 'duration': duration}
        members = []
        for record in chosen:
            model_id = str(record['values'][0])
            previous_member = next((m for m in (old or {}).get('members', []) if m['id'] == model_id), None)
            if previous_member and previous_member['identity'] != beheer.identity(record):
                prior[model_id] = {'status': record['values'][14]}
            else:
                prior.setdefault(model_id, {'status': record['values'][14]})
            members.append({'id': model_id, 'identity': beheer.identity(record), 'title': str(record['values'][1]) + ' ' + str(record['values'][2])})
            values = record['values'][:]
            values[14] = status
            beheer.put_fields(local, record, values)
        prior = {model_id: prior[model_id] for model_id in seen}
        bundle = {'id': key, 'requestId': (old or {}).get('requestId', metadata.get('requestId', '')), 'name': text(metadata.get('name', ''), 500) or 'Bundel voor ' + buyer, 'buyer': buyer, 'contact': text(metadata.get('contact', ''), 500), 'note': text(metadata.get('note', ''), 2000), 'status': status, 'total': total, 'date': date if status == 'Verkocht' else None, 'members': members, 'appointment': appointment, 'prior': prior}
    bundles[key] = bundle
    commit_local(directory, local, data, before_local, 'Verkoopbundel geannuleerd' if cancel else 'Verkoopbundel gewijzigd' if old else 'Verkoopbundel gemaakt')
    return key


def undo_history(directory, metadata):
    local = bijwerken.load_local(directory)
    before_local = copy.deepcopy(local)
    data = bijwerken.merge_local(bijwerken.snapshot(directory), copy.deepcopy(local))
    key = text(metadata.get('id', ''), 100)
    event_id = text(metadata.get('eventId', ''), 100)
    event = beheer.event_for(local, key, event_id)
    for change in event['changes']:
        if change['field'] == 'photos':
            for path in change['before']:
                if not collectie.safe_photo(directory, path).is_file():
                    raise ValueError('Een oudere foto ontbreekt. Herstel die eerst via Back-ups.')
    beheer.undo_event(local, data, key, event_id)
    commit_local(directory, local, data, before_local, 'Wijziging teruggedraaid', undo_of=event_id)
    return key


class Handler(BaseHTTPRequestHandler):
    def send_json(self, status, value):
        self.send_bytes(status, json.dumps(value, ensure_ascii=False).encode('utf-8'), 'application/json; charset=utf-8')

    def send_bytes(self, status, raw, content_type):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(raw)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Cross-Origin-Resource-Policy', 'same-origin')
        if content_type in ('application/zip', 'application/vnd.apache.parquet'):
            self.send_header('Content-Disposition', 'attachment; filename="collectie.' + ('zip' if content_type == 'application/zip' else 'parquet') + '"')
        self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' blob:; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'")
        self.end_headers()
        self.wfile.write(raw)

    def valid_host(self):
        return self.headers.get('Host') == f'127.0.0.1:{self.server.server_port}'

    def do_GET(self):
        if not self.valid_host():
            self.send_json(403, {'error': 'Open de app via het lokale adres 127.0.0.1.'})
            return
        route = unquote(urlsplit(self.path).path)
        directory = self.server.directory
        if route == '/api/config':
            with LOCK:
                self.send_json(200, {'token': self.server.token, 'revision': revision(directory), 'version': 3})
            return
        if route == '/api/history':
            from urllib.parse import parse_qs
            key = parse_qs(urlsplit(self.path).query).get('id', [''])[0]
            with LOCK:
                local = bijwerken.load_local(directory)
                self.send_json(200, {'events': list(reversed(local.get('history', {}).get(key, [])))})
            return
        if route == '/api/backups':
            with LOCK:
                self.send_json(200, {'backups': collectie.list_backups(directory)})
            return
        if route in ('/api/export/collectie.zip', '/api/export/collectie.parquet'):
            with LOCK:
                data = bijwerken.merge_local(bijwerken.snapshot(directory), bijwerken.load_local(directory))
                collectie.enrich(data, directory)
                raw = collectie.export_zip(directory, data) if route.endswith('.zip') else collectie.parquet_bytes(data, bijwerken.load_local(directory))
            self.send_bytes(200, raw, 'application/zip' if route.endswith('.zip') else 'application/vnd.apache.parquet')
            return
        if route in ('/', '/index.html'):
            path = directory / 'index.html'
            content_type = 'text/html; charset=utf-8'
        elif route.startswith('/fotos/'):
            path = (directory / route.lstrip('/')).resolve()
            if not path.is_relative_to((directory / 'fotos').resolve()):
                self.send_json(403, {'error': 'Ongeldig fotopad.'})
                return
            content_type = {'.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.webp': 'image/webp', '.gif': 'image/gif'}.get(path.suffix.lower())
            if not content_type:
                self.send_json(404, {'error': 'Foto niet gevonden.'})
                return
        else:
            self.send_json(404, {'error': 'Niet gevonden.'})
            return
        try:
            self.send_bytes(200, path.read_bytes(), content_type)
        except FileNotFoundError:
            self.send_json(404, {'error': 'Bestand niet gevonden.'})

    def do_POST(self):
        origin = f'http://127.0.0.1:{self.server.server_port}'
        if not self.valid_host() or self.headers.get('Origin') != origin or not secrets.compare_digest(self.headers.get('X-Inventaris-Token', ''), self.server.token):
            # Lees kleine afgewezen verzoeken leeg, anders kan Windows de 403 met een TCP-reset afbreken.
            try:
                rejected_length = int(self.headers.get('Content-Length', '0'))
                if 0 < rejected_length <= 1024 * 1024:
                    self.rfile.read(rejected_length)
            except ValueError:
                pass
            self.send_json(403, {'error': 'Open de app via Start inventaris.cmd en vernieuw de pagina.'})
            return
        route = urlsplit(self.path).path
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if length < 0 or length > MAX_BODY:
                self.send_json(413, {'error': 'De totale upload mag maximaal 80 MB zijn.'})
                return
            raw = self.rfile.read(length)
            with LOCK:
                if self.headers.get('X-Inventaris-Revision') != revision(self.server.directory):
                    self.send_json(409, {'error': 'De collectie is intussen in een ander venster gewijzigd. Vernieuw de pagina voordat je opslaat.'})
                    return
                if route == '/api/refresh':
                    bijwerken.main(self.server.directory)
                    self.send_json(200, {'ok': True})
                elif route in ('/api/backups/create', '/api/backups/restore'):
                    metadata = json.loads(raw.decode('utf-8'))
                    if route.endswith('/create'):
                        backup = collectie.create_backup(self.server.directory, 'Handmatig herstelpunt')
                        self.send_json(200, {'ok': True, 'backup': backup})
                    else:
                        missing = collectie.restore_backup(self.server.directory, metadata.get('id'))
                        rebuild(self.server.directory)
                        self.send_json(200, {'ok': True, 'missing': missing})
                elif route == '/api/status':
                    metadata = json.loads(raw.decode('utf-8'))
                    key = save_status(self.server.directory, metadata)
                    self.send_json(200, {'ok': True, 'id': key})
                elif route in ('/api/bundles', '/api/history/undo'):
                    metadata = json.loads(raw.decode('utf-8'))
                    if not isinstance(metadata, dict):
                        raise ValueError('Ongeldige gegevens.')
                    key = save_bundle(self.server.directory, metadata) if route == '/api/bundles' else undo_history(self.server.directory, metadata)
                    self.send_json(200, {'ok': True, 'id': key})
                elif route in ('/api/models/delete', '/api/models/restore'):
                    metadata = json.loads(raw.decode('utf-8'))
                    if not isinstance(metadata, dict):
                        raise ValueError('Ongeldige modelgegevens.')
                    key = change_model_state(self.server.directory, metadata, restore=route.endswith('/restore'))
                    self.send_json(200, {'ok': True, 'id': key})
                elif route == '/api/appointments':
                    metadata = json.loads(raw.decode('utf-8'))
                    if not isinstance(metadata, dict):
                        raise ValueError('Ongeldige afspraak.')
                    key = save_appointment(self.server.directory, metadata)
                    self.send_json(200, {'ok': True, 'id': key})
                elif route == '/api/models':
                    content_type = self.headers.get('Content-Type', '')
                    if not content_type.startswith('multipart/form-data;'):
                        raise ValueError('Ongeldige upload.')
                    message = BytesParser(policy=policy.default).parsebytes(b'Content-Type: ' + content_type.encode('ascii') + b'\r\nMIME-Version: 1.0\r\n\r\n' + raw)
                    if not message.is_multipart():
                        raise ValueError('Ongeldige upload.')
                    parts = {}
                    for part in message.iter_parts():
                        name = part.get_param('name', header='content-disposition')
                        if name in parts:
                            raise ValueError('Dubbel uploadveld.')
                        parts[name] = part.get_payload(decode=True)
                    metadata = json.loads(parts.pop('metadata', b'{}').decode('utf-8'))
                    if not isinstance(metadata, dict):
                        raise ValueError('Ongeldige modelgegevens.')
                    key = save_model(self.server.directory, metadata, parts)
                    make_thumbnails(self.server.directory)
                    self.send_json(200, {'ok': True, 'id': key})
                else:
                    self.send_json(404, {'error': 'Niet gevonden.'})
        except DuplicateModelError as exc:
            self.send_json(409, {'error': str(exc), 'duplicates': exc.matches})
        except (ValueError, TypeError, KeyError) as exc:
            self.send_json(400, {'error': str(exc)})
        except PermissionError:
            self.send_json(409, {'error': 'Een bestand is vergrendeld. Sla Excel op en sluit het bestand als je Excel opnieuw wilt inlezen.'})
        except Exception as exc:
            self.log_error('%s', exc)
            self.send_json(500, {'error': 'Opslaan of bijwerken is niet gelukt. Controleer het appvenster; je invoer blijft in het formulier staan.'})


class LocalHTTPServer(ThreadingHTTPServer):
    # Windows laat met SO_REUSEADDR twee apps dezelfde poort binden.
    # Exclusief binden voorkomt dat de browser bij een andere app terechtkomt.
    allow_reuse_address = False

    def server_bind(self):
        if hasattr(socket, 'SO_EXCLUSIVEADDRUSE'):
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


def make_thumbnails(directory):
    """Miniaturen voor het foto-overzicht bijwerken; een fout mag opslaan niet hinderen."""
    try:
        with THUMB_LOCK:
            miniaturen.ensure_for_data(directory, bijwerken.snapshot(directory))
    except Exception as exc:
        print(f'Miniaturen bijwerken mislukt: {exc}', file=sys.stderr, flush=True)


def create_server(directory=DIRECTORY, port=DEFAULT_PORT):
    server = LocalHTTPServer(('127.0.0.1', port), Handler)
    server.directory = Path(directory).resolve()
    server.token = secrets.token_urlsafe(32)
    server.daemon_threads = True
    return server


def start_server(directory=DIRECTORY, port=None):
    if port is not None:
        return create_server(directory, port)
    try:
        return create_server(directory, DEFAULT_PORT)
    except OSError:
        # De OS kiest en reserveert een vrije poort in dezelfde bind-operatie.
        return create_server(directory, 0)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--port', type=int, default=None)
    parser.add_argument('--no-browser', action='store_true')
    args = parser.parse_args()
    if not DIRECTORY.is_dir():
        raise SystemExit(f'Gegevensmap niet gevonden: {DIRECTORY}\nMaak de koppeling naar OneDrive opnieuw aan (zie LEESMIJ.txt).')
    rebuild(DIRECTORY)
    # Ontbrekende miniaturen op de achtergrond aanmaken; de pagina toont intussen de originelen.
    threading.Thread(target=make_thumbnails, args=(DIRECTORY,), daemon=True).start()
    try:
        server = start_server(port=args.port)
    except OSError as exc:
        raise SystemExit('De lokale poort is niet beschikbaar. Start zonder --port om automatisch een vrije poort te kiezen.') from exc
    url = f'http://127.0.0.1:{server.server_port}/'
    print(f'Inventaris: {url}\nLaat dit venster open tijdens het werken. Stop met Ctrl+C.', flush=True)
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == '__main__':
    main()

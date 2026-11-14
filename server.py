"""Lokale inventarisapp: Excel blijft leesbron, browsertoevoegingen staan apart."""
import argparse
import copy
import datetime
from email import policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import math
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
    temp = path.with_name(path.name + '.tmp')
    try:
        temp.write_text(content, encoding='utf-8')
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


def rebuild(directory):
    data = bijwerken.merge_local(bijwerken.snapshot(directory), bijwerken.load_local(directory))
    atomic_write(directory / 'index.html', bijwerken.render_html(directory, data))


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


def save_model(directory, metadata, uploads):
    """Roep aan onder LOCK. Bewaar lokale modellen en media voor Excel apart."""
    local = bijwerken.load_local(directory)
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
    merged = bijwerken.merge_local(data, copy.deepcopy(local))
    html = bijwerken.render_html(directory, merged)
    written = []
    committed = False
    try:
        for path, raw in prepared:
            if raw is not None:
                target = directory / path
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(raw)
                written.append(target)
        atomic_write(directory / bijwerken.LOCAL_FILE, json.dumps(local, ensure_ascii=False, indent=2) + '\n')
        committed = True
        atomic_write(directory / 'index.html', html)
    finally:
        # Na JSON-opslag is de informatie duurzaam; herstart herbouwt eventueel de HTML.
        if not committed:
            for path in written:
                path.unlink(missing_ok=True)
    return key


def change_model_state(directory, metadata, *, restore=False):
    """Verplaats naar de lokale prullenbak of herstel, zonder bronbestanden te wissen."""
    key = text(metadata.get('id', ''), 100)
    local = bijwerken.load_local(directory)
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
        if record.get('origin') == 'local':
            local['models'] = [r for r in local['models'] if str(r['values'][0]) != key]
            record.pop('appointment', None)
            local.setdefault('archived_models', []).append(record)
        else:
            baseline = record.get('originalValues', record['values'])
            local.setdefault('hidden_excel', {})[key] = {'identity': baseline[1:6], 'title': str(record['values'][1]) + ' ' + str(record['values'][2])}
    merged = bijwerken.merge_local(data, copy.deepcopy(local))
    html = bijwerken.render_html(directory, merged)
    atomic_write(directory / bijwerken.LOCAL_FILE, json.dumps(local, ensure_ascii=False, indent=2) + '\n')
    atomic_write(directory / 'index.html', html)
    return key


def save_appointment(directory, metadata):
    """Bewaar of verwijder de afspraak met een koper voor één model."""
    key = text(metadata.get('id', ''), 100)
    local = bijwerken.load_local(directory)
    data = bijwerken.merge_local(bijwerken.snapshot(directory), copy.deepcopy(local))
    record = next((r for r in data['records'] if str(r['values'][0]) == key), None)
    if not record or metadata.get('identity') != record['values'][1:6]:
        raise ValueError('Het model is intussen gewijzigd of verwijderd. Vernieuw de pagina.')
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
    merged = bijwerken.merge_local(data, copy.deepcopy(local))
    html = bijwerken.render_html(directory, merged)
    atomic_write(directory / bijwerken.LOCAL_FILE, json.dumps(local, ensure_ascii=False, indent=2) + '\n')
    atomic_write(directory / 'index.html', html)
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
                self.send_json(200, {'token': self.server.token, 'revision': revision(directory)})
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

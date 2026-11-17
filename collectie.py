"""Volledige export, herstelpunten en controles voor de gecombineerde collectie."""
import datetime as dt
import hashlib
import io
import json
from pathlib import Path
import re
import unicodedata
import uuid
import zipfile

FILES = ('index.html', 'lokale_modellen.json', 'aanvullingen.json', 'inventaris.parquet')


def now():
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec='seconds')


def known(value):
    return value is not None and str(value).strip().casefold() not in ('', 'onbekend')


def normalized(value):
    return re.sub(r'[^a-z0-9]', '', ''.join(c for c in unicodedata.normalize('NFKD', str(value).casefold()) if not unicodedata.combining(c))) if known(value) else ''


def duplicates(records, values, exclude=''):
    if not all(known(values[i]) for i in (1, 2)):
        return []
    result = []
    for record in records:
        other = record['values']
        if str(other[0]) == exclude or any(normalized(values[i]) != normalized(other[i]) for i in (1, 2)):
            continue
        if any(known(values[i]) and known(other[i]) and normalized(values[i]) != normalized(other[i]) for i in (3, 4, 5, 7)):
            continue
        result.append({'id': str(other[0]), 'title': str(other[1]) + ' ' + str(other[2])})
    return result


def ad_prices(ad):
    # Alleen expliciete vraagprijzen/eurobedragen; bouwjaren en artikelnummers tellen niet mee.
    pattern = r'(?:vraagprijs\s*:?\s*(?:€\s*|EUR\s*)?|€\s*)(\d+(?:[. ]\d{3})*(?:[,.]\d{1,2})?)(?!\d)'
    result = []
    for match in re.finditer(pattern, ad.replace('**', ''), re.I):
        value = match.group(1).replace(' ', '')
        if ',' in value:
            value = value.replace('.', '').replace(',', '.')
        elif re.fullmatch(r'\d{1,3}(?:\.\d{3})+', value):
            value = value.replace('.', '')
        result.append(float(value))
    return result


def checks(record, directory=None):
    extra = record.get('extra') or {}
    prices = [p for p in record['values'][11:14] if type(p) in (int, float)]
    issues = []
    if not extra.get('photos'):
        issues.append({'code': 'photos', 'message': 'Foto’s ontbreken'})
    elif directory is not None:
        missing = sum(not safe_photo(Path(directory), p).is_file() for p in extra['photos'])
        if missing:
            issues.append({'code': 'photos', 'message': str(missing) + ' fotobestand(en) niet gevonden'})
    if not extra.get('ad', '').strip():
        issues.append({'code': 'ad', 'message': 'Zoekertjestekst ontbreekt'})
    if not prices:
        issues.append({'code': 'price', 'message': 'Vraagprijs ontbreekt'})
    stated = ad_prices(extra.get('ad', ''))
    if stated and prices and not any(abs(a - b) < .005 for a in stated for b in prices):
        issues.append({'code': 'price_mismatch', 'message': 'Prijs in zoekertjestekst wijkt af van de inventaris'})
    return issues


def enrich(data, directory=None):
    for record in data['records']:
        record['quality'] = {'issues': checks(record, directory), 'duplicates': duplicates(data['records'], record['values'], str(record['values'][0]))}
    return data


def photo_paths(value):
    if isinstance(value, dict):
        for key, item in value.items():
            if key == 'photos' and isinstance(item, list):
                yield from (p for p in item if isinstance(p, str))
            else:
                yield from photo_paths(item)
    elif isinstance(value, list):
        for item in value:
            yield from photo_paths(item)


def safe_photo(directory, relative):
    if not relative.startswith('fotos/') or relative.startswith('fotos/_mini/'):
        raise ValueError('Ongeldig fotopad.')
    path = (directory / relative).resolve()
    if not path.is_relative_to((directory / 'fotos').resolve()):
        raise ValueError('Ongeldig fotopad.')
    return path


def inventory_files(directory):
    return {name: (directory / name).read_bytes() for name in FILES if (directory / name).is_file()}


def referenced_photos(files):
    values = []
    for name, raw in files.items():
        if name.endswith('.json'):
            values.append(json.loads(raw))
        elif name == 'index.html':
            match = re.search(r'<script id="inventory-data" type="application/json">(.*?)</script>', raw.decode('utf-8'), re.S)
            if match:
                values.append(json.loads(match.group(1)))
    return sorted(set(photo_paths(values)))


def atomic_bytes(path, raw):
    temp = path.with_name(path.name + '.tmp')
    try:
        temp.write_bytes(raw)
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


def create_backup(directory, reason):
    directory = Path(directory)
    files = inventory_files(directory)
    if 'index.html' not in files:
        return None
    root = directory / 'backups'
    pool = root / 'fotobestanden'
    pool.mkdir(parents=True, exist_ok=True)
    cache_path = root / 'foto-cache.json'
    cache = json.loads(cache_path.read_text(encoding='utf-8')) if cache_path.exists() else {}
    manifest = {'version': 1, 'created': now(), 'reason': reason, 'photos': {}, 'missing': []}
    for relative in referenced_photos(files):
        source = safe_photo(directory, relative)
        if not source.is_file():
            manifest['missing'].append(relative)
            continue
        stat = source.stat()
        cached = cache.get(relative, {})
        if cached.get('size') == stat.st_size and cached.get('mtime') == stat.st_mtime_ns and re.fullmatch(r'[a-f0-9]{64}', cached.get('sha', '')) and (pool / cached['sha']).is_file():
            digest = cached['sha']
        else:
            raw = source.read_bytes()
            digest = hashlib.sha256(raw).hexdigest()
            if not (pool / digest).exists():
                atomic_bytes(pool / digest, raw)
            cache[relative] = {'size': stat.st_size, 'mtime': stat.st_mtime_ns, 'sha': digest}
        manifest['photos'][relative] = digest
    atomic_bytes(cache_path, json.dumps(cache).encode('utf-8'))
    identifier = dt.datetime.now(dt.timezone.utc).strftime('%Y%m%dT%H%M%S%fZ') + '-' + uuid.uuid4().hex[:8] + '.zip'
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        for name, raw in files.items():
            archive.writestr(name, raw)
        archive.writestr('manifest.json', json.dumps(manifest, ensure_ascii=False))
    atomic_bytes(root / identifier, buffer.getvalue())
    return {'id': identifier, **manifest}


def list_backups(directory):
    result = []
    for path in sorted((Path(directory) / 'backups').glob('*.zip'), reverse=True):
        with zipfile.ZipFile(path) as archive:
            item = json.loads(archive.read('manifest.json'))
        result.append({'id': path.name, 'created': item['created'], 'reason': item['reason'], 'photos': len(item['photos']), 'missing': len(item['missing'])})
    return result


def restore_backup(directory, identifier):
    directory = Path(directory)
    if not isinstance(identifier, str) or not re.fullmatch(r'\d{8}T\d{12}Z-[a-f0-9]{8}\.zip', identifier):
        raise ValueError('Ongeldig herstelpunt.')
    root = directory / 'backups'
    with zipfile.ZipFile(root / identifier) as archive:
        manifest = json.loads(archive.read('manifest.json'))
        files = {name: archive.read(name) for name in FILES if name in archive.namelist()}
    if 'index.html' not in files:
        raise ValueError('Herstelpunt bevat geen inventarispagina.')
    photos = []
    for relative, digest in manifest['photos'].items():
        target = safe_photo(directory, relative)
        if not re.fullmatch(r'[a-f0-9]{64}', digest):
            raise ValueError('Ongeldige foto in het herstelpunt.')
        raw = (root / 'fotobestanden' / digest).read_bytes()
        if hashlib.sha256(raw).hexdigest() != digest:
            raise ValueError('Een back-upfoto is beschadigd; herstel is niet uitgevoerd.')
        photos.append((target, raw))
    # Een eerste herstelpunt kan van vóór het eerste paginamodel zijn.
    files.setdefault('lokale_modellen.json', b'{"models": [], "excel_media": {}}')
    files.setdefault('aanvullingen.json', b'{}')
    create_backup(directory, 'Vóór herstellen van ' + identifier)
    for target, raw in photos:
        target.parent.mkdir(parents=True, exist_ok=True)
        atomic_bytes(target, raw)
    for name, raw in files.items():
        atomic_bytes(directory / name, raw)
    return manifest.get('missing', [])


def parquet_bytes(data, local=None):
    import pyarrow as pa
    import pyarrow.parquet as pq
    arrays, names = [], []
    for index, name in enumerate(data['headers']):
        values = [r['values'][index] if known(r['values'][index]) else None for r in data['records']]
        numeric = index in (11, 12, 13) and all(v is None or type(v) in (int, float) for v in values)
        arrays.append(pa.array(values if numeric else [None if v is None else str(v) for v in values], type=pa.float64() if numeric else pa.string()))
        names.append(name)
    extras = {
        'Opmerking': (pa.string(), lambda r: r.get('note', '')),
        'Bron': (pa.string(), lambda r: r.get('origin', 'excel')),
        'Foto’s': (pa.list_(pa.string()), lambda r: r.get('extra', {}).get('photos', [])),
        'Zoekertjestekst': (pa.string(), lambda r: r.get('extra', {}).get('ad', '')),
        'ChatGPT URL': (pa.string(), lambda r: r.get('extra', {}).get('chatUrl') or ('https://chatgpt.com/c/' + r['extra']['chatId'] if r.get('extra', {}).get('chatId') else None)),
        'Verkoopprijs': (pa.float64(), lambda r: r.get('sale', {}).get('price')),
        'Verkoopdatum': (pa.string(), lambda r: r.get('sale', {}).get('date')),
        'Afspraak': (pa.string(), lambda r: json.dumps(r['appointment'], ensure_ascii=False) if r.get('appointment') else None),
        'Bundel ID': (pa.string(), lambda r: r.get('bundle', {}).get('id')),
        'Bundel naam': (pa.string(), lambda r: r.get('bundle', {}).get('name')),
        'Wijzigingsgeschiedenis': (pa.string(), lambda r: json.dumps((local or {}).get('history', {}).get(str(r['values'][0]), []), ensure_ascii=False)),
    }
    for name, (kind, getter) in extras.items():
        names.append(name)
        arrays.append(pa.array([getter(r) for r in data['records']], type=kind))
    table = pa.Table.from_arrays(arrays, names=names).replace_schema_metadata({'exported': now(), 'scope': 'volledige actieve collectie', 'excel_sha256': data.get('sha256', '')})
    output = io.BytesIO()
    pq.write_table(table, output, compression='snappy')
    return output.getvalue()


def export_zip(directory, data):
    directory = Path(directory)
    files = inventory_files(directory)
    local = json.loads(files.get('lokale_modellen.json', b'{}'))
    output = io.BytesIO()
    missing = []
    with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('collectie.json', json.dumps(data, ensure_ascii=False, indent=2))
        archive.writestr('collectie.parquet', parquet_bytes(data, local))
        archive.writestr('verkoopbundels.json', json.dumps(local.get('bundles', {}), ensure_ascii=False, indent=2))
        archive.writestr('wijzigingsgeschiedenis.json', json.dumps(local.get('history', {}), ensure_ascii=False, indent=2))
        for name, raw in files.items():
            archive.writestr(name, raw)
        for relative in referenced_photos(files):
            path = safe_photo(directory, relative)
            if path.is_file():
                archive.write(path, relative)
            else:
                missing.append(relative)
        archive.writestr('export-info.json', json.dumps({'created': now(), 'models': len(data['records']), 'missing_photos': missing}, ensure_ascii=False, indent=2))
    return output.getvalue()

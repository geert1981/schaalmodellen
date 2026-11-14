"""Lees Excel uitsluitend; genereer een Parquet-export en offline HTML-inventaris."""
import datetime
import hashlib
import json
import pathlib
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from zoneinfo import ZoneInfo

SOURCE = pathlib.Path(r'C:\Users\ntm556\OneDrive - De Stroomlijn\Geert Kril\Schaalmodellen\Inventaris.xlsx')
HEADERS = ['Nr.', 'Merk (echte auto)', 'Model / uitvoering', 'Bouwjaar (echte auto)', 'Merk modelauto (fabrikant)', 'Modelnr. / artikelnummer', 'Conditie', 'Schaal', 'Oplage', 'Locatie', 'Doos (ja/nee/onbekend)', 'Vraagprijs zonder doos', 'Vraagprijs met doos', 'Vraagprijs met doos als nieuw', '2dehands (status zoekertje)']
FIELDS = ['nr', 'automerk', 'model', 'bouwjaar', 'merk modelauto', 'modelnr', 'conditie', 'schaal', 'oplage', 'locatie', 'doos', 'vraagprijs zonder doos', 'vraagprijs met doos', 'vraagprijs met doos als nieuw', '2e hands']
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
LOCAL_FILE = 'lokale_modellen.json'
CODE_DIR = pathlib.Path(__file__).resolve().parent
DATA_DIR = CODE_DIR / 'data'  # junction naar OneDrive
TEMPLATE = CODE_DIR / 'inventaris.template.html'


def load_local(directory):
    path = directory / LOCAL_FILE
    if not path.exists():
        return {'models': [], 'excel_media': {}}
    data = json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(data.get('models'), list) or not isinstance(data.get('excel_media'), dict):
        raise ValueError('Ongeldige structuur in ' + LOCAL_FILE)
    for key in ('excel_edits', 'hidden_excel'):
        if not isinstance(data.get(key, {}), dict):
            raise ValueError('Ongeldige structuur in ' + LOCAL_FILE + ': ' + key)
    if not isinstance(data.get('archived_models', []), list):
        raise ValueError('Ongeldig archief in ' + LOCAL_FILE)
    return data


def local_revision(local):
    return hashlib.sha256(json.dumps(local, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()


def merge_local(data, local):
    # Haal eerdere lokale aanvullingen uit de momentopname voordat we opnieuw samenvoegen.
    data['records'] = [r for r in data['records'] + data.pop('hiddenRecords', []) if r.get('origin') != 'local']
    data['hiddenRecords'] = []
    data['localWarnings'] = []
    data['localRevision'] = local_revision(local)
    applied = {name: set() for name in ('excel_media', 'excel_edits', 'hidden_excel')}
    for record in data['records']:
        record['origin'] = 'excel'
        if 'originalValues' in record:
            record['values'] = record.pop('originalValues')
            record['note'] = record.pop('originalNote')
        record.pop('locallyEdited', None)
        if 'originalExtra' in record:
            original = record.pop('originalExtra')
            record.pop('extra', None)
            if original is not None:
                record['extra'] = original
        key = str(record['values'][0])
        identity = record['values'][1:6]
        hidden = local.get('hidden_excel', {}).get(key)
        if hidden and hidden['identity'] == identity:
            data['hiddenRecords'].append(record)
            applied['hidden_excel'].add(key)
        media = local['excel_media'].get(key)
        if media and media['identity'] == identity:
            record['originalExtra'] = record.get('extra')
            record['extra'] = {**(record.get('extra') or {}), **media['extra']}
            applied['excel_media'].add(key)
        edits = local.get('excel_edits', {}).get(key)
        if edits and edits['identity'] == identity:
            record['originalValues'] = record['values'][:]
            record['originalNote'] = record.get('note', '')
            for index, value in edits['values'].items():
                record['values'][int(index)] = value
            if 'note' in edits:
                record['note'] = edits['note']
            record['locallyEdited'] = True
            applied['excel_edits'].add(key)
    for name, label in [('excel_media', 'Foto’s en tekst'), ('excel_edits', 'Pagina-aanpassingen'), ('hidden_excel', 'De verwijdering')]:
        for key in local.get(name, {}):
            if key not in applied[name]:
                data['localWarnings'].append(label + ' voor Excel-model ' + key + ' zijn bewaard, maar niet toegepast omdat het model ontbreekt of de identiteit is gewijzigd.')
    data['records'] = [r for r in data['records'] if str(r['values'][0]) not in applied['hidden_excel']]
    used = {str(r['values'][0]) for r in data['records']}
    for record in local['models']:
        key = str(record['values'][0])
        if key in used:
            raise ValueError('Dubbel modelnummer: ' + key)
        used.add(key)
        data['records'].append(record)
    data['trash'] = [{'id': str(r['values'][0]), 'title': str(r['values'][1]) + ' ' + str(r['values'][2]), 'origin': 'local'} for r in local.get('archived_models', [])]
    data['trash'].extend({'id': key, 'title': item['title'], 'origin': 'excel'} for key, item in local.get('hidden_excel', {}).items())
    return data


def snapshot(directory):
    html = (directory / 'index.html').read_text(encoding='utf-8')
    match = re.search(r'<script id="inventory-data" type="application/json">(.*?)</script>', html, re.S)
    if not match:
        raise ValueError('Geen inventarisgegevens gevonden in index.html. Voer Bijwerken.cmd uit.')
    return json.loads(match.group(1))


def render_html(directory, data):
    payload = json.dumps(data, ensure_ascii=False).replace('<', '\\u003c').replace('>', '\\u003e').replace('&', '\\u0026')
    template = TEMPLATE.read_text(encoding='utf-8')
    if template.count('__INVENTARIS_DATA__') != 1:
        raise ValueError('Template moet precies één gegevensmarkering bevatten.')
    return template.replace('__INVENTARIS_DATA__', payload)

def extract(path, *, include_raw=False):
    raw = path.read_bytes()
    import io
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        shared = [''.join(t.text or '' for t in s.findall('.//m:t', NS)) for s in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', NS)] if 'xl/sharedStrings.xml' in z.namelist() else []
        book = ET.fromstring(z.read('xl/workbook.xml'))
        sheet = next(s for s in book.findall('m:sheets/m:sheet', NS) if s.get('name') == 'Inventaris')
        rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        target = next(r.get('Target') for r in rels if r.get('Id') == rid)
        target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
        rows = []
        for r in ET.fromstring(z.read(target)).findall('m:sheetData/m:row', NS):
            cells = {}
            for c in r.findall('m:c', NS):
                letters = ''.join(x for x in c.get('r') if x.isalpha())
                v = c.find('m:v', NS)
                value = v.text if v is not None else None
                typ = c.get('t')
                if typ == 's' and value is not None:
                    value = shared[int(value)]
                elif typ == 'inlineStr':
                    value = ''.join(t.text or '' for t in c.findall('m:is//m:t', NS))
                elif value is not None and typ == 'b':
                    value = value == '1'
                elif value is not None and typ not in ('str', 'e', 'd'):
                    value = float(value)
                    if value.is_integer():
                        value = int(value)
                cells[letters] = value
            rows.append(cells)
        columns = {str(v).strip().lower(): k for k, v in rows[0].items() if v is not None}
        if 'automerk' not in columns and 'merk' in columns:
            columns['automerk'] = columns['merk']
        missing = [f for f in FIELDS if f not in columns]
        if missing:
            raise ValueError('Verwachte kolommen ontbreken: ' + ', '.join(missing))
        records = []
        for row in rows[1:]:
            values = [row.get(columns[f]) for f in FIELDS]
            note = row.get(columns.get('opmerking', ''))
            # Vooraf ingevulde nummers en standaardschaal vormen geen modelregistratie.
            if not any(v is not None and str(v).strip() for i, v in enumerate(values) if i not in (0, 7)) and not note:
                continue
            if isinstance(values[7], (int, float)) and 0 <= values[7] < 1:
                minutes = round(values[7] * 24 * 60)
                values[7] = f'{minutes // 60}:{minutes % 60:02}'
            values = ['Onbekend' if v is None or str(v).strip() == '' else v for v in values]
            records.append({'values': values, 'note': note or ''})
    data = {'headers': HEADERS, 'records': records, 'source': str(path), 'sha256': hashlib.sha256(raw).hexdigest(), 'updated': datetime.datetime.now(ZoneInfo('Europe/Brussels')).isoformat(timespec='seconds')}
    if include_raw:
        data['_raw_rows'] = rows
    return data


def parquet_table(rows, data):
    try:
        import pyarrow as pa
    except ImportError as exc:
        raise RuntimeError('Parquet vereist pyarrow. Installeer met: python -m pip install -r requirements.txt') from exc
    columns = [(col, str(value)) for col, value in rows[0].items() if value is not None and str(value).strip()]
    names = [name for _, name in columns]
    if len(set(names)) != len(names):
        raise ValueError('Dubbele Excel-kolomnamen: Parquet-export afgebroken.')
    known = {col for col, _ in columns}
    body = [row for row in rows[1:] if any(v is not None for v in row.values())]
    if any(v is not None and col not in known for row in body for col, v in row.items()):
        raise ValueError('Excel bevat gegevens in een kolom zonder koptekst.')
    arrays = []
    for col, _ in columns:
        values = [row.get(col) for row in body]
        present = [v for v in values if v is not None]
        # Parquet heeft per kolom een vast type; gemengde Excel-kolommen worden tekst.
        if not present or all(isinstance(v, str) for v in present):
            array = pa.array(values, type=pa.string())
        elif all(type(v) is bool for v in present):
            array = pa.array(values, type=pa.bool_())
        elif all(type(v) in (int, float) for v in present):
            array = pa.array(values, type=pa.float64() if any(type(v) is float for v in present) else pa.int64())
        else:
            array = pa.array([None if v is None else str(v) for v in values], type=pa.string())
        arrays.append(array)
    metadata = {key: str(data[key]) for key in ('source', 'sha256', 'updated')}
    metadata['sheet'] = 'Inventaris'
    return pa.Table.from_arrays(arrays, names=names).replace_schema_metadata(metadata)

def main(directory=None):
    directory = pathlib.Path(directory) if directory is not None else DATA_DIR
    if not directory.is_dir():
        raise RuntimeError(f'Gegevensmap niet gevonden: {directory}. Maak de koppeling naar OneDrive opnieuw aan (zie LEESMIJ.txt).')
    data = extract(SOURCE, include_raw=True)
    table = parquet_table(data.pop('_raw_rows'), data)
    import pyarrow.parquet as pq
    extras_path = directory / 'aanvullingen.json'
    extras = json.loads(extras_path.read_text(encoding='utf-8')) if extras_path.exists() else {}
    for record in data['records']:
        extra = extras.get(str(record['values'][0]))
        if extra and extra.get('identity') == record['values'][1:6]:
            record['extra'] = extra
    data['unmatched'] = extras.get('_unmatched', [])
    merge_local(data, load_local(directory))
    html = render_html(directory, data)
    destination = directory / 'index.html'
    temporary = directory / 'index.html.tmp'
    parquet_path = directory / 'inventaris.parquet'
    parquet_tmp = directory / 'inventaris.parquet.tmp'
    try:
        pq.write_table(table, parquet_tmp, compression='snappy')
        temporary.write_text(html, encoding='utf-8')
        if hashlib.sha256(SOURCE.read_bytes()).hexdigest() != data['sha256']:
            raise RuntimeError('Excel veranderde tijdens het lezen. Probeer opnieuw.')
        parquet_tmp.replace(parquet_path)
        temporary.replace(destination)
    finally:
        parquet_tmp.unlink(missing_ok=True)
        temporary.unlink(missing_ok=True)
    try:
        import miniaturen
        miniaturen.ensure_for_data(directory, data)
    except Exception as exc:
        print(f'Miniaturen bijwerken mislukt: {exc}', file=sys.stderr)
    print(f'Bijgewerkt: {len(data["records"])} modellen in HTML; {table.num_rows} Excel-rijen en {table.num_columns} kolommen in Parquet. Excel ongewijzigd.\n{destination}\n{parquet_path}')

if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(f'Bijwerken mislukt: {exc}\nSluit Excel als het bestand vergrendeld is. De vorige pagina blijft behouden.', file=sys.stderr)
        sys.exit(1)

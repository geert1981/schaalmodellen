"""Verkoopbundels en wijzigingsgeschiedenis, los van de Excel-bron."""
import copy
import datetime
import uuid


def identity(record):
    return None if record.get('origin') == 'local' else record.get('originalValues', record['values'])[1:6]


def put_fields(local, record, values, note=None, extra=None):
    key = str(record['values'][0])
    if record.get('origin') == 'local':
        stored = next(r for r in local['models'] if str(r['values'][0]) == key)
        stored['values'] = copy.deepcopy(values)
        if note is not None:
            stored['note'] = note
        if extra is not None:
            stored['extra'] = copy.deepcopy(extra)
        return
    baseline = record.get('originalValues', record['values'])
    edits = {'identity': identity(record), 'values': {str(i): values[i] for i in range(1, 15) if values[i] != baseline[i]}}
    if note is None:
        note = record.get('note', '')
    if note != record.get('originalNote', record.get('note', '')):
        edits['note'] = note
    local.setdefault('excel_edits', {}).pop(key, None)
    if edits['values'] or 'note' in edits:
        local['excel_edits'][key] = edits
    if extra is not None:
        local['excel_media'][key] = {'identity': identity(record), 'extra': copy.deepcopy(extra)}


def attach(data, local):
    records = data['records']
    for record in records + data['hiddenRecords']:
        record.pop('bundle', None)
        record.pop('appointmentScope', None)
        record['historyCount'] = len(local.get('history', {}).get(str(record['values'][0]), []))
    by_id = {str(r['values'][0]): r for r in records}
    data['bundles'] = []
    for bundle in local.get('bundles', {}).values():
        view = copy.deepcopy(bundle)
        view['members'] = []
        for member in bundle['members']:
            record = by_id.get(member['id'])
            available = record is not None and identity(record) == member['identity']
            view['members'].append({**member, 'available': available})
            if bundle['status'] == 'Geannuleerd':
                continue
            if not available:
                data['localWarnings'].append('Model ' + member['id'] + ' uit bundel ' + bundle['name'] + ' ontbreekt of heeft een gewijzigde identiteit. Controleer de bundel.')
                continue
            record['bundle'] = {k: copy.deepcopy(bundle[k]) for k in ('id', 'name', 'buyer', 'contact', 'total', 'status', 'date', 'note')}
            record['bundle']['memberIds'] = [m['id'] for m in bundle['members']]
            record.pop('appointment', None)
            if bundle.get('appointment'):
                record['appointment'] = {**bundle['appointment'], 'buyer': bundle['buyer'], 'note': bundle['note']}
                record['appointmentScope'] = 'bundle'
            if bundle['status'] == 'Verkocht':
                # Een bundeltotaal is geen individuele verkoopprijs.
                record['sale'] = {'price': None, 'date': bundle['date']}
        data['bundles'].append(view)


def all_records(data, local):
    result = {str(r['values'][0]): (r, 'active') for r in data['records']}
    for record in data.get('hiddenRecords', []) + local.get('archived_models', []):
        result[str(record['values'][0])] = (record, 'trashed')
    return result


def state(record, visibility):
    values = {f'value:{i}': copy.deepcopy(record['values'][i]) for i in range(1, 15)}
    extra = record.get('extra') or {}
    url = extra.get('chatUrl') or ('https://chatgpt.com/c/' + extra['chatId'] if extra.get('chatId') else '')
    values.update({'note': record.get('note', ''), 'photos': copy.deepcopy(extra.get('photos', [])), 'ad': extra.get('ad', ''), 'chatUrl': url, 'appointment': copy.deepcopy(record.get('appointment')), 'sale': copy.deepcopy(record.get('sale')), 'bundle': record.get('bundle', {}).get('id'), 'bundleInfo': copy.deepcopy(record.get('bundle')), 'visibility': visibility})
    return values


def log_changes(local, before_data, before_local, after_data, action, undo_of=None):
    before, after = all_records(before_data, before_local), all_records(after_data, local)
    for key in before.keys() | after.keys():
        old, new = before.get(key), after.get(key)
        old_state = state(*old) if old else {}
        new_state = state(*new) if new else {}
        changes = [{'field': field, 'before': old_state.get(field), 'after': new_state.get(field)} for field in sorted(old_state.keys() | new_state.keys()) if old_state.get(field) != new_state.get(field)]
        if not changes:
            continue
        record = (new or old)[0]
        bundled = bool(old_state.get('bundle') or new_state.get('bundle'))
        blocked = not old or not new or identity(old[0]) != identity(new[0]) or any(c['field'] in ('visibility', 'bundle', 'bundleInfo') for c in changes) or (bundled and any(c['field'] in ('value:14', 'appointment', 'sale') for c in changes))
        event = {'id': uuid.uuid4().hex, 'at': datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='seconds'), 'action': action, 'identity': identity(record), 'changes': changes, 'photos': sorted(set(old_state.get('photos', []) + new_state.get('photos', []))), 'undoable': not blocked}
        if undo_of:
            event['undoOf'] = undo_of
        local.setdefault('history', {}).setdefault(key, []).append(event)


def event_for(local, key, event_id):
    event = next((e for e in local.get('history', {}).get(key, []) if e['id'] == event_id), None)
    if not event:
        raise ValueError('Deze wijziging staat niet in de geschiedenis van het model.')
    return event


def undo_event(local, data, key, event_id):
    event = event_for(local, key, event_id)
    record = next((r for r in data['records'] if str(r['values'][0]) == key), None)
    if not record or identity(record) != event['identity']:
        raise ValueError('Dit model ontbreekt of heeft inmiddels een andere identiteit.')
    if not event['undoable']:
        raise ValueError('Gebruik de prullenbak of de verkoopbundel om deze actie terug te draaien.')
    current = state(record, 'active')
    if any(current.get(c['field']) != c['after'] for c in event['changes']):
        raise ValueError('Een van deze velden is daarna opnieuw gewijzigd. Draai eerst die latere wijziging terug.')
    if record.get('bundle') and any(c['field'] in ('value:14', 'appointment', 'sale') for c in event['changes']):
        raise ValueError('Beheer status en afspraak nu via de verkoopbundel.')
    restored = {**current, **{c['field']: copy.deepcopy(c['before']) for c in event['changes']}}
    values = record['values'][:]
    for i in range(1, 15):
        values[i] = restored[f'value:{i}']
    extra = {**(record.get('extra') or {}), **{k: restored[k] for k in ('photos', 'ad', 'chatUrl')}}
    put_fields(local, record, values, restored['note'], extra)
    for field, storage in (('appointment', 'appointments'), ('sale', 'sales')):
        if not any(c['field'] == field for c in event['changes']):
            continue
        local.setdefault(storage, {}).pop(key, None)
        if restored[field] is not None:
            local[storage][key] = {'identity': identity(record), **restored[field]}
    return event

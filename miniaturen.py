"""Kleine voorbeeldfoto's (miniaturen) voor het foto-overzicht.

De originele foto's blijven ongewijzigd. Miniaturen staan in fotos/_mini/ met
hetzelfde relatieve pad plus '.jpg', bijvoorbeeld:
    fotos/web/abc.png  ->  fotos/_mini/web/abc.png.jpg
De pagina valt terug op het origineel als een miniatuur ontbreekt, dus zonder
Pillow of bij een fout blijft alles werken (alleen trager).

Eenmalig of handmatig alles aanmaken:  python miniaturen.py
"""
import pathlib
import sys

MAX_SIZE = 480      # langste zijde in pixels
QUALITY = 80
PREFIX = 'fotos/'
MINI_DIR = 'fotos/_mini/'


def thumb_path(photo):
    """Relatief pad van de miniatuur voor een fotopad uit de inventaris."""
    if not photo.startswith(PREFIX) or photo.startswith(MINI_DIR):
        return None
    return MINI_DIR + photo[len(PREFIX):] + '.jpg'


def photos_in(data):
    for record in data.get('records', []) + data.get('hiddenRecords', []):
        yield from (record.get('extra') or {}).get('photos', [])


def ensure(directory, photos):
    """Maak ontbrekende of verouderde miniaturen. Geeft (gemaakt, mislukt) terug."""
    try:
        from PIL import Image, ImageOps
    except ImportError:
        return 0, 0
    directory = pathlib.Path(directory).resolve()
    fotos = (directory / 'fotos').resolve()
    made = failed = 0
    for photo in dict.fromkeys(photos):
        relative = thumb_path(photo) if isinstance(photo, str) else None
        if not relative:
            continue
        source = (directory / photo).resolve()
        target = (directory / relative).resolve()
        if not source.is_relative_to(fotos) or not target.is_relative_to(fotos) or not source.is_file():
            continue
        if target.exists() and target.stat().st_mtime >= source.stat().st_mtime:
            continue
        try:
            with Image.open(source) as image:
                image = ImageOps.exif_transpose(image)
                image.thumbnail((MAX_SIZE, MAX_SIZE))
                if image.mode not in ('RGB', 'L'):
                    background = Image.new('RGB', image.size, (255, 255, 255))
                    rgba = image.convert('RGBA')
                    background.paste(rgba, mask=rgba.getchannel('A'))
                    image = background
                target.parent.mkdir(parents=True, exist_ok=True)
                temporary = target.with_name(target.name + '.tmp')
                image.save(temporary, 'JPEG', quality=QUALITY, optimize=True, progressive=True)
                temporary.replace(target)
                made += 1
        except Exception:  # Een kapotte foto mag de app niet tegenhouden.
            failed += 1
    return made, failed


def ensure_for_data(directory, data):
    return ensure(directory, list(photos_in(data)))


def main(directory=None):
    import bijwerken
    directory = pathlib.Path(directory) if directory else bijwerken.DATA_DIR
    try:
        import PIL  # noqa: F401
    except ImportError:
        print('Pillow ontbreekt. Installeer met: python -m pip install -r requirements.txt')
        return 1
    made, failed = ensure_for_data(directory, bijwerken.snapshot(directory))
    print(f'Miniaturen aangemaakt: {made}' + (f' · mislukt: {failed}' if failed else ''))
    return 0


if __name__ == '__main__':
    sys.exit(main())

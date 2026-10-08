"""Independent media fidelity oracle for media.mjs, using Python ZIP/XML.

python3 tools/lectern/test/media-check.py ORIGINALS OPERATIONS_DIR
python3 tools/lectern/test/media-check.py --sounds ORIGINALS RESULTS.jsonl SAVED_ROOT
Checks bytes and intended relationship targets, unedited picture properties,
playback behaviours, and exact undo/redo package restoration (except timestamps).
"""
from collections import Counter
import copy
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'ooxml'))
from package import Package, R, MC, split, at

MEDIA_TAGS = {'videoFile', 'audioFile', 'wavAudioFile', 'audioCd', 'media'}
def descendants(el, name):
    return [e for e in el.iter() if split(e.tag)[1] == name]

def canonical(el, pkg, owner, ignore=()):
    ns, tag = split(el.tag)
    if tag in ignore: return None
    attributes = []
    for key, value in sorted(el.attrib.items()):
        ans, name = split(key)
        if ans == MC or ans == 'http://www.w3.org/XML/1998/namespace': continue
        if ans in R and value:
            dep = pkg.rels[owner][value]
            target = ('external', dep['target']) if dep['external'] else ('sha256', hashlib.sha256(pkg.data[dep['target']]).hexdigest())
            value = (dep['type'], target)
        elif name == 'id' and tag in ('cNvPr', 'cTn'): value = '#definition'
        elif name == 'spid': value = '#shape'
        elif name == 'val' and tag == 'tn': value = '#time'
        attributes.append((key, value))
    children = tuple(c for child in el if (c := canonical(child, pkg, owner, ignore)) is not None)
    if tag == 'effectLst' and 'outerShdw' in ignore and not attributes and not children: return None
    return (el.tag, tuple(attributes), (el.text or '').strip(), children)

def pictures(pkg):
    found = []
    for part, root in pkg.xml.items():
        if split(root.tag)[1] != 'sld': continue
        for pic in descendants(root, 'pic'):
            if not any(split(e.tag)[1] in MEDIA_TAGS for e in pic.iter()): continue
            nv = descendants(pic, 'cNvPr')[0]
            found.append((part, at(nv, 'id'), pic))
    return found

def playback(pkg, part, shape):
    values = []
    for el in pkg.xml[part].iter():
        tag = split(el.tag)[1]
        if tag not in ('audio', 'video') and not (tag == 'cTn' and el.get('presetClass') == 'mediacall'): continue
        if any(at(e, 'spid') == shape for e in descendants(el, 'spTgt')):
            values.append(repr(canonical(el, pkg, part)))
    return Counter(values)

def equal_states(before, after):
    assert set(before.data) == set(after.data), 'Undo/redo changed the package part names'
    for name, data in before.data.items():
        if data == after.data[name]: continue
        assert name == 'docProps/core.xml', 'Undo/redo changed ' + name
        a, b = copy.deepcopy(before.xml[name]), copy.deepcopy(after.xml[name])
        for root in (a, b):
            for el in list(root):
                if split(el.tag)[1] in ('created', 'modified'): root.remove(el)
        assert canonical(a, before, name) == canonical(b, after, name), 'Undo/redo changed core properties'

def sound_inventory(pkg):
    result = Counter()
    for part, root in pkg.xml.items():
        if split(root.tag)[1] != 'sld': continue
        for el in root.iter():
            tag = split(el.tag)[1]
            if tag in ('audio', 'video', 'sndAc', 'sndTgt') or tag == 'hlinkClick' and descendants(el, 'snd'):
                result[repr(canonical(el, pkg, part))] += 1
    return result

def sounds():
    originals, report, directory = map(Path, sys.argv[2:5])
    failures = 0
    for line in report.read_text().splitlines():
        row = json.loads(line)
        if row['status'] != 'ok': continue
        name, scenario = row['file'], row['scenario']
        for suffix in ('', '-undo', '-redo', '-second', '-draft', '-recovered'):
            file = directory / (scenario + suffix) / name
            if not file.exists(): continue
            result = dict(file=name, scenario=scenario + suffix, status='ok')
            try:
                before, after = Package(originals / name), Package(file)
                assert sound_inventory(before) == sound_inventory(after), 'Sound properties, playback, dependency bytes or intended targets changed'
            except Exception as error:
                result.update(status='failed', error=str(error)); failures += 1
            print(json.dumps(result))
    return int(failures != 0)

def main():
    originals, directory = map(Path, sys.argv[1:3])
    failures = 0
    for line in (directory / 'results.jsonl').read_text().splitlines():
        row = json.loads(line); name, scenario = row['file'], row['scenario']
        result = dict(file=name, scenario=scenario, status=row['status'])
        if row['status'] != 'ok':
            print(json.dumps(dict(result, error=row.get('error', row.get('reason'))))); failures += row['status'] == 'failed'; continue
        try:
            original = Package(originals / name)
            saved = Package(directory / scenario / name)
            before_file = directory / (scenario + '-before') / name
            before = Package(before_file) if before_file.exists() else original
            a, b = pictures(before), pictures(saved)
            target = next((p for p in a if p[:2] == (row.get('sourcePart'), str(row.get('shapeId')))), a[0])
            expected = len(a)
            if scenario == 'duplicate': expected += 1
            if scenario == 'duplicate-slide': expected += sum(p[0] == target[0] for p in a)
            if scenario in ('delete', 'replace'): expected -= 1
            if scenario == 'copy': expected = 1
            assert len(b) == expected, f'Media pictures: expected {expected}, got {len(b)}'
            ignore = {'xfrm'} if scenario in ('geometry', 'copy') else {'outerShdw'} if scenario == 'shadow' else {'srcRect'} if scenario == 'crop' else set()
            wanted = Counter(repr(canonical(pic, before, part, ignore)) for part, _, pic in a)
            selected = repr(canonical(target[2], before, target[0], ignore))
            if scenario == 'duplicate': wanted[selected] += 1
            if scenario in ('delete', 'replace'): wanted[selected] -= 1
            if scenario == 'duplicate-slide': wanted.update(repr(canonical(pic, before, part, ignore)) for part, _, pic in a if part == target[0])
            if scenario == 'copy': wanted = Counter({selected: 1})
            actual = Counter(repr(canonical(pic, saved, part, ignore)) for part, _, pic in b)
            assert +wanted == actual, 'Unedited media properties, dependency bytes or targets changed'
            if scenario in ('geometry', 'crop', 'shadow'):
                edited = next(pic for part, identity, pic in b if (part, identity) == target[:2])
                prop = {'geometry': 'xfrm', 'crop': 'srcRect', 'shadow': 'outerShdw'}[scenario]
                old = [canonical(el, before, target[0]) for el in descendants(target[2], prop)]
                new = [canonical(el, saved, target[0]) for el in descendants(edited, prop)]
                assert new and new != old, 'The requested ' + scenario + ' edit was not written'
                if scenario == 'geometry' and row.get('expectedBox'):
                    transform = descendants(edited, 'xfrm')[0]
                    off, extent = descendants(transform, 'off')[0], descendants(transform, 'ext')[0]
                    box = [float(off.get('x')), float(off.get('y')), float(extent.get('cx')), float(extent.get('cy'))]
                    assert all(abs(a - b * 12700) <= 1 for a, b in zip(box, row['expectedBox'])), 'Moved/resized transform does not match the edited box'
            for part, identity, pic in b:
                key = repr(canonical(pic, saved, part, ignore))
                src = next(p for p in a if repr(canonical(p[2], before, p[0], ignore)) == key)
                assert playback(saved, part, identity) == playback(before, src[0], src[1]), 'Playback behaviours or properties changed'
            if scenario == 'save':
                assert Counter(repr(canonical(pic, original, part)) for part, _, pic in pictures(original)) == actual, 'Unedited open/save changed a media frame'
            for suffix, expected_state in [('undo', before), ('redo', saved)]:
                file = directory / (scenario + '-' + suffix) / name
                if file.exists(): equal_states(expected_state, Package(file))
            result.update(mediaPictures=len(b), restoredStates=sum((directory / (scenario + '-' + s) / name).exists() for s in ('undo', 'redo')))
        except Exception as error:
            result.update(status='failed', error=str(error)); failures += 1
        print(json.dumps(result))
    return int(failures != 0)

if __name__ == '__main__': sys.exit(sounds() if sys.argv[1] == '--sounds' else main())

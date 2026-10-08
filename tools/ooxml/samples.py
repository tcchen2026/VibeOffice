"""Choose independently authored examples for the Office handoff. No fixture bytes enter git.
Usage: python3 tools/ooxml/samples.py CORPORA RESULTS_ROOT DESTINATION
RESULTS_ROOT has quire/save, ledger/save, lectern/save (and optionally <app>/text).
"""
import hashlib
import json
from pathlib import Path
import re
import shutil
import sys
import zipfile

FEATURES = {
    'quire': {'content-controls': r'<(?:\w+:)?sdt\b', 'ole': r'<(?:\w+:)?OLEObject\b', 'smartart': r'<(?:\w+:)?relIds\b',
              'custom-xml': r'customXml/item', 'labels': r'LabelInfo', 'fonts': r'<(?:\w+:)?embedRegular\b', 'macros': r'vbaProject.bin',
              'template': r'attachedTemplate', 'permissions': r'<(?:\w+:)?permStart\b', 'altchunk': r'<(?:\w+:)?altChunk\b'},
    'ledger': {'pivots': r'<(?:\w+:)?pivotTableDefinition\b', 'controls': r'<(?:\w+:)?controls\b', 'ole': r'<(?:\w+:)?oleObjects\b',
               'connections': r'xl/connections.xml', 'slicers': r'xl/slicers/', 'threads': r'<(?:\w+:)?threadedComment\b',
               'custom-xml': r'customXml/item', 'labels': r'LabelInfo', 'macros': r'vbaProject.bin'},
    'lectern': {'video': r'<(?:\w+:)?videoFile\b', 'audio': r'<(?:\w+:)?audioFile\b', 'smartart': r'<(?:\w+:)?relIds\b',
                'sections': r'<(?:\w+:)?sectionLst\b', 'comments': r'ppt/comments/', 'fonts': r'<(?:\w+:)?embeddedFontLst\b',
                'effects': r'<(?:\w+:)?(?:glow|sp3d|reflection)\b', 'ole': r'<(?:\w+:)?oleObj\b', 'macros': r'vbaProject.bin',
                'custom-xml': r'customXml/item', 'labels': r'LabelInfo', 'handout': r'ppt/handoutMasters/'},
}

def main():
    corpus, results, dest = map(Path, sys.argv[1:4])
    manifest = []
    for app, features in FEATURES.items():
        todo = dict(features)
        files = corpus / {'quire': 'word', 'ledger': 'excel', 'lectern': 'powerpoint'}[app]
        # Prefer small, readily inspectable examples; require a successful baseline save.
        for file in sorted(files.iterdir(), key=lambda f: (f.stat().st_size, f.name)):
            saved = results / app / 'save' / file.name
            if not saved.exists(): continue
            try:
                with zipfile.ZipFile(file) as z:
                    text = '\n'.join(z.namelist()) + '\n' + '\n'.join(z.read(n).decode('utf-8', 'replace') for n in z.namelist() if n.endswith('.xml') and z.getinfo(n).file_size < 10_000_000)
            except Exception:
                continue
            for feature, pattern in list(todo.items()):
                if not re.search(pattern, text): continue
                folder = dest / app / feature
                folder.mkdir(parents=True, exist_ok=True)
                row = dict(app=app, feature=feature, source=file.name, sha256=hashlib.sha256(file.read_bytes()).hexdigest(), files=[])
                for kind, source in [('original', file), ('saved', saved), ('edited', results / app / 'text' / file.name)]:
                    if not source.exists(): continue
                    target = folder / (kind + file.suffix)
                    if target.exists() and target.read_bytes() != source.read_bytes():
                        raise ValueError(f'Refusing to overwrite a different handoff file: {target}; use a new destination')
                    shutil.copyfile(source, target)
                    row['files'].append(str(target.relative_to(dest)))
                manifest.append(row)
                del todo[feature]
            if not todo: break
        for feature in todo:
            manifest.append(dict(app=app, feature=feature, excluded='No matching successfully saved fixture in this corpus'))
    dest.mkdir(parents=True, exist_ok=True)
    (dest / 'manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
    print(json.dumps(dict(samples=sum('files' in m for m in manifest), missing=sum('excluded' in m for m in manifest))))

if __name__ == '__main__':
    main()

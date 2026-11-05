"""Fetch pinned, sparse public corpora; never delete or reset an existing checkout."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import urllib.request

SOURCES = {
    'libreoffice': ('https://github.com/LibreOffice/core', '2821d29a87b28785d74fa64d975465e4c99d4416', ['sw/qa', 'sc/qa', 'sd/qa', 'oox/qa']),
    'poi': ('https://github.com/apache/poi', 'ae62bb5116b9aee19ebd5834e3a82066132c9f7f', ['test-data']),
    'openxml-sdk': ('https://github.com/dotnet/Open-XML-SDK', '431ab05cf160248cc3885a4a766026d4f8243792', ['test']),
}
EXTENSIONS = {'word': ['docx', 'docm', 'dotx', 'dotm'], 'excel': ['xlsx', 'xlsm', 'xltx', 'xltm'],
              'powerpoint': ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm']}
BY_EXT = {'.' + ext: app for app, exts in EXTENSIONS.items() for ext in exts}
LEDGER_FEATURES = [
    ('XLibur/XLibur', 'cc7547ff9bf67077e47fbffc38d313b3c6cd31aa',
     'XLibur.Tests/Resource/TryToLoad/SlicersOnPivotAndTable.xlsx', '4dc1bd8cac0c5c50d15c7460ce2c5479974442c4ea6c7f441f35f3f77c70a280'),
    ('XLibur/XLibur', 'cc7547ff9bf67077e47fbffc38d313b3c6cd31aa',
     'XLibur.Tests/Resource/TryToLoad/Timelines_Missing_21232.xlsx', '662c40550a6b296deee93c744bc7bcb37a93196104ef907cd2dafea5a167f6b9'),
    ('Ryvier/CoffeeSalesProject', '4b94ad3d35d7f50f385051d5adaac426d391e03f',
     'coffeeSalesProject.xlsx', 'd8e288300dfb18e1912289ccfb4e346398667afd928151193be9868f9aaee048'),
]

def ledger_features(root):
    dest = root / 'ledger-features'
    dest.mkdir(parents=True, exist_ok=True)
    manifest = []
    for repo, commit, source, digest in LEDGER_FEATURES:
        url = f'https://raw.githubusercontent.com/{repo}/{commit}/{source}'
        file = dest / Path(source).name
        data = file.read_bytes() if file.exists() else urllib.request.urlopen(url, timeout=60).read()
        if hashlib.sha256(data).hexdigest() != digest:
            raise ValueError(f'Fixture hash mismatch: {file}')
        if not file.exists(): file.write_bytes(data)
        manifest.append(dict(file=file.name, repo=repo, commit=commit, path=source, url=url, sha256=digest))
    (dest / 'manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
    print(f'ledger-features: {len(manifest)} pinned supplemental files (separate from the main corpus)')

def git(folder, *args, check=True):
    p = subprocess.run(['git', '-C', str(folder), *args], check=check, stdout=subprocess.PIPE, text=True)
    return p.stdout.strip() if p.returncode == 0 else None

def legacy(root):
    """Materialize only binary fixtures, without changing the existing sparse checkouts.

    The Git tree at each existing pin is the inventory. Git's blob digest verifies
    downloaded bytes; SHA-256 in the separate manifest makes later runs auditable.
    Destination directories may be symlinks to a larger filesystem. No source code
    from the LibreOffice checkout is opened.
    """
    from concurrent.futures import ThreadPoolExecutor
    import time
    from urllib.parse import quote
    kinds = {'.doc': 'word', '.dot': 'word', '.xls': 'excel', '.xlt': 'excel',
             '.ppt': 'powerpoint', '.pps': 'powerpoint', '.pot': 'powerpoint'}
    entries = []
    for name in ('poi', 'libreoffice'):
        url, commit, folders = SOURCES[name]
        folder = root / name
        if git(folder, 'rev-parse', '--verify', 'HEAD') != commit:
            raise ValueError(f'{name}: prepare the existing corpus at {commit} first')
        if git(folder, 'remote', 'get-url', 'origin').removesuffix('.git').rstrip('/') != url:
            raise ValueError(f'Unexpected remote in {folder}')
        tree = subprocess.check_output(['git', '-C', str(folder), 'ls-tree', '-rz', commit, '--', *folders])
        for line in tree.split(b'\0'):
            if not line: continue
            meta, path = line.split(b'\t', 1)
            mode, kind, oid = meta.decode().split()
            rel = path.decode('utf-8')
            ext = Path(rel).suffix.lower()
            if ext not in kinds or kind != 'blob' or mode not in ('100644', '100755'): continue
            key = hashlib.sha256(path).hexdigest()[:12]
            dest = f'legacy-{kinds[ext]}/{name}__{key}__{Path(rel).name}'
            raw = url.replace('https://github.com/', 'https://raw.githubusercontent.com/')
            entries.append(dict(source=name, commit=commit, path=rel, file=dest,
                                format=kinds[ext], extension=ext, git_blob=oid,
                                url=f'{raw}/{commit}/{quote(rel)}'))
    for app in set(kinds.values()): (root / ('legacy-' + app)).mkdir(exist_ok=True)

    def fetch(entry):
        target = root / entry['file']
        original = root / entry['source'] / entry['path']
        if target.is_file(): data = target.read_bytes()
        elif original.is_file(): data = original.read_bytes()
        else:
            for attempt in range(4):
                try:
                    with urllib.request.urlopen(entry['url'], timeout=90) as response:
                        data = response.read(256 * 1024 * 1024 + 1)
                    break
                except Exception:
                    if attempt == 3: raise
                    time.sleep(2 ** attempt)
        if len(data) > 256 * 1024 * 1024: raise ValueError(f'Fixture too large: {target}')
        oid = hashlib.sha1(b'blob ' + str(len(data)).encode() + b'\0' + data).hexdigest()
        if oid != entry['git_blob']: raise ValueError(f'Git blob mismatch: {target}')
        if not target.exists():
            temp = target.with_name(target.name + '.part')
            temp.write_bytes(data)
            temp.replace(target)
        return dict(entry, size=len(data), sha256=hashlib.sha256(data).hexdigest())

    manifest = []
    with ThreadPoolExecutor(max_workers=6) as pool:
        for entry in pool.map(fetch, sorted(entries, key=lambda e: e['file'])):
            manifest.append(entry)
            if len(manifest) % 100 == 0: print(f'legacy: {len(manifest)}/{len(entries)} verified', flush=True)
    for dest, rows in [(root / 'legacy-manifest.json', manifest)] + [
            (root / ('legacy-' + app) / 'manifest.json', [e for e in manifest if e['format'] == app])
            for app in sorted(set(kinds.values()))]:
        temp = dest.with_name(dest.name + '.tmp')
        temp.write_text(json.dumps(rows, indent=2) + '\n')
        temp.replace(dest)
    for app in sorted(set(kinds.values())):
        print(f'legacy-{app}: {sum(e["format"] == app for e in manifest)} files')

def main():
    root = Path(sys.argv[1] if len(sys.argv) > 1 else os.environ.get('CORPORA', '/work/corpora')).expanduser().resolve()
    repo = Path(__file__).resolve().parents[2]
    if root == repo or repo in root.parents:
        raise ValueError('Choose a corpus directory outside the repository')
    selected = sys.argv[2:] or list(SOURCES)
    if selected == ['ledger-features']:
        ledger_features(root)
        return
    if selected == ['legacy']:
        legacy(root)
        return
    if any(s not in SOURCES for s in selected):
        raise ValueError('Sources: ' + ', '.join(SOURCES))
    root.mkdir(parents=True, exist_ok=True)
    for name in selected:
        url, commit, folders = SOURCES[name]
        folder = root / name
        folder.mkdir(exist_ok=True)
        if not (folder / '.git').exists():
            if any(folder.iterdir()):
                raise ValueError(f'Refusing to replace nonempty directory {folder}')
            git(folder, 'init', '-q')
            git(folder, 'remote', 'add', 'origin', url)
        if git(folder, 'remote', 'get-url', 'origin').removesuffix('.git').rstrip('/') != url:
            raise ValueError(f'Unexpected remote in {folder}')
        if git(folder, 'status', '--porcelain'):
            raise ValueError(f'Uncommitted files in {folder}; use another corpus directory')
        if git(folder, 'rev-parse', '--verify', 'HEAD', check=False) != commit:
            exts = sorted({k[1:] for k in BY_EXT} | {k[1:].upper() for k in BY_EXT})
            patterns = [f'/{d}/**/*.{ext}' for d in folders for ext in exts]
            git(folder, 'sparse-checkout', 'set', '--no-cone', *patterns)
            print(f'Fetching {name} at {commit}', flush=True)
            git(folder, 'fetch', '--depth=1', '--filter=blob:none', 'origin', commit)
            git(folder, 'checkout', '--detach', commit)
        print(f'{name}: {commit}', flush=True)
    manifest = []
    for name, (_, commit, _) in SOURCES.items():
        folder = root / name
        if not (folder / '.git').exists() or git(folder, 'rev-parse', '--verify', 'HEAD', check=False) != commit:
            continue
        files = subprocess.check_output(['git', '-C', str(folder), 'ls-files', '-z']).decode().split('\0')
        for rel in sorted(files):
            file = folder / rel
            app = BY_EXT.get(file.suffix.lower())
            if not app or not file.is_file():
                continue
            data = file.read_bytes()
            key = hashlib.sha256(rel.encode()).hexdigest()[:12]
            link = root / app / f'{name}__{key}__{file.name}'
            link.parent.mkdir(exist_ok=True)
            target = Path('..') / name / rel
            if link.is_symlink():
                if os.readlink(link) != str(target):
                    raise ValueError(f'Unexpected symlink: {link}')
            elif link.exists():
                raise ValueError(f'Refusing to replace {link}')
            else:
                link.symlink_to(target)
            manifest.append(dict(source=name, commit=commit, path=rel, file=str(link.relative_to(root)),
                                 size=len(data), sha256=hashlib.sha256(data).hexdigest()))
    tmp = root / f'manifest.{os.getpid()}.tmp'
    tmp.write_text(json.dumps(manifest, indent=2) + '\n')
    tmp.replace(root / 'manifest.json')
    for app in EXTENSIONS:
        print(f'{app}: {sum(m["file"].startswith(app + "/") for m in manifest)} files')

if __name__ == '__main__':
    main()

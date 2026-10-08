"""Fetch pinned, sparse public corpora; never delete or reset an existing checkout."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

SOURCES = {
    'libreoffice': ('https://github.com/LibreOffice/core', '2821d29a87b28785d74fa64d975465e4c99d4416', ['sw/qa', 'sc/qa', 'sd/qa', 'oox/qa']),
    'poi': ('https://github.com/apache/poi', 'ae62bb5116b9aee19ebd5834e3a82066132c9f7f', ['test-data']),
    'openxml-sdk': ('https://github.com/dotnet/Open-XML-SDK', '431ab05cf160248cc3885a4a766026d4f8243792', ['test']),
}
EXTENSIONS = {'word': ['docx', 'docm', 'dotx', 'dotm'], 'excel': ['xlsx', 'xlsm', 'xltx', 'xltm'],
              'powerpoint': ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm']}
BY_EXT = {'.' + ext: app for app, exts in EXTENSIONS.items() for ext in exts}

def git(folder, *args, check=True):
    p = subprocess.run(['git', '-C', str(folder), *args], check=check, stdout=subprocess.PIPE, text=True)
    return p.stdout.strip() if p.returncode == 0 else None

def main():
    root = Path(sys.argv[1] if len(sys.argv) > 1 else os.environ.get('CORPORA', str(Path.home() / 'corpora'))).expanduser().resolve()
    repo = Path(__file__).resolve().parents[2]
    if root == repo or repo in root.parents:
        raise ValueError('Choose a corpus directory outside the repository')
    selected = sys.argv[2:] or list(SOURCES)
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

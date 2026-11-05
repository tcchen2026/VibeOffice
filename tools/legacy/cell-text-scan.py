#!/usr/bin/env python3
"""Locate Excel text-limit candidates in converted and existing OOXML corpora.

This is evidence for the Office acceptance rule, not a replacement package checker.
The Office controls cover a referenced shared string; unused strings are also
listed here so the rule's owner can decide how to handle them.
"""
import collections
import json
import pathlib
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
LIMIT = 32767
PART_LIMIT = 128 * 1024 * 1024


def length(value):
    value = re.sub(r'_x([0-9a-fA-F]{4})_', lambda m: chr(int(m[1], 16)), value)
    return len(value.encode('utf-16-le', errors='surrogatepass')) // 2


def rich_text(node):
    return ''.join(n.text or '' for n in node.findall(NS + 't') + node.findall(NS + 'r/' + NS + 't'))


def scan(groups):
    files = collections.Counter(); oversized = []; unreadable = []
    for group, folder in groups:
        for path in sorted(pathlib.Path(folder).iterdir()):
            if path.suffix.lower() not in {'.xlsx', '.xlsm', '.xltx', '.xltm'}:
                continue
            files[group] += 1
            try:
                with zipfile.ZipFile(path) as z:
                    for part in z.infolist():
                        shared = part.filename == 'xl/sharedStrings.xml'
                        sheet = part.filename.startswith('xl/worksheets/') and part.filename.endswith('.xml')
                        if not (shared or sheet):
                            continue
                        if part.file_size > PART_LIMIT:
                            unreadable.append({'file': str(path), 'part': part.filename, 'reason': 'part exceeds scan bound'})
                            continue
                        index = 0
                        with z.open(part) as source:
                            for _, node in ET.iterparse(source, events=('end',)):
                                if shared and node.tag == NS + 'si':
                                    size = length(rich_text(node)); where = {'index': index}; index += 1
                                elif sheet and node.tag == NS + 'c':
                                    if node.get('t') == 'inlineStr':
                                        item = node.find(NS + 'is'); size = length(rich_text(item)) if item is not None else 0
                                    elif node.get('t') == 'str':
                                        size = length(node.findtext(NS + 'v') or '')
                                    else:
                                        size = 0
                                    where = {'cell': node.get('r')}
                                else:
                                    if node.tag == NS + 'row':
                                        node.clear()
                                    continue
                                if size > LIMIT:
                                    oversized.append({'group': group, 'file': str(path), 'part': part.filename, **where, 'length': size})
                                node.clear()
            except (OSError, ValueError, RuntimeError, zipfile.BadZipFile, ET.ParseError) as error:
                unreadable.append({'file': str(path), 'reason': str(error)})
    return {'files': dict(files), 'oversized': oversized, 'unreadable': unreadable}


if __name__ == '__main__':
    folders = sys.argv[1:] or ['tmp/legacy/artifacts/C1/packages', '/work/corpora/excel']
    if len(folders) != 2:
        raise SystemExit('usage: cell-text-scan.py CONVERTED_PACKAGES OOXML_CORPUS')
    print(json.dumps(scan(zip(['converted', 'ooxml'], folders)), indent=2))

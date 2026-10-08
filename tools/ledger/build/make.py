#!/usr/bin/env python3
"""Assemble public/ledger/index.html: a full document with Ledger's CSS inlined (after the shared
public/common/luna.css) and the scripts in load order.

    python3 tools/ledger/build/make.py
"""
import os
D = os.path.dirname(os.path.abspath(__file__))
APP = os.path.join(D, '..', '..', '..', 'public', 'ledger')
# load order; 'common/x' is public/common/x.js, shared by the suite, anything else public/ledger/js/x.js
SCRIPTS = ['common/core', 'common/xml', 'common/opc', 'common/opc-order', 'xml', 'common/zip', 'common/sha', 'common/crypto', 'common/icons',
           'common/geometry', 'common/metafile', 'common/ui', 'common/numfmt', 'formula', 'model', 'calc',
           'fn-core', 'fn-lookup', 'fn-stat', 'fn-fin', 'fn-eng', 'common/dml', 'common/charts', 'xchart',
           'preserve', 'xlsx-read', 'xlsx-write', 'csv', 'xmlss', 'styles', 'cf', 'layout', 'render', 'ops', 'clipboard',
           'fninfo', 'grid', 'editor', 'common/spell', 'spell', 'filter', 'drawing', 'commands', 'dialogs',
           'dialogs2', 'dialogs3', 'print', 'panes', 'app']
tpl = open(os.path.join(D, 'template.html')).read()
css = open(os.path.join(D, 'ledger.css')).read()   # after public/common/luna.css, linked by the template
def src(s):
    rel = '../' + s + '.js' if s.startswith('common/') else 'js/' + s + '.js'
    if not os.path.exists(os.path.join(APP, rel)):
        raise SystemExit(f'missing script: {os.path.relpath(os.path.join(APP, rel))}')
    return rel
tags = '\n'.join(f'<script src="{src(s)}"></script>' for s in SCRIPTS)
body = tpl.replace('/*CSS*/', css).replace('<!--SCRIPTS-->', tags)
out = ('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
       '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n' + body + '\n</html>\n')
open(os.path.join(APP, 'index.html'), 'w').write(out)
print('public/ledger/index.html', len(out), 'bytes;', tags.count('<script'), 'scripts')

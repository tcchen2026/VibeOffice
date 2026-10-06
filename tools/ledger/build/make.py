#!/usr/bin/env python3
"""Assemble public/ledger/index.html: a full document with the CSS inlined and the scripts in load order.

    python3 tools/ledger/build/make.py
"""
import os
D = os.path.dirname(os.path.abspath(__file__))
APP = os.path.join(D, '..', '..', '..', 'public', 'ledger')
SCRIPTS = ['core', 'xml', 'zip', 'sha', 'crypto', 'icons', 'geometry', 'metafile', 'ui',
           'numfmt', 'formula', 'model', 'calc', 'fn-core', 'fn-lookup', 'fn-stat', 'fn-fin', 'fn-eng',
           'dml', 'charts', 'xchart', 'xlsx-read', 'xlsx-write', 'csv', 'xmlss', 'styles', 'cf', 'layout', 'render', 'ops', 'clipboard',
           'fninfo', 'grid', 'editor', 'spell', 'filter', 'drawing', 'commands', 'dialogs', 'dialogs2', 'dialogs3', 'print', 'panes', 'app']
tpl = open(os.path.join(D, 'template.html')).read()
css = open(os.path.join(D, 'base.css')).read() + '\n' + open(os.path.join(D, 'ledger.css')).read()
tags = '\n'.join(f'<script src="js/{s}.js"></script>' for s in SCRIPTS if os.path.exists(os.path.join(APP, 'js', s + '.js')))
body = tpl.replace('/*CSS*/', css).replace('<!--SCRIPTS-->', tags)
out = ('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
       '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n' + body + '\n</html>\n')
open(os.path.join(APP, 'index.html'), 'w').write(out)
print('public/ledger/index.html', len(out), 'bytes;', tags.count('<script'), 'scripts')

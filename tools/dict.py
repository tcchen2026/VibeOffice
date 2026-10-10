#!/usr/bin/env python3
"""Rebuild public/common/dict/en_GB.words: the British English spelling list (Quire, Ledger).

    python3 tools/dict.py          # needs bash and perl (SCOWL's mk-list)

Downloads SCOWL 2020.12.07 (pinned by SHA-256) and takes SCOWL's own en_GB-ise list at size 60 with
accents stripped: the list SCOWL's Hunspell en_GB-ise dictionary is made from, and the British
counterpart of en_US.words (SCOWL's en_US list at size 60, byte for byte; the script checks that first).
SCOWL's licence is permissive (public/common/dict/LICENSES.txt).

Commonness levels (the last digit of each line, used to rank suggestions): a word that is also in
en_US.words keeps its level there (from the English Speller Database); a British-only spelling gets the
level its smallest SCOWL size most often has in en_US.words (35 or less: 2, 40: 3, 50: 4, 55 and 60: 6).

Format, as common/spell.js reads it: sorted as SCOWL lists it, one word per line, front-coded as
<length of the prefix shared with the previous word, 0-9a-z><the rest><level digit>.
"""
import hashlib
import os
import subprocess
import sys
import tarfile
import tempfile
import urllib.request

URL = 'https://downloads.sourceforge.net/project/wordlist/SCOWL/2020.12.07/scowl-2020.12.07.tar.gz'
SHA256 = '5587667caa20c4891390c2d42dbb4d5c4c3f41bee77af1457ece3ba23fb859cc'
SIZES = [10, 20, 35, 40, 50, 55, 60]
LEVEL_OF_SIZE = {10: 2, 20: 2, 35: 2, 40: 3, 50: 4, 55: 6, 60: 6}
P = '0123456789abcdefghijklmnopqrstuvwxyz'
DICT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'public', 'common', 'dict')


def decode(text):
    out, prev = {}, ''
    for ln in text.split('\n'):
        if ln:
            w = prev[:P.index(ln[0])] + ln[1:-1]
            out[w] = int(ln[-1])
            prev = w
    return out


def encode(words, level):
    lines, prev = [], ''
    for w in words:
        k = 0
        while k < min(len(w), len(prev), len(P) - 1) and w[k] == prev[k]:
            k += 1
        lines.append(P[k] + w[k:] + str(level[w]))
        prev = w
    return '\n'.join(lines) + '\n'


def mk_list(scowl, spelling, size):
    out = subprocess.run(['./mk-list', '-d', 'final', '--accents=strip', spelling, str(size)], cwd=scowl,
                         check=True, capture_output=True).stdout.decode('ascii')
    return [w for w in out.split('\n') if w]


def main():
    with tempfile.TemporaryDirectory() as tmp:
        data = urllib.request.urlopen(URL).read()
        if hashlib.sha256(data).hexdigest() != SHA256:
            sys.exit('SCOWL download does not match the pinned SHA-256')
        path = os.path.join(tmp, 'scowl.tar.gz')
        with open(path, 'wb') as f:
            f.write(data)
        with tarfile.open(path) as t:
            t.extractall(tmp, filter='data')
        scowl = os.path.join(tmp, 'scowl-2020.12.07')

        us_file = open(os.path.join(DICT, 'en_US.words'), encoding='ascii').read()
        us = decode(us_file)
        us_words = mk_list(scowl, 'en_US', 60)
        if list(us) != us_words or encode(us_words, us) != us_file:
            sys.exit('en_US.words is no longer SCOWL en_US size 60: check the source before building en_GB')

        size = {}
        for s in reversed(SIZES):
            for w in mk_list(scowl, 'en_GB-ise', s):
                size[w] = s
        gb = mk_list(scowl, 'en_GB-ise', 60)
        level = {w: us.get(w, LEVEL_OF_SIZE[size[w]]) for w in gb}
        with open(os.path.join(DICT, 'en_GB.words'), 'w', encoding='ascii', newline='\n') as f:
            f.write(encode(gb, level))
        print(f'en_GB.words: {len(gb)} words, {sum(1 for w in gb if w not in us)} British-only')


if __name__ == '__main__':
    main()

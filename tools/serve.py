#!/usr/bin/env python3
"""Local dev server for public/ that never lets the browser cache anything.

    python3 tools/serve.py [port]      (default 8765, bound to 127.0.0.1)

Every response carries Cache-Control: no-store, and conditional requests are
ignored, so edited HTML/JS modules are always re-fetched on reload.

POST /__feedback takes the F11 bug reports of myvr games (public/common/myvr/app/feedback.js)
and writes issues/<yyyymmdd>/<hhmmss>-<game>-<words>.json and .png (with an empty "solution",
filled in when the issue is fixed and the pair moves to issues/<yyyymmdd>/solved/). Only same-origin requests
from 127.0.0.1 are accepted; file names are built here from [a-z0-9-] only (at most 50 chars).
"""
import base64
import binascii
import datetime
import http.server
import json
import os
import re
import sys

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'public')
ISSUES = os.path.realpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'issues'))
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8760
MAX_BODY = 32 * 1024 * 1024
NAME_MAX = 50                      # file name length, extension included
PNG_PREFIX = 'data:image/png;base64,'


class FeedbackError(Exception):
    def __init__(self, code, msg):
        super().__init__(msg)
        self.code = code


def slug(text):
    """lower case, every run of characters outside a-z0-9 becomes one '-'"""
    return re.sub(r'[^a-z0-9]+', '-', text.lower()).strip('-')


def save_feedback(body, now=None):
    """Validate a report and write it; returns the paths relative to the repo."""
    if not isinstance(body, dict):
        raise FeedbackError(400, 'expected a JSON object')
    comment, game = body.get('comment'), body.get('game')
    data, marks, image = body.get('data'), body.get('marks', []), body.get('image')
    if not isinstance(comment, str) or not comment.strip() or len(comment) > 2000:
        raise FeedbackError(400, 'comment: 1..2000 characters')
    if not isinstance(game, str) or len(game) > 100:
        raise FeedbackError(400, 'game: a short string')
    if data is not None and not isinstance(data, dict):
        raise FeedbackError(400, 'data: an object')
    if not isinstance(marks, list) or len(marks) > 100:
        raise FeedbackError(400, 'marks: a short list')
    png = None
    if image is not None:
        if not isinstance(image, str) or not image.startswith(PNG_PREFIX):
            raise FeedbackError(400, 'image: a PNG data URL')
        try:
            png = base64.b64decode(image[len(PNG_PREFIX):], validate=True)
        except (binascii.Error, ValueError):
            raise FeedbackError(400, 'image: bad base64')
        if not png.startswith(b'\x89PNG\r\n\x1a\n'):
            raise FeedbackError(400, 'image: not a PNG')

    now = now or datetime.datetime.now()
    folder = os.path.join(ISSUES, now.strftime('%Y%m%d'))
    game_s = slug(game)[:20].strip('-') or 'game'
    words = slug(' '.join(comment.split()[:8])) or 'issue'
    base = f"{now.strftime('%H%M%S')}-{game_s}-{words}"
    room = NAME_MAX - len('.json')
    os.makedirs(folder, exist_ok=True)
    for n in range(1, 100):
        suffix = '' if n == 1 else f'-{n}'
        stem = base[:room - len(suffix)].strip('-') + suffix
        paths = [os.path.realpath(os.path.join(folder, stem + ext)) for ext in ('.json', '.png')]
        if any(os.path.dirname(p) != os.path.realpath(folder) or not p.startswith(ISSUES + os.sep) for p in paths):
            raise FeedbackError(400, 'bad file name')
        record = {
            'time': now.isoformat(timespec='seconds'), 'game': game, 'comment': comment.strip(), 'solution': '',
            'screenshot': stem + '.png' if png else None, 'marks': marks, 'data': data,
        }
        try:
            with open(paths[0], 'x', encoding='utf-8') as f:     # 'x': never overwrite a report
                json.dump(record, f, indent=1, ensure_ascii=False)
        except FileExistsError:
            continue
        if png:
            with open(paths[1], 'wb') as f:
                f.write(png)
        rel = lambda p: os.path.relpath(p, os.path.dirname(ISSUES))
        return {'ok': True, 'json': rel(paths[0]), 'png': rel(paths[1]) if png else None}
    raise FeedbackError(500, 'too many reports with this name')


class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map,
                      '.js': 'text/javascript', '.mjs': 'text/javascript',
                      '.glb': 'model/gltf-binary', '.wasm': 'application/wasm'}

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

    def send_head(self):
        # Drop validators so we always answer 200 with a full body instead of 304.
        for h in ('If-Modified-Since', 'If-None-Match'):
            if h in self.headers:
                del self.headers[h]
        return super().send_head()

    def do_POST(self):
        if self.path.split('?')[0] != '/__feedback':
            self.send_error(404)
            return
        try:
            out = save_feedback(self._feedback_body())
            print(f"feedback saved: {out['json']}", flush=True)
            self._reply(200, out)
        except FeedbackError as e:
            self._reply(e.code, {'ok': False, 'error': str(e)})

    def _feedback_body(self):
        # same machine, same origin: other web pages in the browser must not be able to post here
        if self.client_address[0] != '127.0.0.1':
            raise FeedbackError(403, 'only accepted from 127.0.0.1')
        here = f'127.0.0.1:{PORT}'
        if self.headers.get('Host') != here:
            raise FeedbackError(403, 'bad Host')
        origin = self.headers.get('Origin')
        if origin is not None and origin != f'http://{here}':
            raise FeedbackError(403, 'bad Origin')
        if self.headers.get('Content-Type', '').split(';')[0].strip() != 'application/json':
            raise FeedbackError(415, 'expected application/json')
        try:
            n = int(self.headers.get('Content-Length', ''))
        except ValueError:
            raise FeedbackError(411, 'missing Content-Length')
        if n < 0 or n > MAX_BODY:
            raise FeedbackError(413, 'report too large')
        try:
            return json.loads(self.rfile.read(n).decode('utf-8'))
        except (UnicodeDecodeError, json.JSONDecodeError):
            raise FeedbackError(400, 'bad JSON')

    def _reply(self, code, obj):
        data = json.dumps(obj).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, max-age=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()


if __name__ == '__main__':
    http.server.ThreadingHTTPServer.allow_reuse_address = True
    with http.server.ThreadingHTTPServer(('127.0.0.1', PORT), NoCacheHandler) as httpd:
        print(f'Serving {os.path.normpath(ROOT)} at http://127.0.0.1:{PORT}/ (no-cache)', flush=True)
        httpd.serve_forever()

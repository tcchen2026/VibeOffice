"""Private warm LibreOffice worker for render.py; one JSON request per line."""
import json
from pathlib import Path
import resource
import subprocess
import sys
import time
import uuid

import uno
import unohelper
from com.sun.star.beans import PropertyValue
from com.sun.star.task import XInteractionHandler


class AbortInteraction(unohelper.Base, XInteractionHandler):
    def handle(self, request):
        for kind in ('XInteractionAbort', 'XInteractionDisapprove'):
            for continuation in request.getContinuations():
                action = continuation.queryInterface(uno.getTypeByName('com.sun.star.task.' + kind))
                if action:
                    action.select()
                    return


def prop(name, value):
    result = PropertyValue()
    result.Name, result.Value = name, value
    return result


profile, memory = Path(sys.argv[1]), int(sys.argv[2]) * 1024 * 1024
resource.setrlimit(resource.RLIMIT_AS, (memory, memory))
resource.setrlimit(resource.RLIMIT_FSIZE, (512 * 1024 * 1024, 512 * 1024 * 1024))
pipe = 'vo_render_' + uuid.uuid4().hex
office = subprocess.Popen([
    'soffice', '-env:UserInstallation=' + profile.as_uri(), '--headless', '--nologo',
    '--norestore', '--nodefault', '--nofirststartwizard',
    '--accept=pipe,name=' + pipe + ';urp;StarOffice.ComponentContext',
], stdout=subprocess.DEVNULL, stderr=sys.stderr)
try:
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    deadline = time.monotonic() + 20
    while True:
        try:
            context = resolver.resolve('uno:pipe,name=' + pipe + ';urp;StarOffice.ComponentContext')
            break
        except Exception:
            if office.poll() is not None or time.monotonic() >= deadline:
                raise RuntimeError('LibreOffice worker did not start')
            time.sleep(0.05)
    desktop = context.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', context)
    print(json.dumps({'ready': True}), flush=True)
    for line in sys.stdin:
        request = json.loads(line)
        document = None
        try:
            document = desktop.loadComponentFromURL(uno.systemPathToFileUrl(request['source']), '_blank', 0, (
                prop('Hidden', True), prop('ReadOnly', True), prop('AsTemplate', False),
                prop('UpdateDocMode', 0), prop('MacroExecutionMode', 0),
                prop('InteractionHandler', AbortInteraction()),
            ))
            if document is None:
                raise ValueError('LibreOffice did not open the document')
            filters = {
                'com.sun.star.text.TextDocument': 'writer_pdf_Export',
                'com.sun.star.sheet.SpreadsheetDocument': 'calc_pdf_Export',
                'com.sun.star.presentation.PresentationDocument': 'impress_pdf_Export',
            }
            export = next((name for service, name in filters.items() if document.supportsService(service)), None)
            if not export:
                raise ValueError('Document did not open as Writer, Calc or Impress')
            document.storeToURL(uno.systemPathToFileUrl(request['pdf']), (
                prop('FilterName', export), prop('Overwrite', True),
            ))
            result = {'status': 'ok', 'filter': export}
        except Exception as error:
            message = str(error)
            result = {'status': 'failed', 'error': message,
                      'workerFailure': office.poll() is not None or 'bridge' in message.lower() and 'disposed' in message.lower()}
        finally:
            if document is not None:
                try:
                    document.close(True)
                except Exception:
                    try:
                        document.dispose()
                    except Exception as error:
                        # Export may already have succeeded. Report its outcome
                        # and retire this process before accepting another file.
                        result.update(workerRestart=True, closeError=str(error))
        print(json.dumps(result), flush=True)
finally:
    office.terminate()
    try:
        office.wait(timeout=5)
    except subprocess.TimeoutExpired:
        office.kill()

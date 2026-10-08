"""Attempt accounting shared by the feature inventories. Failed pairs never disappear."""
from pathlib import Path

class Audit:
    def __init__(self):
        self.results = []
        self.errors = []

    def pairs(self, corpus, saved, extensions):
        corpus, saved = Path(corpus), Path(saved)
        outputs = {}
        for f in sorted(saved.iterdir()):
            if f.suffix.lower() in extensions:
                outputs.setdefault(f.stem, []).append(f)
        for original in sorted(corpus.iterdir()):
            if original.suffix.lower() not in extensions:
                continue
            choices = outputs.get(original.stem, [])
            target = saved / original.name
            if not target.exists() and len(choices) == 1:
                target = choices[0]
            if not target.exists():
                self.results.append(dict(file=original.name, status='failed', reason='Missing saved copy'))
                continue
            self.results.append(dict(file=original.name, status='ok'))
            yield original.name, str(original), str(target)

    def failure(self, file, error, feature=None, excluded=False):
        self.errors.append(dict(file=file, status='excluded' if excluded else 'failed', feature=feature, error=str(error)))
        for result in reversed(self.results):
            if result['file'] == file:
                if result['status'] != 'failed':
                    result['status'] = 'excluded' if excluded else 'failed'
                break

    def report(self):
        return dict(attempted=len(self.results), ok=sum(r['status'] == 'ok' for r in self.results),
                    failed=sum(r['status'] == 'failed' for r in self.results),
                    excluded=sum(r['status'] == 'excluded' for r in self.results), results=self.results, errors=self.errors)

"""Aggregate the shared-layer fields of a completed measure.mjs --shared run."""
import argparse
from collections import Counter
import json
from pathlib import Path
p=argparse.ArgumentParser(description=__doc__);p.add_argument('run');p.add_argument('--out');args=p.parse_args()
rows=[json.loads(l) for l in (Path(args.run)/'runs.jsonl').read_text().splitlines()]
originals=[r for r in rows if not r['mutation']]
result={'cases':len(rows),'states':dict(Counter(r['status'] for r in rows)),'originals':len(originals),'formats':{}}
for kind in ['doc','xls','ppt','other']:
 selected=[r for r in originals if (r.get('sniff',{}).get('format') if r.get('sniff',{}).get('format') in ['doc','xls','ppt'] else 'other')==kind]
 sums=Counter();pictures=Counter();errors=Counter();geometry=Counter();issues=Counter();metadata=Counter(); files_with_metadata_issues=0
 for r in selected:
  s=r.get('shared',{});
  for k in ['properties','propertyParts','drawings','shapes','tables']:sums[k]+=s.get(k,0)
  pictures.update(s.get('pictures',{}));errors.update(s.get('pictureErrors',{}));geometry.update(s.get('geometry',{}));issues.update(i['code'] for i in s.get('issues',[]));metadata.update(i['code'] for i in s.get('propertyIssues',[]));files_with_metadata_issues+=bool(s.get('propertyIssues'))
 largest=max(selected,key=lambda r:r['size'],default=None)
 result['formats'][kind]={'files':len(selected),'shared_runs':sum('shared'in r for r in selected),'totals':dict(sums),'pictures':dict(pictures),'picture_errors':dict(errors),'geometry':dict(geometry),'issues':dict(issues),'metadata_issues':dict(metadata),'metadata_issue_files':files_with_metadata_issues,'largest':{k:largest[k] for k in ['file','size','ms']} if largest else None}
text=json.dumps(result,indent=2)+'\n';Path(args.out or str(Path(args.run)/'shared-summary.json')).write_text(text);print(text)

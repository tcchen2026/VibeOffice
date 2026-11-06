"""Compare legacy Word body text with independent piece-table and LO references.

python3 -I tools/legacy/doc-agreement.py RUN_DIR SOURCE_REFERENCES.json OUTPUT.json
The source comparison keeps all source text, including revisions. LO comparison
uses its current body text; text boxes, comments and note/header parts are separate.
"""
import collections,json,pathlib,re,sys,zipfile,xml.etree.ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def norm(text):
    return re.sub(r'\s+', '', text).replace('\u00ad','')
def source_text(text):
    result=[];fields=[]
    for ch in text:
        n=ord(ch)
        if n==19:fields.append(False);continue
        if n==20:
            if fields:fields[-1]=True
            continue
        if n==21:
            if fields:fields.pop()
            continue
        if not all(fields):continue
        if n==30:result.append('\u2011')
        elif n==31:result.append('\u00ad')
        elif n>=32 or ch in '\t\n\r':result.append(ch)
    return ''.join(result)
def xml_text(node):
    out=[]
    def walk(n):
        name=n.tag.removeprefix(W)
        if name in ('txbxContent','del','moveFrom','instrText','delInstrText'):return
        if name=='t':out.append(n.text or '')
        elif name in ('br','cr','tab'):out.append('\n' if name!='tab' else '\t')
        elif name=='noBreakHyphen':out.append('\u2011')
        elif name=='softHyphen':out.append('\u00ad')
        elif name=='sym':out.append(chr(int(n.get(W+'char','0'),16)))
        else:
            for c in n:walk(c)
        if name=='p':out.append('\n')
    walk(node);return ''.join(out)
def content(file):
    with zipfile.ZipFile(file) as z:
        result={}
        doc=ET.fromstring(z.read('word/document.xml'))
        result['main']=xml_text(doc.find(W+'body'))
        for part in ['footnotes','endnotes']:
            if 'word/'+part+'.xml' in z.namelist():
                root=ET.fromstring(z.read('word/'+part+'.xml'))
                result[part]=[xml_text(n) for n in root if n.get(W+'type','normal')=='normal']
        result['headers']=[xml_text(ET.fromstring(z.read(name))) for name in z.namelist() if re.fullmatch(r'word/(?:header|footer)\d+\.xml',name)]
        return result
run=pathlib.Path(sys.argv[1]);refs={r['sha256']:r for r in json.loads(pathlib.Path(sys.argv[2]).read_text())};output=pathlib.Path(sys.argv[3]);reference_root=pathlib.Path('tmp/legacy/artifacts/references');rows=[];counts=collections.Counter()
for line in (run/'runs.jsonl').read_text().splitlines():
    row=json.loads(line)
    if row.get('mutation') or row.get('sniff',{}).get('format')!='doc' or row['status']!='converted':continue
    h=row['sha256'];source=refs.get(h,{});converted=content(pathlib.Path(row.get('pkg') or run/'packages'/(h+'.'+row['type'])));record={'file':row['file'],'sha256':h,'converted':converted};counts['converted']+=1
    stories=source.get('references',{}).get('stories')
    if stories:
        expected=source_text(stories['main']);same=norm(expected)==norm(converted['main']);record['source']={'expected':expected,'same':same,'revisions':source.get('features',{}).get('revisionFlags')};counts['sourceCompared']+=1;counts['sourceEqual']+=same
    result=reference_root/h/'result.json'
    if result.exists():
        ref=json.loads(result.read_text())['conversions'].get('docx',{})
        if ref.get('status')=='converted':
            try:
                lo=content(reference_root/ref['file']);same=norm(lo['main'])==norm(converted['main']);record['libreoffice']={'same':same,'text':lo,'file':ref['file']};counts['loCompared']+=1;counts['loEqual']+=same
                for name in ['footnotes','endnotes','headers']:
                    a=sorted(norm(x) for x in converted.get(name,[]) if norm(x));b=sorted(norm(x) for x in lo.get(name,[]) if norm(x));counts[name+'Compared']+=1;counts[name+'Equal']+=a==b
                    # Producers may duplicate inherited header/footer parts.
                    # Keep both the raw part-multiset and distinct-text measures;
                    # neither by itself establishes section-by-section layout.
                    if name=='headers':counts['headerTextSetEqual']+=set(a)==set(b)
            except (zipfile.BadZipFile,ET.ParseError,KeyError) as e:record['referenceError']=str(e)
    rows.append(record)
output.write_text(json.dumps({'counts':counts,'rows':rows},ensure_ascii=False));print(json.dumps(counts))

// Exercise the shared property/OfficeArt layers independently of unfinished host writers.
import { createHash } from 'node:crypto';
export async function auditShared(B, bytes, sniff) {
  const out = { properties: 0, propertyIssues: [], propertyParts: 0, drawings: 0, shapes: 0, tables: 0, pictures: {}, pictureErrors: {}, geometry: {}, issues: [] };
  if (bytes[0] !== 0xd0 || bytes[1] !== 0xcf) return out;
  const cfb = B.cfb(bytes), A = B.officeart;
  const attempt = async (label, fn) => { try { await fn(); } catch (e) { if (!['LegacyError','CFBError'].includes(e.name)) throw e; out.issues.push({ label, code: e.code }); } };
  await attempt('properties', () => {
    const props = B.propertyParts(cfb); out.propertyParts = props.parts.length;
    for (const set of [...props.summary, ...props.document]) { out.properties += Object.keys(set.values).length; out.propertyIssues.push(...set.issues); }
    out.propertyPartHashes = Object.fromEntries(props.parts.map(p => [p.name,createHash('sha256').update(p.data).digest('hex')]));
  });
  if (sniff.encrypted) return out;
  const groups = []; let delayed;
  if (sniff.format === 'xls') {
    await attempt('BIFF drawings', () => {
      const records = B.xls.records(cfb.stream('Workbook') || cfb.stream('Book')), buckets = new Map(); let current = 'globals', matching = '';
      for (const rec of records) {
        if (rec.id === 0x809) current = 'substream:' + rec.offset;
        const key = rec.id === 0xeb ? 'drawingGroup' : rec.id === 0xec ? current : rec.id === 0x3c ? matching : '';
        if (key) { if (!buckets.has(key)) buckets.set(key, []); buckets.get(key).push(rec.r.bytes); } matching = key;
      }
      for (const [name, chunks] of buckets) {
        const size = chunks.reduce((n,b) => n+b.length,0), joined = new Uint8Array(size); let p=0;
        for (const b of chunks) { joined.set(b,p);p+=b.length; } groups.push({name,bytes:joined});
      }
    });
  } else if (sniff.format === 'doc' && sniff.version === 'Word97+') {
    await attempt('Word drawings', () => {
      const fib = B.doc.fib(cfb.stream('WordDocument')), dgg = fib.pair(50); if (!dgg.length) return;
      const r = new B.Reader(cfb.stream(fib.table)).reader(dgg.offset,dgg.length);let p=0;
      while(p<r.length) { if(p) p++; const rec=A.record(r,p);groups.push({name:'drawing:'+p,bytes:r.sub(p,rec.end-p)});p=rec.end; }
      delayed=[cfb.stream('Data'), cfb.stream('WordDocument')];
    });
  } else if(sniff.format === 'ppt' && sniff.version === 'PowerPoint97+') {
    await attempt('PowerPoint drawings', () => {
      const ppt=B.ppt.read(cfb), seen=new Set();
      for(const rec of [ppt.document,...ppt.slides.map(s=>s.record),...ppt.masters.map(s=>s.record),...ppt.notes.map(s=>s.record)]) {
        B.ppt.walk(rec, node=> { if ([1035,1036].includes(node.type) && !seen.has(node.r.bytes.byteOffset)) {seen.add(node.r.bytes.byteOffset);groups.push({name:'drawing:'+node.r.bytes.byteOffset,bytes:node.r.bytes});} });
      }
      delayed=cfb.stream('Pictures');
    });
  }
  for (const group of groups) await attempt(group.name, async () => {
    const records=A.records(group.bytes);out.drawings++;
    for (const entry of A.store(records,delayed)) {
      try {const media=await entry.read();out.pictures[media.ext]=(out.pictures[media.ext]||0)+1;}
      catch(e){if(!['LegacyError','CFBError'].includes(e.name))throw e;out.pictureErrors[e.code]=(out.pictureErrors[e.code]||0)+1;}
    }
    const visit=shapes=>{for(const shape of shapes){out.shapes++;if(shape.table)out.tables++; if(!(shape.flags&1)) {
      try { B.shapes.geometry(shape);out.geometry.ok=(out.geometry.ok||0)+1; }
      catch(e){if(e.name!=='LegacyError')throw e;out.geometry[e.code]=(out.geometry[e.code]||0)+1;}
    } visit(shape.children);}};
    visit(A.shapes(records,sniff.format));
  });
  return out;
}

import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { load } from '../../tools/legacy/load.mjs';
import { rec, container, u16, u32, bytes, variant, properties, target, condition,
  timeNode, setNode, setBehavior, keyframeNode, timeline, oldAnimation, shape, deck, linkedVideo, sound
} from './ppt-animation-fixtures.mjs';

const B = load(), P = B.ppt;
if (process.env.LEGACY_BEFORE_D4) {
  delete P.timing;
  delete P.media;
  for (const name of ['ppt-shapes', 'ppt-links', 'ppt-features', 'ppt-write']) {
    const file = process.env.LEGACY_BEFORE_D4 + '/' + name + '.js';
    vm.runInThisContext(fs.readFileSync(file, 'utf8'), { filename: file });
  }
}
async function convert(fixture) {
  const result = await P.convert(fixture.bytes), zip = await L.zip.read(result.pkg);
  return { ...result, zip, xml: await zip.get('ppt/slides/slide1.xml').text() };
}
const nodes = (xml, name) => {
  const found = [];
  B.checkXml(xml, node => { if (node.name === name) found.push(node); });
  return found;
};
const effect = (id, role = 1) => timeNode([properties([[9, 1], [11, 1], [20, role]]), setNode(id)],
  { fill: 3 });
const animationLoss = result => result.losses.find(l => l.id === 'legacy-slide-animation' && l.notify);

test('PPT extended timing keeps source order, targets, repeat and click conditions', async () => {
  const modifier = Buffer.alloc(8);
  modifier.writeFloatLE(2.5, 4);
  const fixture = deck([shape(2050), shape(2051)], [timeline([
    timeNode([condition(1, 0, -1), rec(0xf129, modifier), effect(2051), effect(2050, 2)], { fill: 3 })])]);
  const r = await convert(fixture);
  assert.deepEqual(nodes(r.xml, 'p:spTgt').map(n => n.attrs.spid), ['2051', '2050']);
  assert.match(r.xml, /repeatCount="2500"/);
  assert.match(r.xml, /<p:cond delay="indefinite"\/>/);
  assert.deepEqual(nodes(r.xml, 'p:cTn').filter(n => n.attrs.presetID).map(n => n.attrs.nodeType),
    ['clickEffect', 'withEffect']);
  assert.ok(!animationLoss(r));
});

test('PPT extended timing takes precedence over its older compatibility animation', async () => {
  const r = await convert(deck([shape(2050, { client: [oldAnimation({ effect: 5 })] })],
    [timeline([effect(2050)])]));
  assert.equal(nodes(r.xml, 'p:set').length, 1);
  assert.equal(nodes(r.xml, 'p:animEffect').length, 0);
  assert.ok(!animationLoss(r));
});

test('PPT the Office float sentinel means indefinite repeat, not damaged timing', async () => {
  const modifier = bytes(u32(0), u32(0x7f7fffff));
  const r = await convert(deck([shape(2050)], [timeline([
    setNode(2050, { extra: [rec(0xf129, modifier)] })])]));
  assert.match(r.xml, /repeatCount="indefinite"/);
  assert.ok(!animationLoss(r));
});

test('PPT old builds sort by source order and skip empty paragraph clicks as Office does', async () => {
  const r = await convert(deck([
    shape(2050, { text: 'First\r\rThird\r  ', client: [oldAnimation({ order: 2, build: 2, effect: 5 })] }),
    shape(2051, { client: [oldAnimation({ order: 1, effect: 12, direction: 0 })] })
  ]));
  const targets = nodes(r.xml, 'p:spTgt').map(n => n.attrs.spid);
  assert.equal(targets[0], '2051');
  assert.equal(nodes(r.xml, 'p:cTn').filter(n => n.attrs.nodeType === 'clickEffect').length, 3);
  assert.deepEqual(nodes(r.xml, 'p:pRg').map(n => [n.attrs.st, n.attrs.end]),
    [['0', '0'], ['0', '0'], ['2', '2'], ['2', '2']]);
  assert.match(r.xml, /<p:bldP spid="2050"/);
  assert.ok(!animationLoss(r));
});

test('PPT old reverse builds retain paragraph indices and automatic delay', async () => {
  const r = await convert(deck([shape(2050, { text: 'First\rSecond',
    client: [oldAnimation({ build: 2, flags: 5, delay: 1200 })] })]));
  assert.deepEqual(nodes(r.xml, 'p:pRg').map(n => n.attrs.st), ['1', '0']);
  assert.match(r.xml, /rev="1"/);
  assert.match(r.xml, /delay="1200"/);
  assert.equal(nodes(r.xml, 'p:cond').filter(n => n.attrs.delay === '499').length, 2);
  assert.equal(nodes(r.xml, 'p:cTn').filter(n => n.attrs.nodeType === 'afterEffect').length, 2);
});

test('PPT 54 Office-authored old effects keep native presets, directions and filters', async () => {
  const measured = JSON.parse(fs.readFileSync(new URL('./ppt-animation-office.json', import.meta.url)));
  const r = await convert(deck(measured.effects.map((row, i) => shape(2050 + i, {
    text: row.name, client: [oldAnimation({ order: i + 1, effect: row.old[0], direction: row.old[1] })]
  }))));
  assert.deepEqual(nodes(r.xml, 'p:cTn').filter(n => n.attrs.presetID)
    .map(n => [n.attrs.presetID, n.attrs.presetSubtype]), measured.effects.map(row => row.preset));
  assert.deepEqual(nodes(r.xml, 'p:animEffect').map(n => ({ ...n.attrs })),
    measured.effects.flatMap(row => row.filters));
  for (const [kind, count] of Object.entries(measured.behaviorCounts))
    assert.equal(nodes(r.xml, 'p:' + kind).length, count, kind + ' behavior count from Office');
  assert.equal(nodes(r.xml, 'p:bldP').filter(n => n.attrs.animBg === '1').length, 54);
  assert.ok(!animationLoss(r));
});

test('PPT unspecified keyframe time stays unspecified, as Office converts Novell 6655408', async () => {
  const r = await convert(deck([shape(2050)], [timeline([
    keyframeNode(2050, -1), keyframeNode(2050, -1000), keyframeNode(2050, 1000)])]));
  assert.deepEqual(nodes(r.xml, 'p:tav').map(n => n.attrs.tm), [undefined, 'indefinite', '100000']);
  assert.ok(!animationLoss(r));
});

test('PPT subordinate hide effects follow the next click, as Office converts 23884', async () => {
  const sub = timeNode([properties([[5, 2], [13, 1]]), setBehavior(2050)], { kind: 3, duration: 1 });
  sub.writeUInt16LE(0xf145, 2);
  const r = await convert(deck([shape(2050)], [timeline([
    timeNode([effect(2050), sub])])]));
  const after = nodes(r.xml, 'p:cTn').find(n => n.attrs.afterEffect === '1');
  assert.equal(after?.attrs.masterRel, 'nextClick');
  assert.ok(!animationLoss(r));
});

test('PPT text targets use paragraphs only at complete UTF-16 paragraph boundaries', async () => {
  const r = await convert(deck([shape(2050, { text: 'A😀B\rSecond' })],
    [timeline([setNode(2050, { mode: 2, start: 0, end: 5 }),
      setNode(2050, { mode: 2, start: 1, end: 3 })])]));
  assert.match(r.xml, /<p:pRg st="0" end="0"\/>/);
  assert.match(r.xml, /<p:charRg st="1" end="2"\/>/);
});

test('PPT a damaged behavior drops its dependent effect without losing sibling effects', async () => {
  const broken = timeNode([container(0xf131, [rec(0xf13a, u32(1))])], { kind: 3 });
  const fixture = deck([shape(2050)], [timeline([broken,
    setNode(2050, { extra: [condition(1, 4, 0, 2, 3)] }), effect(2050)])]);
  const r = await convert(fixture);
  assert.ok(animationLoss(r));
  const ids = new Set(nodes(r.xml, 'p:cTn').map(n => n.attrs.id));
  for (const n of nodes(r.xml, 'p:tn')) assert.ok(ids.has(n.attrs.val), 'No reference to a dropped time node');
  assert.equal(nodes(r.xml, 'p:set').length, 1);
  assert.match(r.xml, /<p:sp>/);
});

test('PPT a damaged condition must not turn a timed effect into an immediate effect', async () => {
  const r = await convert(deck([shape(2050)], [timeline([
    setNode(2050, { extra: [container(0xf125, [rec(0xf128, u32(2))], 1)] }), effect(2050)])]));
  assert.ok(animationLoss(r));
  assert.equal(nodes(r.xml, 'p:set').length, 1);
});

test('PPT malformed timing properties drop one effect without emitting invalid numeric XML', async () => {
  const r = await convert(deck([shape(2050)], [timeline([
    timeNode([properties([[9, 'not a number']]), setNode(2050)]), effect(2050)])]));
  assert.ok(animationLoss(r));
  assert.doesNotMatch(r.xml, /presetID="not a number"/);
  assert.equal(nodes(r.xml, 'p:set').length, 1);
});

test('PPT linked video relationships retain the source path and do not fetch media', async () => {
  const r = await convert(deck([shape(2050, { client: [rec(3009, u32(7))] })], [],
    { documentRecords: [container(1033, [linkedVideo(7, 'C:\\Media\\A # clip.avi')])] }));
  assert.match(r.xml, /<a:videoFile r:link="rId\d+"\/>/);
  const rels = await r.zip.get('ppt/slides/_rels/slide1.xml.rels').text();
  assert.match(rels, /relationships\/video" Target="file:\/\/\/C:\/Media\/A%20%23%20clip.avi" TargetMode="External"/);
  assert.ok(![...r.zip.keys()].some(n => n.endsWith('.avi')));
});

test('PPT legacy media keeps looping, rewind, hide, slide count and the click-to-pause sequence', async () => {
  const r = await convert(deck([shape(2050, { client: [rec(3009, u32(7)),
    oldAnimation({ build: 0, flags: 256 | 4096, slideCount: 1 })] })], [],
  { documentRecords: [container(1033, [linkedVideo(7, 'C:\\clip.avi', 3)])] }));
  assert.match(r.xml, /<p:cMediaNode vol="80000" numSld="1" showWhenStopped="0">/);
  assert.match(r.xml, /repeatCount="indefinite"/);
  assert.match(r.xml, /cmd="playFrom\(0\.0\)"/);
  assert.match(r.xml, /cmd="togglePause"/);
  assert.match(r.xml, /nodeType="interactiveSeq"/);
  assert.ok(!animationLoss(r));
});

test('PPT a damaged media link reports its loss and preserves a valid sibling link', async () => {
  const r = await convert(deck([shape(2050, { client: [rec(3009, u32(7))] }),
    shape(2051, { client: [rec(3009, u32(8))] })], [], { documentRecords: [container(1033, [
    linkedVideo(7, 'javascript:alert(1)'), linkedVideo(8, 'https://example.invalid/clip.avi')])] }));
  const rels = await r.zip.get('ppt/slides/_rels/slide1.xml.rels').text();
  assert.doesNotMatch(rels, /javascript:/);
  assert.match(rels, /https:\/\/example.invalid\/clip.avi/);
  assert.equal(nodes(r.xml, 'a:videoFile').length, 1);
  assert.ok(r.losses.some(l => l.id === 'legacy-presentation-media' && l.notify));
});

test('PPT a legacy blocking media action without a stored duration reports the lost wait', async () => {
  const r = await convert(deck([shape(2050, { client: [rec(3009, u32(7)),
    oldAnimation({ build: 0, flags: 256 | 1024 })] })], [],
  { documentRecords: [container(1033, [linkedVideo(7, 'C:\\clip.avi')])] }));
  assert.ok(animationLoss(r)?.detail.some(d => d.reason === 'PPT_TIMING_MEDIA_DURATION'));
  assert.match(r.xml, /cmd="playFrom\(0\.0\)"/);
});

test('PPT embedded sound bytes and media targets survive together', async () => {
  const audio = Buffer.from('RIFF\x04\0\0\0WAVE');
  const object = container(4111, [rec(4100, bytes(u32(7), u32(0))), rec(4115, bytes(u32(12), u32(906)))]);
  const media = timeNode([target(2050, 3)], { kind: 4 });
  const r = await convert(deck([shape(2050, { client: [rec(3009, u32(7))] })], [timeline([media])],
    { documentRecords: [container(2020, [sound(12, 'beep.wav', audio)]), container(1033, [object])] }));
  const part = [...r.zip.keys()].find(n => /^ppt\/media\/audio\d+\.wav$/.test(n));
  assert.ok(part);
  assert.deepEqual(Buffer.from(await r.zip.get(part).bytes()), audio);
  assert.match(r.xml, /<a:audioFile r:link="rId\d+"/);
  assert.match(r.xml, /<p:audio>/);
  assert.ok(!animationLoss(r));
});

test('PPT missing animation sounds drop only that sound and report the content loss', async () => {
  const r = await convert(deck([shape(2050)], [timeline([
    timeNode([target(99, 3, -1, -1, 2)], { kind: 4 }), effect(2050)])]));
  assert.ok(animationLoss(r));
  assert.equal(nodes(r.xml, 'p:audio').length, 0);
  assert.equal(nodes(r.xml, 'p:set').length, 1);
});

test('PPT a null transition sound reference has no missing-sound notice', async () => {
  const transition = rec(1017, bytes(u32(0), u32(0), Buffer.from([0, 5]), u16(17), Buffer.alloc(4)));
  const r = await convert(deck([], [transition]));
  assert.match(r.xml, /<p:dissolve\/>/);
  assert.doesNotMatch(r.xml, /<p:sndAc>/);
  assert.ok(!r.losses.some(l => /sound|media/.test(l.id)));
});

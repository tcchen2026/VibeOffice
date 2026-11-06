// Independent binary records for animation and media regressions (MS-PPT).
import { rec, container, u16, u32, presentation, slideAtom, scheme } from './ppt-fixtures.mjs';
export { rec, container, u16, u32 };
export const bytes = (...items) => Buffer.concat(items);
export const string = text => Buffer.from(text + '\0', 'utf16le');
export const variant = (id, value) => rec(0xf142, typeof value === 'string'
  ? bytes(Buffer.from([3]), string(value))
  : bytes(Buffer.from([1]), u32(value)), id);
export const properties = entries => container(0xf13d, entries.map(([id, value]) => variant(id, value)));
export function target(id, mode = 0, start = -1, end = -1, type = 1) {
  return container(0xf13c, [rec(0x2afb, bytes(u32(mode), u32(type), u32(id), u32(start), u32(end)))]);
}
export function condition(list = 1, event = 0, delay = 0, object = 0, id = 0, visual = []) {
  return container(0xf125, [rec(0xf128, bytes(u32(object), u32(event), u32(id), u32(delay))), ...visual], list);
}
export function timeNode(children, options = {}) {
  const data = Buffer.alloc(32);
  data.writeUInt32LE(options.kind || 0, 8);
  data.writeUInt32LE(options.fill || 0, 12);
  data.writeInt32LE(options.duration ?? -1, 24);
  data.writeUInt32LE(8 | 16 | (options.fill ? 1 : 0), 28);
  return container(0xf144, [rec(0xf127, data), ...children]);
}
export function setBehavior(id, options = {}) {
  const common = container(0xf12a, [rec(0xf133, bytes(u32(4), Buffer.alloc(12))),
    container(0xf13e, [variant(0, 'style.visibility')]),
    target(id, options.mode || 0, options.start ?? -1, options.end ?? -1)]);
  return container(0xf131, [rec(0xf13a, bytes(u32(1), u32(0))), variant(1, 'visible'), common]);
}
export const setNode = (id, options = {}) => timeNode([
  ...options.extra || [], setBehavior(id, options)], { kind: 3, duration: options.duration ?? 1, fill: 3 });
export function keyframeNode(id, time) {
  const common = container(0xf12a, [rec(0xf133, bytes(u32(4), Buffer.alloc(12))),
    container(0xf13e, [variant(0, 'ppt_w')]), target(id)]);
  const keys = container(0xf13f, [rec(0xf143, u32(time)), variant(0, '#ppt_w')]);
  return timeNode([container(0xf12b, [rec(0xf134, bytes(u32(1), u32(56), u32(1))), common, keys])],
    { kind: 3, duration: 500, fill: 3 });
}
export function timeline(children, options = {}) {
  const sequence = timeNode([properties([[20, 4]]), ...children], { kind: 1 });
  return container(5000, [container(5002, [rec(4026, string('___PPT10')),
    rec(5003, bytes(timeNode([properties([[20, 9]]), sequence]), ...options.builds || []))])]);
}
export function oldAnimation(options = {}) {
  const data = Buffer.alloc(28);
  data.writeUInt32LE(options.flags || 0, 4);
  data.writeUInt32LE(options.sound || 0, 8);
  data.writeInt32LE(options.delay || 0, 12);
  data.writeInt16LE(options.order || 0, 16);
  data.writeUInt16LE(options.slideCount || 0, 18);
  data[20] = options.build ?? 1;
  data[21] = options.effect || 0;
  data[22] = options.direction || 0;
  data[23] = options.after || 0;
  data[24] = options.textMode || 0;
  return container(4116, [rec(4081, data, 0, 1)]);
}
export function shape(id, options = {}) {
  const text = options.text == null ? [] : [rec(0xf00d, bytes(rec(3999, u32(4)),
    rec(4000, Buffer.from(options.text, 'utf16le'))))];
  return container(0xf004, [rec(0xf00a, bytes(u32(id), u32(0xa00)), 1, 2),
    rec(0xf010, bytes(...[100, 200, 1200, 600].map(u32))), ...text,
    ...(options.client ? [rec(0xf011, bytes(...options.client))] : [])]);
}
export function deck(shapes, records = [], options = {}) {
  return presentation({ documentRecords: options.documentRecords || [], pages: {
    masters: [{ id: 0x80000000, ref: 2, record: container(1016, [slideAtom(),
      scheme(['FFFFFF', '000000', '808080', '123456', 'BBE0E3', '333399', '009999', '99CC00']),
      ...options.masterRecords || []]) }],
    slides: [{ id: 256, ref: 3, record: container(1006, [slideAtom(0x80000000, 0, 7, 16),
      rec(1036, container(0xf002, shapes)), ...records]) }], notes: []
  } });
}
export const linkedVideo = (id, url, flags = 0) => container(4102, [container(4101, [
  rec(4100, bytes(u32(id), u32(flags))), rec(4026, string(url))])]);
export const sound = (id, name, data) => container(2022, [rec(4026, string(name), 0),
  rec(4026, string('.wav'), 1), rec(4026, string(String(id)), 2), rec(2023, data)]);

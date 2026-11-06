/* MS-CFB 2.2–2.7: package unchanged VBA streams in a version-3 compound file.
 * This is storage assembly for vbaProject.bin, not a binary Office document writer.
 * Reading always uses the shared compound-file reader in common/crypto.js. */
(function (root) {
  'use strict';
  const B = root.L.legacy;
  const V = B.vba = {};
  const SECTOR = 512;
  const MINI_SECTOR = 64;
  const MINI_CUTOFF = 4096;
  const DIRECTORY_ENTRY = 128;
  const FAT_ENTRIES = SECTOR / 4;
  const HEADER_DIFAT = 109;
  const FREE = 0xffffffff;
  const END = 0xfffffffe;
  const FAT = 0xfffffffd;
  const DIFAT = 0xfffffffc;
  const STORAGE = 1;
  const STREAM = 2;
  const ROOT = 5;
  const bad = detail => B.fail('VBA_STORAGE', 'Invalid VBA project storage', detail);
  const upper = name => name.split('').map(c => {
    const value = c.toUpperCase();
    return value.length === 1 ? value : c;
  }).join('');
  const compare = (a, b) => a.name.length - b.name.length ||
    (a.key < b.key ? -1 : a.key > b.key ? 1 : 0);

  V.pack = function (source, prefix = '') {
    const storageName = prefix.replace(/\/$/, '');
    const storage = storageName
      ? source.entries.find(e => upper(e.path) === upper(storageName) && e.type === STORAGE)
      : source.root;
    if (!storage) bad('Missing project root');
    const sourcePrefix = storageName ? storage.path + '/' : '';
    const entries = [{ ...storage, name: 'Root Entry', path: '', type: ROOT, size: 0 }];
    const selected = source.entries.filter(e => upper(e.path).startsWith(upper(sourcePrefix)));
    if (selected.length > 65535) bad('Too many project entries');
    const byPath = new Map([['', entries[0]]]);
    let size = 0;
    let miniCount = 0;
    let dataSectors = 0;
    for (const original of selected) {
      const path = original.path.slice(sourcePrefix.length);
      if (!path || ![STORAGE, STREAM].includes(original.type)) bad('Invalid project entry');
      const name = original.name;
      if (!name || name.length > 31 || /[\0/\\:!]/.test(name)) bad('Invalid project entry name');
      const data = original.type === STREAM ? source.stream(original.path) : null;
      if (data && data.length !== original.size) bad('Truncated project stream: ' + path);
      if (original.type === STREAM && !data) bad('Missing project stream: ' + path);
      size += data?.length || 0;
      if (size > B.limits.bytes) bad('Project exceeds the byte limit');
      const entry = { ...original, path, name, key: upper(name), data,
        size: data?.length || 0, id: entries.length, children: [], left: FREE, right: FREE, child: FREE };
      const key = path.split('/').map(upper).join('/');
      if (byPath.has(key)) bad('Duplicate project entry');
      byPath.set(key, entry);
      entries.push(entry);
      if (entry.type === STREAM && entry.size) {
        if (entry.size < MINI_CUTOFF) miniCount += Math.ceil(entry.size / MINI_SECTOR);
        else dataSectors += Math.ceil(entry.size / SECTOR);
      }
    }
    Object.assign(entries[0], { id: 0, children: [], left: FREE, right: FREE, child: FREE });
    for (const entry of entries.slice(1)) {
      const slash = entry.path.lastIndexOf('/');
      const parentPath = slash < 0 ? '' : entry.path.slice(0, slash);
      const parent = byPath.get(parentPath.split('/').map(upper).join('/'));
      if (!parent || ![STORAGE, ROOT].includes(parent.type)) bad('Missing project parent');
      parent.children.push(entry);
    }
    // A balanced sibling tree with the incomplete bottom level red. Comparisons
    // use CFB's length-first, single-UTF-16-unit uppercase ordering.
    for (const entry of entries) {
      const children = entry.children.sort(compare);
      let redLevel = 0;
      for (let count = children.length; count; count = Math.floor((count - 1) / 2)) redLevel++;
      const tree = (lo, hi, depth) => {
        if (lo > hi) return FREE;
        const mid = (lo + hi) >>> 1;
        const child = children[mid];
        child.color = depth === redLevel ? 0 : 1;
        child.left = tree(lo, mid - 1, depth + 1);
        child.right = tree(mid + 1, hi, depth + 1);
        return child.id;
      };
      entry.child = tree(0, children.length - 1, 0);
    }
    const directorySectors = Math.ceil(entries.length * DIRECTORY_ENTRY / SECTOR);
    const miniSectors = Math.ceil(miniCount * MINI_SECTOR / SECTOR);
    const miniFatSectors = Math.ceil(miniCount / FAT_ENTRIES);
    const payloadSectors = directorySectors + miniSectors + miniFatSectors + dataSectors;
    let fatSectors = 0;
    let difatSectors = 0;
    for (;;) {
      const nextFat = Math.ceil((payloadSectors + fatSectors + difatSectors) / FAT_ENTRIES);
      const nextDifat = Math.ceil(Math.max(0, nextFat - HEADER_DIFAT) / (FAT_ENTRIES - 1));
      if (nextFat === fatSectors && nextDifat === difatSectors) break;
      fatSectors = nextFat;
      difatSectors = nextDifat;
    }
    const sectors = payloadSectors + fatSectors + difatSectors;
    if ((sectors + 1) * SECTOR > B.limits.bytes) bad('Packaged project exceeds the byte limit');
    const out = new Uint8Array((sectors + 1) * SECTOR);
    const view = new DataView(out.buffer);
    const word = (at, value) => view.setUint16(at, value, true);
    const dword = (at, value) => view.setUint32(at, value, true);
    const offset = sector => (sector + 1) * SECTOR;
    const fatStart = payloadSectors;
    const difatStart = fatStart + fatSectors;
    out.fill(255, offset(fatStart));
    const chain = (first, count, at = offset(fatStart)) => {
      for (let i = 0; i < count; i++) dword(at + (first + i) * 4, i + 1 < count ? first + i + 1 : END);
    };
    out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    word(24, 0x3e);
    word(26, 3);
    word(28, 0xfffe);
    word(30, 9);
    word(32, 6);
    dword(44, fatSectors);
    dword(48, 0);
    dword(56, MINI_CUTOFF);
    dword(60, miniFatSectors ? directorySectors + miniSectors : END);
    dword(64, miniFatSectors);
    dword(68, difatSectors ? difatStart : END);
    dword(72, difatSectors);
    for (let i = 0; i < HEADER_DIFAT; i++) dword(76 + i * 4, i < fatSectors ? fatStart + i : FREE);
    for (let i = 0; i < difatSectors; i++) {
      for (let j = 0; j < FAT_ENTRIES - 1; j++) {
        const index = HEADER_DIFAT + i * (FAT_ENTRIES - 1) + j;
        dword(offset(difatStart + i) + j * 4, index < fatSectors ? fatStart + index : FREE);
      }
      dword(offset(difatStart + i) + SECTOR - 4, i + 1 < difatSectors ? difatStart + i + 1 : END);
    }
    chain(0, directorySectors);
    chain(directorySectors, miniSectors);
    chain(directorySectors + miniSectors, miniFatSectors);
    for (let i = 0; i < fatSectors; i++) dword(offset(fatStart) + (fatStart + i) * 4, FAT);
    for (let i = 0; i < difatSectors; i++) dword(offset(fatStart) + (difatStart + i) * 4, DIFAT);
    const miniFatOffset = offset(directorySectors + miniSectors);
    if (miniFatSectors) out.fill(255, miniFatOffset, miniFatOffset + miniFatSectors * SECTOR);
    let miniAt = 0;
    let sectorAt = directorySectors + miniSectors + miniFatSectors;
    entries[0].start = miniSectors ? directorySectors : END;
    entries[0].size = miniCount * MINI_SECTOR;
    for (const entry of entries) {
      if (entry.type === STREAM) {
        const small = entry.size < MINI_CUTOFF;
        const count = Math.ceil(entry.size / (small ? MINI_SECTOR : SECTOR));
        entry.start = !count ? END : small ? miniAt : sectorAt;
        if (count) {
          chain(entry.start, count, small ? miniFatOffset : offset(fatStart));
          out.set(entry.data, small ? offset(directorySectors) + miniAt * MINI_SECTOR : offset(sectorAt));
          if (small) miniAt += count;
          else sectorAt += count;
        }
      }
      const at = SECTOR + entry.id * DIRECTORY_ENTRY;
      for (let i = 0; i < entry.name.length; i++) word(at + i * 2, entry.name.charCodeAt(i));
      word(at + 64, (entry.name.length + 1) * 2);
      out[at + 66] = entry.type;
      out[at + 67] = entry.color ?? 1;
      dword(at + 68, entry.left);
      dword(at + 72, entry.right);
      dword(at + 76, entry.child);
      if (entry.type !== STREAM) {
        if (entry.clsid?.length === 16) out.set(entry.clsid, at + 80);
        if (entry.type !== ROOT && entry.created?.length === 8) out.set(entry.created, at + 100);
        if (entry.modified?.length === 8) out.set(entry.modified, at + 108);
      }
      dword(at + 116, entry.type === STORAGE ? 0 : entry.start);
      dword(at + 120, entry.size);
    }
    return out;
  };
})(typeof window !== 'undefined' ? window : globalThis);

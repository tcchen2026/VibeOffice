/* Binary PowerPoint media objects, action sounds and transition sounds.
 * Derived from Apache POI hslf/record/{Sound,ExMediaAtom,ExVideoContainer}.java
 * and hslf/model/MovieShape.java (Apache-2.0).
 * Derived from Apache OpenOffice sd/source/filter/ppt/pptin.cxx (Apache-2.0).
 * MS-PPT 2.4.16, 2.10: preserve bytes and links without fetching external files. */
(function (root) {
  'use strict';
  const B = root.L.legacy, P = B.ppt, D = P.drawings, el = B.el, N = s => s;
  const M = P.media = {};
  const R = { objects: 1033, soundCollection: 2020, sound: 2022, soundData: 2023,
    reference: 3009, string: 4026, quickTime: 4074, media: 4100, video: 4101,
    avi: 4102, movie: 4103, midi: 4109, cd: 4110, embeddedWav: 4111,
    linkedWav: 4112, cdAtom: 4114, wavAtom: 4115, client: 0xf011 };
  const text = r => r.r.text(0, r.r.length & ~1, 1200).split('\0')[0];
  const types = { wav: 'audio/wav', mp3: 'audio/mpeg', mid: 'audio/midi', midi: 'audio/midi',
    wma: 'audio/x-ms-wma', aif: 'audio/aiff', aiff: 'audio/aiff', au: 'audio/basic' };
  const report = (book, page, shape, reason) => {
    page.mediaLoss ||= B.loss('legacy-presentation-media', N('Some sounds or videos could not be kept'),
      page.where, 'drop', page.where, true);
    if (!page.mediaLoss.detail) {
      page.mediaLoss.detail = [];
      book.losses.push(page.mediaLoss);
    }
    if (page.mediaLoss.detail.length < 12)
      page.mediaLoss.detail.push({ shape: shape?.id, reason: String(reason?.code || reason) });
  };
  M.path = function (value) {
    if (!value || /[\x00-\x1f]/.test(value) || /^(?:javascript|vbscript|data):/i.test(value.trim()))
      B.fail('PPT_MEDIA_LINK', 'Invalid media link');
    const filePath = value.replace(/\\/g, '/');
    const fileUri = prefix => encodeURI(prefix + filePath).replace(/[?#]/g, encodeURIComponent);
    if (/^[a-z]:[\\/]/i.test(value)) return fileUri('file:///');
    if (/^\\\\/.test(value)) return fileUri('file:');
    return value.replace(/\\/g, '/');
  };
  M.prepare = function (book, pkg) {
    const sounds = new Map(), objects = new Map(), relationships = new Map();
    let next = 0;
    const local = fn => {
      try { return fn(); }
      catch (error) {
        if (error.name !== 'LegacyError') throw error;
        return null;
      }
    };
    for (const collection of book.documentRecords.filter(r => r.type === R.soundCollection))
      for (const record of P.children(collection, () => {}).filter(r => r.type === R.sound))
        local(() => {
          const children = P.children(record), strings = new Map();
          for (const child of children) if (child.type === R.string) strings.set(child.instance, text(child));
          const id = Number(strings.get(2)), data = children.find(c => c.type === R.soundData);
          if (!Number.isInteger(id) || !data) return;
          const name = strings.get(0) || '', ext = (strings.get(1) || name.split('.').pop() || '')
            .replace(/^\./, '').toLowerCase();
          if (!types[ext]) return;
          if (!sounds.has(id)) sounds.set(id, { name, ext, data: data.r.bytes });
        });
    for (const list of book.documentRecords.filter(r => r.type === R.objects))
      for (const record of P.children(list, () => {}))
        local(() => {
          if (![R.avi, R.movie, R.midi, R.cd, R.embeddedWav, R.linkedWav, R.quickTime].includes(record.type))
            return;
          let children = P.children(record);
          if (record.type === R.avi || record.type === R.movie) {
            const video = children.find(c => c.type === R.video);
            if (!video) return;
            children = P.children(video);
          }
          const media = children.find(c => c.type === R.media)?.r;
          if (!media) return;
          const id = media.u32(0), flags = media.u32(4);
          const item = { id, flags, kind: [R.avi, R.movie, R.quickTime].includes(record.type) ? 'video' : 'audio' };
          if (record.type === R.embeddedWav) {
            const wav = children.find(c => c.type === R.wavAtom)?.r;
            if (!wav) return;
            item.sound = wav.u32(0);
            item.duration = wav.i32(4);
          } else if (record.type === R.cd) {
            const cd = children.find(c => c.type === R.cdAtom)?.r;
            if (!cd) return;
            cd.check(0, 8);
            const position = at => Math.round((cd.u8(at + 1) * 60 + cd.u8(at + 2) + cd.u8(at + 3) / 75) * 1000);
            item.cd = { startTrack: cd.u8(0), startTime: position(0), endTrack: cd.u8(4), endTime: position(4) };
          } else {
            const link = children.find(c => c.type === R.string);
            if (!link) return;
            try { item.path = M.path(text(link)); }
            catch (error) {
              if (error.name !== 'LegacyError') throw error;
              item.problem = error.code;
            }
          }
          if (!objects.has(id)) objects.set(id, item);
        });
    const addSound = id => {
      const sound = sounds.get(id);
      if (!sound) return null;
      if (!sound.part) {
        do { sound.part = 'ppt/media/audio' + ++next + '.' + sound.ext; } while (pkg.parts.has(sound.part));
        pkg.add(sound.part, sound.data, types[sound.ext]);
      }
      return sound;
    };
    const relation = (type, target, external = false) => {
      const key = book.textPart + ':' + type + ':' + target + ':' + external;
      if (!relationships.has(key)) relationships.set(key,
        pkg.rel(book.textPart, type, external ? target : D.relative(book.textPart, target), { external }));
      return relationships.get(key);
    };
    book.presentationMedia = { sounds, objects, addSound, relation };
    const shapes = (list, depth = 0) => {
      if (depth > B.limits.depth) B.fail('PPT_MEDIA_LIMIT', 'Too many nested media shapes');
      return list.flatMap(shape => [shape, ...shapes(shape.children || [], depth + 1)]);
    };
    for (const page of D.pages(book))
      for (const shape of shapes(page.textShapes || [])) {
        const client = shape.records?.find(r => r.type === R.client);
        if (!client) continue;
        local(() => {
          const reference = P.children(client).find(r => r.type === R.reference);
          if (!reference) return;
          const item = objects.get(reference.r.u32(0));
          if (!item) return;
          if (item.problem) {
            report(book, page, shape, item.problem);
            return;
          }
          if (item.sound != null && !sounds.has(item.sound)) {
            report(book, page, shape, 'Missing embedded sound');
            return;
          }
          shape.media = item;
        });
      }
  };
  M.sound = function (book, id, tag = 'a:snd') {
    const sound = book.presentationMedia?.addSound(id);
    if (!sound) return '';
    const rid = book.presentationMedia.relation('audio', sound.part);
    return el(tag, { 'r:embed': rid, name: sound.name || null });
  };
  M.target = (book, id) => {
    const sound = M.sound(book, id, 'p:sndTgt');
    return sound ? el('p:tgtEl', {}, sound) : '';
  };
  M.shape = function (book, page, shape) {
    const item = shape.media;
    if (!item) return '';
    try {
      if (item.cd) return el('a:audioCd', {},
        el('a:st', { track: item.cd.startTrack, time: item.cd.startTime }) +
        el('a:end', { track: item.cd.endTrack, time: item.cd.endTime }));
      const media = book.presentationMedia, sound = item.sound == null ? null : media.addSound(item.sound);
      const target = sound?.part || item.path;
      if (!target) return '';
      const rid = media.relation(item.kind, target, !!item.path);
      return el(item.kind === 'video' ? 'a:videoFile' : 'a:audioFile', { 'r:link': rid });
    } catch (error) {
      if (error.name !== 'LegacyError') throw error;
      report(book, page, shape, error);
      return '';
    }
  };
  M.transition = function (book, page, record) {
    const flags = record.r.u16(10);
    if (flags & 256) return el('p:sndAc', {}, el('p:endSnd'));
    const soundId = record.r.u32(4);
    if (!(flags & 16) || !soundId) return '';
    const sound = M.sound(book, soundId, 'p:snd');
    if (!sound) {
      report(book, page, null, 'Missing transition sound');
      return '';
    }
    return el('p:sndAc', {}, el('p:stSnd', { loop: flags & 64 ? 1 : null }, sound));
  };
  M.report = report;
})(typeof window !== 'undefined' ? window : globalThis);

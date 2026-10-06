/* Lectern — built-in clip art (original vector drawings, 64×64). */
(function () {
  'use strict';
  const L = window.L;
  const O = 'stroke="#222" stroke-width="2" stroke-linejoin="round"';
  const items = [
    ['Light bulb', 'idea light bulb think innovation creative', `<path d="M32 6c-11 0-19 8-19 18 0 7 4 11 7 15 2 3 3 5 3 8h18c0-3 1-5 3-8 3-4 7-8 7-15 0-10-8-18-19-18z" fill="#ffe259" ${O}/><path d="M24 22a9 9 0 0 1 8-8" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round"/><rect x="23" y="47" width="18" height="5" fill="#b9c2cc" ${O}/><rect x="24" y="52" width="16" height="4" fill="#9aa3ad" ${O}/><path d="M28 56h8l-2 4h-4z" fill="#555" ${O}/>`],
    ['Globe', 'globe world international earth global', `<circle cx="32" cy="32" r="25" fill="#5aa9e6" ${O}/><path d="M18 18c6 2 6 8 12 8s4 8 10 6 6 6 4 10-8 2-10 8-8 4-10-2 2-8-4-10-8-2-6-8 0-10 4-12z" fill="#6cc04a" ${O}/><path d="M7 32h50M32 7c-9 8-9 42 0 50M32 7c9 8 9 42 0 50" fill="none" stroke="#1e4f80" stroke-width="1.3" opacity=".6"/>`],
    ['Rocket', 'rocket launch start growth space', `<path d="M32 4c9 7 13 18 12 32l-6 8H26l-6-8C19 22 23 11 32 4z" fill="#eef2f7" ${O}/><circle cx="32" cy="24" r="6" fill="#4aa3df" ${O}/><path d="M20 36l-8 12 10-3zM44 36l8 12-10-3z" fill="#e74c3c" ${O}/><path d="M26 44h12l-2 6h-8z" fill="#888" ${O}/><path d="M28 50c0 6 4 10 4 10s4-4 4-10z" fill="#ff9f1a" ${O}/>`],
    ['Trophy', 'trophy award win success prize achievement', `<path d="M18 8h28v10c0 10-6 17-14 17S18 28 18 18z" fill="#f5c542" ${O}/><path d="M18 12H9c0 9 5 14 11 14M46 12h9c0 9-5 14-11 14" fill="none" ${O}/><rect x="28" y="35" width="8" height="10" fill="#d4a017" ${O}/><path d="M20 45h24v6H20z" fill="#8b5a2b" ${O}/><path d="M16 51h32v7H16z" fill="#6d4421" ${O}/><path d="M25 13v10" stroke="#fff" stroke-width="3" stroke-linecap="round"/>`],
    ['Gears', 'gears process settings engineering machine work', `<g ${O} fill="#9aa7b8"><path d="M26 10l4 0 1 5 4 2 4-3 3 3-3 4 2 4 5 1v4l-5 1-2 4 3 4-3 3-4-3-4 2-1 5h-4l-1-5-4-2-4 3-3-3 3-4-2-4-5-1v-4l5-1 2-4-3-4 3-3 4 3 4-2z"/></g><circle cx="28" cy="27" r="6" fill="#fff" ${O}/><g ${O} fill="#e8a33d"><path d="M46 36h3l1 3 3 1 2-2 2 2-2 2 1 3 3 1v3l-3 1-1 3 2 2-2 2-2-2-3 1-1 3h-3l-1-3-3-1-2 2-2-2 2-2-1-3-3-1v-3l3-1 1-3-2-2 2-2 2 2 3-1z"/></g><circle cx="47.5" cy="47.5" r="3.5" fill="#fff" ${O}/>`],
    ['Growth chart', 'chart graph growth results increase sales profit', `<rect x="6" y="8" width="52" height="48" fill="#fff" ${O}/><rect x="12" y="38" width="8" height="14" fill="#5b9bd5" ${O}/><rect x="24" y="30" width="8" height="22" fill="#5b9bd5" ${O}/><rect x="36" y="22" width="8" height="30" fill="#5b9bd5" ${O}/><rect x="48" y="14" width="6" height="38" fill="#5b9bd5" ${O}/><path d="M10 34l12-8 10 2 18-16" fill="none" stroke="#e74c3c" stroke-width="3" stroke-linecap="round"/><path d="M45 10l7 1-1 7" fill="none" stroke="#e74c3c" stroke-width="3" stroke-linecap="round"/>`],
    ['Calendar', 'calendar schedule date plan time deadline', `<rect x="8" y="12" width="48" height="44" rx="3" fill="#fff" ${O}/><path d="M8 15a3 3 0 0 1 3-3h42a3 3 0 0 1 3 3v9H8z" fill="#d64541" ${O}/><path d="M20 6v10M44 6v10" stroke="#222" stroke-width="3" stroke-linecap="round"/><g fill="#9aa7b8"><rect x="14" y="30" width="7" height="6"/><rect x="25" y="30" width="7" height="6"/><rect x="36" y="30" width="7" height="6"/><rect x="14" y="41" width="7" height="6"/><rect x="25" y="41" width="7" height="6"/></g><rect x="36" y="41" width="7" height="6" fill="#d64541"/>`],
    ['Magnifier', 'search magnifier research find investigate', `<circle cx="26" cy="26" r="17" fill="#cfe9ff" ${O}/><path d="M18 20a10 10 0 0 1 8-6" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round"/><path d="M38 38l16 16" stroke="#6d4421" stroke-width="8" stroke-linecap="round"/><path d="M38 38l16 16" stroke="#222" stroke-width="10" stroke-linecap="round" opacity=".15"/>`],
    ['Team', 'people team group staff meeting users', `<circle cx="20" cy="20" r="8" fill="#f2c39b" ${O}/><path d="M6 50c0-12 6-18 14-18s14 6 14 18z" fill="#4a7bc8" ${O}/><circle cx="44" cy="20" r="8" fill="#e0a87c" ${O}/><path d="M30 50c0-12 6-18 14-18s14 6 14 18z" fill="#6cb04a" ${O}/><circle cx="32" cy="28" r="9" fill="#f2c39b" ${O}/><path d="M16 60c0-14 7-21 16-21s16 7 16 21z" fill="#e8641c" ${O}/>`],
    ['House', 'house home real estate property building', `<path d="M8 30L32 8l24 22" fill="none" stroke="#222" stroke-width="3" stroke-linejoin="round"/><path d="M14 26v30h36V26L32 10z" fill="#f4e3c1" ${O}/><path d="M6 31L32 7l26 24-4 3L32 14 10 34z" fill="#c0392b" ${O}/><rect x="27" y="38" width="10" height="18" fill="#8b5a2b" ${O}/><rect x="17" y="32" width="8" height="8" fill="#9fd3ff" ${O}/><rect x="40" y="32" width="8" height="8" fill="#9fd3ff" ${O}/>`],
    ['Computer', 'computer monitor technology pc screen it', `<rect x="6" y="8" width="52" height="36" rx="2" fill="#d5dbe3" ${O}/><rect x="10" y="12" width="44" height="28" fill="#2f6fd0" ${O}/><path d="M14 16h20M14 21h14" stroke="#cfe3ff" stroke-width="2"/><path d="M26 44h12l2 8H24z" fill="#b6bec9" ${O}/><rect x="16" y="52" width="32" height="5" rx="1" fill="#9aa3ad" ${O}/>`],
    ['Telephone', 'telephone phone call contact support', `<path d="M14 8c-4 2-7 6-6 11 3 15 17 31 33 35 5 1 9-2 11-6l-9-8-6 4c-7-3-12-9-15-15l4-6z" fill="#e74c3c" ${O}/><path d="M15 12c-2 1-3 3-3 5" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round"/>`],
    ['Envelope', 'envelope mail email message letter', `<rect x="6" y="14" width="52" height="36" rx="2" fill="#fff8e1" ${O}/><path d="M6 16l26 20 26-20" fill="none" ${O}/><path d="M6 50l20-17M58 50L38 33" fill="none" stroke="#222" stroke-width="1.5"/>`],
    ['Star', 'star rating favorite quality excellent', `<path d="M32 5l8 17 18 2-13 13 4 19-17-9-17 9 4-19L6 24l18-2z" fill="#ffd34e" ${O}/><path d="M32 13l4 10" stroke="#fff" stroke-width="2.5" stroke-linecap="round"/>`],
    ['Check mark', 'check done complete approved yes ok', `<circle cx="32" cy="32" r="26" fill="#3fae49" ${O}/><path d="M19 33l9 9 18-20" fill="none" stroke="#fff" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>`],
    ['Clock', 'clock time hour deadline schedule', `<circle cx="32" cy="32" r="26" fill="#fff" ${O}/><circle cx="32" cy="32" r="22" fill="none" stroke="#4a7bc8" stroke-width="3"/><path d="M32 16v16l10 7" fill="none" stroke="#222" stroke-width="3" stroke-linecap="round"/><g fill="#222"><circle cx="32" cy="12" r="1.5"/><circle cx="52" cy="32" r="1.5"/><circle cx="32" cy="52" r="1.5"/><circle cx="12" cy="32" r="1.5"/></g>`],
    ['Puzzle', 'puzzle solution fit strategy problem piece', `<path d="M10 18h12c-2-8 12-8 10 0h12v12c8-2 8 12 0 10v12H32c2 8-12 8-10 0H10V40c-8 2-8-12 0-10z" fill="#9b59b6" ${O}/><path d="M14 22h6" stroke="#fff" stroke-width="2.5" stroke-linecap="round"/>`],
    ['Target', 'target goal objective aim focus', `<circle cx="30" cy="34" r="24" fill="#fff" ${O}/><circle cx="30" cy="34" r="17" fill="#e74c3c" ${O}/><circle cx="30" cy="34" r="10" fill="#fff" ${O}/><circle cx="30" cy="34" r="4" fill="#e74c3c" ${O}/><path d="M30 34L56 8" stroke="#6d4421" stroke-width="3"/><path d="M50 6l8-2-2 8-5 1z" fill="#3fae49" ${O}/>`],
    ['Money', 'money finance budget dollar coins cost revenue', `<path d="M20 22c-6 8-10 16-10 22 0 10 10 14 22 14s22-4 22-14c0-6-4-14-10-22z" fill="#6cb04a" ${O}/><path d="M22 14h20l-4 8H26z" fill="#5a9a3d" ${O}/><path d="M24 14c-2-6 4-8 8-4 4-4 10-2 8 4" fill="none" ${O}/><text x="32" y="48" text-anchor="middle" font-family="Arial" font-weight="bold" font-size="18" fill="#fff">$</text>`],
    ['Flag', 'flag milestone goal finish achievement', `<path d="M14 6v54" stroke="#6d4421" stroke-width="4" stroke-linecap="round"/><path d="M16 8c10-4 18 4 28 0s12 0 12 0v24s-4-4-12 0-18-4-28 0z" fill="#2e86c1" ${O}/><path d="M16 8h10v6H16zM36 14h10v6H36zM26 20h10v6H26z" fill="#fff" opacity=".85"/>`],
    ['Book', 'book education learning training study read', `<path d="M8 12c8-2 16-2 24 4v40c-8-6-16-6-24-4z" fill="#fff" ${O}/><path d="M56 12c-8-2-16-2-24 4v40c8-6 16-6 24-4z" fill="#fff" ${O}/><path d="M12 22c5-1 10 0 15 2M12 29c5-1 10 0 15 2M37 24c5-2 10-3 15-2M37 31c5-2 10-3 15-2" stroke="#9aa7b8" stroke-width="1.6" fill="none"/><path d="M6 14v42c9-2 18-2 26 3 8-5 17-5 26-3V14" fill="none" stroke="#c0392b" stroke-width="3"/>`],
    ['Leaf', 'leaf environment green nature sustainability eco', `<path d="M10 54C8 30 22 10 54 8c2 28-12 46-40 46z" fill="#6cc04a" ${O}/><path d="M12 52C24 40 34 30 46 18" fill="none" stroke="#2f7a1f" stroke-width="2.5" stroke-linecap="round"/><path d="M24 40l-2-10M32 32l-1-9M38 26l6 1M30 36l8 2" stroke="#2f7a1f" stroke-width="1.6"/>`],
    ['Cloud', 'cloud weather computing storage online', `<path d="M18 48c-8 0-12-5-12-11s5-11 11-11c2-9 9-14 17-14 9 0 16 7 16 15 6 0 10 5 10 10 0 7-5 11-11 11z" fill="#eaf4ff" ${O}/><path d="M16 32c1-3 3-4 6-5" fill="none" stroke="#fff" stroke-width="2.5" stroke-linecap="round"/>`],
    ['Warning', 'warning caution risk alert attention danger', `<path d="M32 6l28 50H4z" fill="#ffd400" ${O}/><path d="M32 22v18" stroke="#222" stroke-width="6" stroke-linecap="round"/><circle cx="32" cy="48" r="3.5" fill="#222"/>`],
  ];
  L.clipart = {
    items: items.map(([name, keys, body], i) => ({ id: 'clip' + i, name, keys, svg: `<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" viewBox="0 0 64 64">${body}</svg>` })),
    search(q) {
      q = String(q || '').toLowerCase().trim();
      if (!q) return L.clipart.items;
      const words = q.split(/\s+/);
      return L.clipart.items.filter((it) => words.every((w) => (it.name + ' ' + it.keys).toLowerCase().includes(w)));
    },
    url: (it) => 'data:image/svg+xml;utf8,' + encodeURIComponent(it.svg),
    /** rasterize to PNG so .pptx files open everywhere */
    async toPNG(it, size) {
      size = size || 512;
      const img = await L.loadImage(L.clipart.url(it));
      const c = document.createElement('canvas');
      c.width = size; c.height = size;
      c.getContext('2d').drawImage(img, 0, 0, size, size);
      return new Promise((res) => c.toBlob((b) => res(b), 'image/png'));
    },
  };
})();

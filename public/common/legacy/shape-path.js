/* Derived from Apache OpenOffice svx/source/customshapes/EnhancedCustomShape2d.cxx,
 * EnhancedCustomShapeGeometry.cxx and filter/source/msfilter/msdffimp.cxx (Apache-2.0).
 * Native OfficeArt guide evaluation and path commands, with cycle/work bounds.
 * Elliptical arcs become short cubic segments (at most22.5 degrees each). */
(function (root) {
  'use strict';
  const N = s => s; // UI refusal text, translated where it is shown.
  const B = root.L.legacy,
    A = B.officeart;
  B.shapePath = function (shape) {
    const props = shape.properties,
      def = B.nativeGeometry?.[shape.type],
      get = (id, fallback = 0) => (props.get(id)?.value ?? fallback) | 0;
    const vertexProp = props.get(325),
      segmentProp = props.get(326),
      guideProp = props.get(342);
    const vertices = vertexProp?.data?.length
      ? A.array(vertexProp).map(v => (v.length === 4 ? [v.i16(0), v.i16(2)] : [v.i32(0), v.i32(4)]))
      : def?.v;
    const segments = segmentProp?.data?.length ? A.array(segmentProp).map(v => v.u16(0)) : def?.s;
    const guides = guideProp?.data?.length
      ? A.array(guideProp).map(g => [g.u16(0), g.u16(2), g.u16(4), g.u16(6)])
      : def?.g || [];
    if (!vertices?.length) B.fail('SHAPE_PATH', N('The custom shape has no vertices'));
    const left = get(320),
      top = get(321),
      width = get(322, def?.w || 21600) - left,
      height = get(323, def?.h || 21600) - top;
    if (width <= 0 || height <= 0) B.fail('SHAPE_PATH', N('Invalid custom-shape coordinate system'));
    const angles = new Set();
    for (const h of def?.handles || []) if (h[0] & 8 && h[2] >= 256 && h[2] < 266) angles.add(h[2] - 256);
    const cache = new Map(),
      active = new Set();
    let work = 0;
    const param = (n, special, depth) => {
      if (!special) return n;
      if (n & 0x400) return guide(n & 1023, depth + 1);
      if (n >= 327 && n <= 336) {
        const i = n - 327;
        return props.has(n) ? get(n) / (angles.has(i) ? 65536 : 1) : def?.a[i] || 0;
      }
      if (n >= 320 && n <= 323) return [left, top, left + width, top + height][n - 320];
      B.fail('SHAPE_GUIDES', N('This shape uses an unsupported geometry parameter'), n);
    };
    const guide = (i, depth = 0) => {
      if (cache.has(i)) return cache.get(i);
      if (++work > B.limits.records || depth > B.limits.depth || active.has(i))
        B.fail('SHAPE_GUIDES', N('Cyclic or excessive custom-shape guide evaluation'));
      const g = guides[i];
      if (!g) B.fail('SHAPE_GUIDES', N('The shape geometry has a missing guide'), i);
      active.add(i);
      const [flags] = g,
        x = param(g[1], flags & 0x2000, depth),
        y = param(g[2], flags & 0x4000, depth),
        z = param(g[3], flags & 0x8000, depth),
        rad = Math.PI / 180;
      let v;
      switch (flags & 255) {
        case 0:
        case 14:
          v = x + y - z;
          break;
        case 1:
          v = (x * y) / (z || 1);
          break;
        case 2:
          v = (x + y) / 2;
          break;
        case 3:
          v = Math.abs(x);
          break;
        case 4:
          v = Math.min(x, y);
          break;
        case 5:
          v = Math.max(x, y);
          break;
        case 6:
          v = x > 0 ? y : z;
          break;
        case 7:
          v = Math.hypot(x, y, z);
          break;
        case 8:
          v = Math.atan2(y, x) / rad;
          break;
        case 9:
          v = x * Math.sin(y * rad);
          break;
        case 10:
          v = x * Math.cos(y * rad);
          break;
        case 11:
          v = x * Math.cos(Math.atan2(z, y));
          break;
        case 12:
          v = x * Math.sin(Math.atan2(z, y));
          break;
        case 13:
          v = Math.sqrt(Math.max(0, x));
          break;
        case 15:
          v = z * Math.sqrt(Math.max(0, 1 - (x / (y || 1)) ** 2));
          break;
        case 16:
          v = x * Math.tan(y * rad);
          break;
        case 128:
          v = Math.sqrt(Math.max(0, z * z - x * x));
          break;
        case 129:
          v = Math.cos(z * rad) * (x - 10800) + Math.sin(z * rad) * (y - 10800) + 10800;
          break;
        case 130:
          v = -Math.sin(z * rad) * (x - 10800) + Math.cos(z * rad) * (y - 10800) + 10800;
          break;
        default:
          B.fail('SHAPE_GUIDES', N('This shape uses an unsupported geometry operation'), flags & 255);
      }
      if (!Number.isFinite(v) || Math.abs(v) > 2 ** 40)
        B.fail('SHAPE_GUIDES', N('Invalid custom-shape guide result'));
      active.delete(i);
      cache.set(i, v);
      return v;
    };
    const coord = n => (n >>> 16 === 0x8000 ? guide(n & 65535) : n | 0);
    let at = 0,
      commands = [],
      paths = [],
      noFill = false,
      noStroke = false,
      current,
      begin;
    const point = () => {
      const v = vertices[at++];
      if (!v) B.fail('SHAPE_PATH', N('Custom-shape path runs beyond its vertices'));
      return [coord(v[0]) - left, coord(v[1]) - top];
    };
    const emit = (op, pts = []) => {
      if (++work > B.limits.records) B.fail('SHAPE_PATH', N('Too many custom-shape commands'));
      commands.push({ op, pts });
      if (pts.length) current = pts.at(-1);
      if (op === 'M') begin = current;
      if (op === 'Z') current = begin;
    };
    const finish = () => {
      if (commands.length) paths.push({ commands, fill: !noFill, stroke: !noStroke });
      commands = [];
      noFill = noStroke = false;
      current = begin = null;
    };
    const ellipse = (cx, cy, rx, ry, start, sweep, connect) => {
      if (!(rx >= 0 && ry >= 0) || ![cx, cy, rx, ry, start, sweep].every(Number.isFinite))
        B.fail('SHAPE_PATH', N('Invalid elliptical arc'));
      sweep = Math.max(-Math.PI * 2, Math.min(Math.PI * 2, sweep));
      const p = [cx + rx * Math.cos(start), cy + ry * Math.sin(start)];
      if (!connect || !current) emit('M', [p]);
      else if (Math.hypot(current[0] - p[0], current[1] - p[1]) > 1e-6) emit('L', [p]);
      const count = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 8))),
        step = sweep / count;
      for (let j = 0; j < count; j++) {
        const a = start + j * step,
          b = a + step,
          k = (4 / 3) * Math.tan(step / 4);
        emit('C', [
          [cx + rx * (Math.cos(a) - k * Math.sin(a)), cy + ry * (Math.sin(a) + k * Math.cos(a))],
          [cx + rx * (Math.cos(b) + k * Math.sin(b)), cy + ry * (Math.sin(b) - k * Math.cos(b))],
          [cx + rx * Math.cos(b), cy + ry * Math.sin(b)]
        ]);
      }
    };
    if (!segments?.length) {
      emit('M', [point()]);
      while (at < vertices.length) emit('L', [point()]);
      if (get(324, 1) === 1) emit('Z');
    }
    for (const value of segments || []) {
      const type = value >>> 13,
        count = value & 8191;
      // AOO's native callout2 and curvedUpArrow definitions include unused trailing
      // commands. CreateSubPath bounds these against nCoordSize and ignores them.
      // Apply that rule only to our trusted native data; file-supplied paths still
      // fail explicitly if their required vertices are missing.
      if (!vertexProp?.data?.length && !segmentProp?.data?.length) {
        const escape = (value >>> 8) & 31,
          n = value & 255;
        const needed =
          type === 0 || type === 2
            ? count || 1
            : type === 1
              ? 3 * (count || 1)
              : (type === 5 || type === 6) && escape <= 9
                ? escape === 0
                  ? n || 1
                  : n
                : 0;
        if (at + needed > vertices.length) continue;
      }
      if (type === 0) for (let i = 0; i < (count || 1); i++) emit('L', [point()]);
      else if (type === 1) for (let i = 0; i < (count || 1); i++) emit('C', [point(), point(), point()]);
      else if (type === 2) for (let i = 0; i < (count || 1); i++) emit('M', [point()]);
      else if (type === 3) emit('Z');
      else if (type === 4) finish();
      else if (type === 5 || type === 6) {
        const escape = (value >>> 8) & 31,
          n = value & 255;
        if (escape === 10) noFill = true;
        else if (escape === 11) noStroke = true;
        else if (escape === 0) for (let i = 0; i < (n || 1); i++) emit('L', [point()]);
        else if (escape >= 3 && escape <= 6) {
          if (n % 4) B.fail('SHAPE_PATH', N('Invalid arc vertex count'));
          for (let i = 0; i < n / 4; i++) {
            const a = point(),
              b = point(),
              s = point(),
              e = point(),
              cx = (a[0] + b[0]) / 2,
              cy = (a[1] + b[1]) / 2,
              rx = Math.abs(b[0] - a[0]) / 2,
              ry = Math.abs(b[1] - a[1]) / 2;
            const start = Math.atan2((s[1] - cy) / (ry || 1), (s[0] - cx) / (rx || 1)),
              end = Math.atan2((e[1] - cy) / (ry || 1), (e[0] - cx) / (rx || 1));
            let sweep = end - start;
            const clockwise = escape >= 5;
            if (clockwise && sweep <= 0) sweep += Math.PI * 2;
            if (!clockwise && sweep >= 0) sweep -= Math.PI * 2;
            ellipse(cx, cy, rx, ry, start, sweep, escape === 3 || escape === 5);
          }
        } else if (escape === 1 || escape === 2) {
          if (n % 3) B.fail('SHAPE_PATH', N('Invalid angle-ellipse vertex count'));
          for (let i = 0; i < n / 3; i++) {
            const c = point(),
              size = point(),
              angle = point();
            ellipse(
              c[0],
              c[1],
              Math.abs(size[0] + left),
              Math.abs(size[1] + top),
              (-(angle[0] + left) * Math.PI) / 180 / 65536,
              (-(angle[1] + top) * Math.PI) / 180 / 65536,
              escape === 1
            );
          }
        } else if (escape === 7 || escape === 8) {
          for (let i = 0; i < n; i++) {
            if (!current) B.fail('SHAPE_PATH', N('Elliptical quadrant has no starting point'));
            const p = point(),
              q = current,
              center = escape === 7 ? [q[0], p[1]] : [p[0], q[1]],
              rx = Math.abs(p[0] - q[0]),
              ry = Math.abs(p[1] - q[1]);
            const start = Math.atan2((q[1] - center[1]) / (ry || 1), (q[0] - center[0]) / (rx || 1)),
              end = Math.atan2((p[1] - center[1]) / (ry || 1), (p[0] - center[0]) / (rx || 1));
            let sweep = end - start;
            if (sweep > Math.PI) sweep -= 2 * Math.PI;
            if (sweep < -Math.PI) sweep += 2 * Math.PI;
            ellipse(center[0], center[1], rx, ry, start, sweep, true);
          }
        } else if (escape === 9) {
          for (let i = 0; i < n / 2; i++) emit('Q', [point(), point()]);
        } else if (![12, 13, 14, 15, 16, 17, 18, 19, 20].includes(escape))
          B.fail('SHAPE_PATH_ESCAPE', N('This shape uses an unsupported path command'), escape);
      } else B.fail('SHAPE_PATH_ESCAPE', N('Unsupported shape path type'));
    }
    finish();
    const rect = def?.t?.[0],
      points = props.get(337)?.data?.length
        ? A.array(props.get(337)).map(v => (v.length === 4 ? [v.i16(0), v.i16(2)] : [v.i32(0), v.i32(4)]))
        : def?.c || [],
      directions = A.array(props.get(338));
    const connections = points.map((v, i) => ({
      x: coord(v[0]) - left,
      y: coord(v[1]) - top,
      angle: directions[i] ? directions[i].i32(0) / 65536 : 0
    }));
    return {
      width,
      height,
      paths,
      connections,
      textRect: rect
        ? {
            l: coord(rect[0][0]) - left,
            t: coord(rect[0][1]) - top,
            r: coord(rect[1][0]) - left,
            b: coord(rect[1][1]) - top
          }
        : null,
      native: !vertexProp
    };
  };
})(typeof window !== 'undefined' ? window : globalThis);

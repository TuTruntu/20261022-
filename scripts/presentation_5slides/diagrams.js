/**
 * Native, editable scientific diagrams for the five introductory slides.
 * All functions draw inside the supplied rectangle (coordinates in inches).
 * Field colours and waveforms are qualitative illustrations, not solver data.
 */
'use strict';

const C = Object.freeze({
  blue: '165CA4', cyan: '2296BC', gold: 'D79B3A', navy: '143B63',
  pale: 'DCECF6', grid: 'A4BDCE', goldPale: 'F3E1BD', white: 'FFFFFF',
});

function checkBox(x, y, w, h) {
  if (![x, y, w, h].every(Number.isFinite) || w <= 0 || h <= 0) {
    throw new TypeError('Diagram bounds must be finite, with positive width and height.');
  }
}

function helpers(slide, pptx) {
  function shape(type, x, y, w, h, fill, stroke = fill, width = 0, extra = {}) {
    slide.addShape(type, {
      x, y, w: Math.max(w, 0.001), h: Math.max(h, 0.001),
      fill: { color: fill }, line: { color: stroke, width }, ...extra,
    });
  }
  function line(x1, y1, x2, y2, color = C.blue, width = 1, transparency = 0) {
    slide.addShape(pptx.ShapeType.line, {
      x: Math.min(x1, x2), y: Math.min(y1, y2),
      w: Math.max(Math.abs(x2 - x1), 0.001),
      h: Math.max(Math.abs(y2 - y1), 0.001),
      flipV: (x2 - x1) * (y2 - y1) < 0,
      line: { color, width, transparency, beginArrowType: 'none', endArrowType: 'none' },
    });
  }
  function poly(points, fill, stroke = fill, width = 0, transparency = 0) {
    const xs = points.map(p => p[0]), ys = points.map(p => p[1]);
    const x = Math.min(...xs), y = Math.min(...ys);
    slide.addShape(pptx.ShapeType.custGeom, {
      x, y, w: Math.max(Math.max(...xs) - x, 0.001),
      h: Math.max(Math.max(...ys) - y, 0.001),
      points: [...points.map(p => ({ x: p[0] - x, y: p[1] - y })), { close: true }],
      fill: { color: fill, transparency }, line: { color: stroke, width },
    });
  }
  function dot(cx, cy, radius, fill, stroke = fill, width = 0) {
    shape(pptx.ShapeType.ellipse, cx - radius, cy - radius, 2 * radius, 2 * radius, fill, stroke, width);
  }
  function arrow(p, q, shaft, headWidth, headLength, color = C.cyan) {
    const dx = q[0] - p[0], dy = q[1] - p[1], length = Math.hypot(dx, dy);
    const ux = dx / length, uy = dy / length, vx = -uy, vy = ux;
    const b = [q[0] - headLength * ux, q[1] - headLength * uy];
    const off = (r, a) => [r[0] + a * vx, r[1] + a * vy];
    poly([off(p, shaft / 2), off(b, shaft / 2), off(b, headWidth / 2), q,
      off(b, -headWidth / 2), off(b, -shaft / 2), off(p, -shaft / 2)], color);
  }
  return { shape, line, poly, dot, arrow };
}

function meshCube(slide, pptx, x, y, w, h) {
  checkBox(x, y, w, h);
  const { line, poly } = helpers(slide, pptx);
  const a = [x + .08 * w, y + .28 * h], u = [.66 * w, 0], v = [.20 * w, -.21 * h], z = [0, .63 * h];
  const P = (i, j, k) => [a[0] + u[0] * i + v[0] * j + z[0] * k,
    a[1] + u[1] * i + v[1] * j + z[1] * k];
  poly([P(0, 0, 0), P(1, 0, 0), P(1, 0, 1), P(0, 0, 1)], 'EDF5FA');
  poly([P(0, 0, 0), P(0, 1, 0), P(1, 1, 0), P(1, 0, 0)], 'DAEBF6');
  poly([P(1, 0, 0), P(1, 1, 0), P(1, 1, 1), P(1, 0, 1)], 'C8E0F0');
  const n = 7;
  for (let i = 0; i <= n; i++) {
    const t = i / n;
    line(...P(t, 0, 0), ...P(t, 0, 1), C.blue, .65);
    line(...P(0, 0, t), ...P(1, 0, t), C.blue, .65);
    line(...P(t, 0, 0), ...P(t, 1, 0), C.blue, .65);
    line(...P(0, t, 0), ...P(1, t, 0), C.blue, .65);
    line(...P(1, 0, t), ...P(1, 1, t), C.blue, .65);
    line(...P(1, t, 0), ...P(1, t, 1), C.blue, .65);
  }
}

function network(slide, pptx, x, y, w, h) {
  checkBox(x, y, w, h);
  const { line, dot } = helpers(slide, pptx);
  const counts = [3, 4, 3], xx = [.14, .50, .86];
  const r = Math.min(w * .046, h * .068);
  const pts = counts.map((n, k) => Array.from({ length: n }, (_, i) =>
    [x + xx[k] * w, y + (.11 + .78 * (i + .5) / n) * h]));
  for (let k = 0; k < 2; k++) {
    for (const p of pts[k]) for (const q of pts[k + 1]) line(...p, ...q, C.grid, .7);
  }
  pts.forEach((layer, k) => layer.forEach((p, i) =>
    dot(...p, r, k === 1 && i === 1 ? C.gold : k === 1 ? C.cyan : C.blue, C.white, .5)));
}

function geometryMeasures(slide, pptx, x, y, w, h) {
  checkBox(x, y, w, h);
  const { line, poly, dot } = helpers(slide, pptx);
  const names = ['体', '面', '线', '点'];
  const labelH = Math.min(.23, h * .30), dh = h - labelH;
  for (let k = 0; k < 4; k++) {
    const bx = x + k * w / 4, bw = w / 4;
    const P = (a, b, c) => [bx + bw * (.14 + .55 * a + .20 * b),
      y + dh * (.92 - .18 * b - .66 * c)];
    // Identical cells make the dimensions of their contained objects comparable.
    for (const [p, q] of [
      [[0, 0, 0], [1, 0, 0]], [[1, 0, 0], [1, 1, 0]], [[1, 1, 0], [0, 1, 0]], [[0, 1, 0], [0, 0, 0]],
      [[0, 0, 1], [1, 0, 1]], [[1, 0, 1], [1, 1, 1]], [[1, 1, 1], [0, 1, 1]], [[0, 1, 1], [0, 0, 1]],
      [[0, 0, 0], [0, 0, 1]], [[1, 0, 0], [1, 0, 1]], [[1, 1, 0], [1, 1, 1]],
    ]) line(...P(...p), ...P(...q), C.grid, .65);
    if (k === 0) {
      poly([P(.16, .16, .18), P(.81, .16, .18), P(.81, .16, .83), P(.16, .16, .83)], 'C4DFF0', C.blue, .8);
      poly([P(.16, .16, .83), P(.16, .81, .83), P(.81, .81, .83), P(.81, .16, .83)], 'E5F0F7', C.blue, .8);
      poly([P(.81, .16, .18), P(.81, .81, .18), P(.81, .81, .83), P(.81, .16, .83)], '9FCBE3', C.blue, .8);
    } else if (k === 1) {
      poly([P(.14, .12, .22), P(.86, .12, .56), P(.86, .86, .75), P(.14, .86, .42)], 'BFDFEF', C.cyan, 1.1);
    } else if (k === 2) {
      line(...P(.17, .16, .18), ...P(.83, .82, .84), C.gold, 3.0);
    } else {
      const p = P(.52, .48, .52);
      dot(...p, Math.min(bw * .09, dh * .075), C.gold, C.white, .5);
    }
    slide.addText(names[k], {
      x: bx, y: y + h - labelH, w: bw, h: labelH,
      fontFace: 'Noto Sans CJK SC', fontSize: 14, color: C.navy,
      align: 'center', valign: 'mid', margin: 0, breakLine: false,
    });
  }
}

function fieldCube(slide, pptx, x, y, w, h) {
  checkBox(x, y, w, h);
  const { line, poly, arrow } = helpers(slide, pptx);
  const P = (a, b, c) => [x + w * (.07 + .68 * a + .22 * b),
    y + h * (.93 - .19 * b - .66 * c)];
  const color = v => {
    const stops = [[0, [45, 113, 175]], [.48, [51, 167, 198]], [.76, [167, 218, 225]], [1, [231, 175, 79]]];
    for (let i = 0; i < stops.length - 1; i++) {
      const [a, ca] = stops[i], [b, cb] = stops[i + 1];
      if (v >= a && v <= b) {
        const t = (v - a) / (b - a);
        return ca.map((n, k) => Math.round(n + (cb[k] - n) * t).toString(16).padStart(2, '0')).join('');
      }
    }
    return C.gold;
  };
  for (const [axis, fixed] of [['z', .10], ['y', .55], ['x', .56]]) {
    const point = (u, v) => axis === 'z' ? P(u, v, fixed) : axis === 'y' ? P(u, fixed, v) : P(fixed, u, v);
    const n = 6;
    for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) {
      const u0 = i / n, u1 = (i + 1) / n, v0 = j / n, v1 = (j + 1) / n;
      const u = (u0 + u1) / 2, v = (v0 + v1) / 2;
      const value = .19 + .77 * Math.exp(-((u - .54) ** 2 / .065 + (v - .49) ** 2 / .16));
      poly([point(u0, v0), point(u1, v0), point(u1, v1), point(u0, v1)], color(value));
    }
  }
  const corners = [[0, 0, 0], [1, 0, 0], [1, 1, 0], [0, 1, 0], [0, 0, 1], [1, 0, 1], [1, 1, 1], [0, 1, 1]].map(v => P(...v));
  for (const [i, j] of [[0, 1], [1, 2], [2, 3], [3, 0], [4, 5], [5, 6], [6, 7], [7, 4], [0, 4], [1, 5], [2, 6], [3, 7]]) {
    line(...corners[i], ...corners[j], C.grid, .65);
  }
  for (const [a, b, c] of [[.28, .02, .32], [.49, .02, .48], [.70, .02, .61]]) {
    const p = P(a, b, c), q = P(a + .10, b, c + .10);
    arrow(p, q, Math.min(w, h) * .008, Math.min(w, h) * .034, Math.min(w, h) * .035, C.white);
  }
}

function transfer(slide, pptx, x, y, w, h) {
  checkBox(x, y, w, h);
  const { line, dot, arrow } = helpers(slide, pptx);
  const r = Math.min(w * .020, h * .046);
  function smallNet(left, col) {
    const xx = [left, left + .12 * w, left + .24 * w];
    const pts = xx.map((a, k) => Array.from({ length: k === 1 ? 3 : 2 }, (_, i) =>
      [a, y + (.43 + .35 * i / (k === 1 ? 2 : 1)) * h]));
    for (let k = 0; k < 2; k++) for (const p of pts[k]) for (const q of pts[k + 1]) line(...p, ...q, C.grid, .7);
    pts.forEach((layer, k) => layer.forEach(p => dot(...p, r, k === 1 ? col : C.blue, C.white, .3)));
  }
  smallNet(x + .08 * w, C.cyan);
  smallNet(x + .68 * w, C.gold);
  arrow([x + .41 * w, y + .59 * h], [x + .59 * w, y + .59 * h], .055 * h, .17 * h, .10 * w, C.cyan);
  // Source and target frequency patterns differ, but their model structure is shared.
  for (const [left, cycles, col] of [[.055, 1.5, C.blue], [.655, 2.5, C.gold]]) {
    const pts = Array.from({ length: 31 }, (_, i) => [x + (left + .29 * i / 30) * w,
      y + (.22 - .07 * Math.sin(2 * Math.PI * cycles * i / 30)) * h]);
    for (let i = 1; i < pts.length; i++) line(...pts[i - 1], ...pts[i], col, 1.15);
  }
}

function ambiguousBoundary(slide, pptx, x, y, w, h) {
  checkBox(x, y, w, h);
  const { line, poly } = helpers(slide, pptx);
  const pw = .40 * w, ph = .81 * h, gy = y + .095 * h;
  for (const [left, angle] of [[x + .025 * w, 0], [x + .575 * w, Math.PI / 4]]) {
    const gx = left, cx = gx + pw / 2, cy = gy + ph / 2;
    for (let i = 0; i <= 5; i++) {
      line(gx + pw * i / 5, gy, gx + pw * i / 5, gy + ph, C.grid, .6);
      line(gx, gy + ph * i / 5, gx + pw, gy + ph * i / 5, C.grid, .6);
    }
    // Same unrotated dimensions => exactly the same area for both boundaries.
    const side = Math.min(pw, ph), a = .31 * side, b = .125 * side;
    const pts = [[-a, -b], [a, -b], [a, b], [-a, b]].map(([u, v]) =>
      [cx + u * Math.cos(angle) - v * Math.sin(angle), cy + u * Math.sin(angle) + v * Math.cos(angle)]);
    poly(pts, 'C3E0EF', C.gold, 1.6);
  }
  const mid = x + .50 * w;
  line(mid - .035 * w, y + .44 * h, mid + .035 * w, y + .44 * h, C.navy, 1.4);
  line(mid - .035 * w, y + .56 * h, mid + .035 * w, y + .56 * h, C.navy, 1.4);
}

module.exports = { meshCube, network, geometryMeasures, fieldCube, transfer, ambiguousBoundary };

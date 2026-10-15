/* Orion stereonet maths (pure functions, no DOM). Vectors are ENU: x = East, y = North, z = Up.
   Lower-hemisphere, equal-area (Lambert azimuthal) projection. Used by app.js and tests/stereo.test.js. */
(function (root) {
  const D2R = Math.PI / 180, R2D = 180 / Math.PI;
  const norm360 = (a) => ((a % 360) + 360) % 360;
  const clamp1 = (x) => Math.max(-1, Math.min(1, x));
  const dot = (a, b) => a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
  const unit = (a) => { const l = Math.hypot(a[0], a[1], a[2]) || 1; return [a[0] / l, a[1] / l, a[2] / l]; };

  /* ---- conversions ---- */
  // Right-hand rule: the plane dips 90 deg clockwise from strike.
  const strikeToDipDir = (strike) => norm360(strike + 90);
  const dipDirToStrike = (dipDir) => norm360(dipDir - 90);

  // Pole to a plane, as a lower-hemisphere line.
  function poleFromPlane(dip, dipDir) {
    return { trend: norm360(dipDir + 180), plunge: 90 - dip };
  }
  function trendPlungeToVec(trend, plunge) {
    const t = trend * D2R, p = plunge * D2R;
    return [Math.cos(p) * Math.sin(t), Math.cos(p) * Math.cos(t), -Math.sin(p)];
  }
  function vecToTrendPlunge(v) {
    const u = v[2] > 0 ? [-v[0], -v[1], -v[2]] : v; // force lower hemisphere
    return { trend: norm360(Math.atan2(u[0], u[1]) * R2D), plunge: Math.asin(clamp1(-u[2])) * R2D };
  }
  function poleToPlane(trend, plunge) {
    return { dip: 90 - plunge, dipDir: norm360(trend + 180) };
  }
  const poleVec = (dip, dipDir) => { const p = poleFromPlane(dip, dipDir); return trendPlungeToVec(p.trend, p.plunge); };

  /* ---- projection ---- */
  // Lambert azimuthal equal-area, lower hemisphere. Primitive circle has radius 1. Returns [x east, y north].
  function project(v) {
    const u = v[2] > 0 ? [-v[0], -v[1], -v[2]] : v;
    const h = Math.hypot(u[0], u[1]);
    if (h < 1e-12) return [0, 0];
    const r = Math.SQRT2 * Math.sin(Math.acos(clamp1(-u[2])) / 2);
    return [(r * u[0]) / h, (r * u[1]) / h];
  }
  // Inverse: net coordinates (radius <= 1) -> lower-hemisphere unit vector.
  function unproject(x, y) {
    const r2 = x * x + y * y;
    if (r2 > 1) return null;
    const k = Math.sqrt(2 - r2); // r^2 = 1 - cos(theta), sin(theta) = r*sqrt(2 - r^2)
    return [x * k, y * k, -(1 - r2)];
  }

  /* ---- great circle (rotation of the strike line about the pole) ---- */
  function greatCircle(dip, dipDir, n) {
    n = n || 90;
    const a = dipDir * D2R, d = dip * D2R;
    const s = [-Math.cos(a), Math.sin(a), 0]; // strike direction (RHR)
    const dn = [Math.sin(a) * Math.cos(d), Math.cos(a) * Math.cos(d), -Math.sin(d)]; // down-dip direction
    const pts = [];
    for (let i = 0; i <= n; i++) {
      const t = (Math.PI * i) / n, c = Math.cos(t), sn = Math.sin(t);
      pts.push([c * s[0] + sn * dn[0], c * s[1] + sn * dn[1], c * s[2] + sn * dn[2]]);
    }
    return pts;
  }

  /* ---- mean plane ---- */
  // Principal axis of the pole scatter matrix: used only to give every pole the same sense (axial data).
  function principalAxis(P) {
    const T = [[0, 0, 0], [0, 0, 0], [0, 0, 0]];
    P.forEach((p) => { for (let i = 0; i < 3; i++) for (let j = 0; j < 3; j++) T[i][j] += p[i] * p[j]; });
    let v = unit(P[0]);
    for (let k = 0; k < 200; k++) v = unit([0, 1, 2].map((i) => T[i][0] * v[0] + T[i][1] * v[1] + T[i][2] * v[2]));
    return v;
  }
  // Convert planes -> pole vectors, align them to one hemisphere about the principal axis, average, normalise.
  function meanPlane(planes) {
    if (!planes.length) return null;
    const P = planes.map((p) => poleVec(p.dip, p.dipDir));
    const ax = principalAxis(P);
    const sum = [0, 0, 0];
    P.forEach((p) => { const s = dot(p, ax) < 0 ? -1 : 1; sum[0] += s * p[0]; sum[1] += s * p[1]; sum[2] += s * p[2]; });
    const R = Math.hypot(sum[0], sum[1], sum[2]);
    let v = R > 1e-12 ? unit(sum) : ax;
    if (v[2] > 0) v = v.map((x) => -x);
    const tp = vecToTrendPlunge(v), pl = poleToPlane(tp.trend, tp.plunge);
    return { n: planes.length, pole: v, trend: tp.trend, plunge: tp.plunge, dip: pl.dip, dipDir: pl.dipDir, strike: dipDirToStrike(pl.dipDir), resultant: R / planes.length };
  }

  /* ---- pole density (axial Gaussian / Fisher-type kernel, Kamb-like smoothing) ---- */
  function kernelKappa(n) { return Math.max(3, Math.min(60, (n + 9) / 3)); }
  // Returns { grid, size, max } where grid[j*size+i] is density at net coords x=-1+2i/(size-1), y=1-2j/(size-1) (NaN outside).
  function density(poles, size, kappa) {
    size = size || 80;
    kappa = kappa || kernelKappa(poles.length);
    const grid = new Float32Array(size * size).fill(NaN);
    let max = 0;
    for (let j = 0; j < size; j++) for (let i = 0; i < size; i++) {
      const v = unproject(-1 + (2 * i) / (size - 1), 1 - (2 * j) / (size - 1));
      if (!v) continue;
      let s = 0;
      for (const p of poles) s += Math.exp(kappa * (Math.abs(dot(v, p)) - 1)); // |.| handles axial data at the primitive
      grid[j * size + i] = s;
      if (s > max) max = s;
    }
    return { grid, size, max };
  }
  // Marching squares: returns line segments [[x1,y1,x2,y2],...] in grid index space for one level.
  function contourSegments(grid, size, level) {
    const segs = [];
    const val = (i, j) => grid[j * size + i];
    const lerp = (a, b) => (level - a) / (b - a);
    for (let j = 0; j < size - 1; j++) for (let i = 0; i < size - 1; i++) {
      const a = val(i, j), b = val(i + 1, j), c = val(i + 1, j + 1), d = val(i, j + 1);
      if (!(a === a && b === b && c === c && d === d)) continue;
      const idx = (a >= level ? 8 : 0) | (b >= level ? 4 : 0) | (c >= level ? 2 : 0) | (d >= level ? 1 : 0);
      if (idx === 0 || idx === 15) continue;
      const T = [i + lerp(a, b), j], R = [i + 1, j + lerp(b, c)], B = [i + lerp(d, c), j + 1], L = [i, j + lerp(a, d)];
      const add = (p, q) => segs.push([p[0], p[1], q[0], q[1]]);
      switch (idx) {
        case 1: case 14: add(L, B); break;
        case 2: case 13: add(B, R); break;
        case 3: case 12: add(L, R); break;
        case 4: case 11: add(T, R); break;
        case 5: add(T, L); add(B, R); break;
        case 6: case 9: add(T, B); break;
        case 7: case 8: add(T, L); break;
        case 10: add(T, R); add(L, B); break;
      }
    }
    return segs;
  }
  // Light blue -> deep blue ramp, t in [0,1].
  function blueRGB(t) {
    t = Math.max(0, Math.min(1, t));
    const a = [191, 225, 255], b = [8, 48, 140];
    return [Math.round(a[0] + (b[0] - a[0]) * t), Math.round(a[1] + (b[1] - a[1]) * t), Math.round(a[2] + (b[2] - a[2]) * t)];
  }
  const blueRamp = (t) => `rgb(${blueRGB(t).join(',')})`;

  const api = { norm360, strikeToDipDir, dipDirToStrike, poleFromPlane, trendPlungeToVec, vecToTrendPlunge, poleToPlane, poleVec, project, unproject, greatCircle, meanPlane, kernelKappa, density, contourSegments, blueRGB, blueRamp };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.OrionStereo = api;
})(typeof window !== 'undefined' ? window : globalThis);

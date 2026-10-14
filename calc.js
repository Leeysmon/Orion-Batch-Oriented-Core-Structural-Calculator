/* Orion core-structure maths. Coordinates are ENU: x = East, y = North, z = Up. */
(function () {
  const D2R = Math.PI / 180, R2D = 180 / Math.PI;
  const norm360 = (a) => ((a % 360) + 360) % 360;
  const dot = (a, b) => a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
  const cross = (a, b) => [a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0]];
  const mul = (a, k) => [a[0] * k, a[1] * k, a[2] * k];
  const add = (a, b) => [a[0] + b[0], a[1] + b[1], a[2] + b[2]];
  const sub = (a, b) => [a[0] - b[0], a[1] - b[1], a[2] - b[2]];
  const len = (a) => Math.hypot(a[0], a[1], a[2]);
  const unit = (a) => { const l = len(a); return l ? mul(a, 1 / l) : a; };

  // Azimuth clockwise from north; incDown = degrees below horizontal (positive = down).
  function dirVector(az, incDown) {
    const a = az * D2R, i = incDown * D2R;
    return [Math.sin(a) * Math.cos(i), Math.cos(a) * Math.cos(i), -Math.sin(i)];
  }

  // Right-hand rotation of v about unit axis k.
  function rotate(v, k, ang) {
    const c = Math.cos(ang), s = Math.sin(ang);
    return add(add(mul(v, c), mul(cross(k, v), s)), mul(k, dot(k, v) * (1 - c)));
  }

  /*
   * Alpha/beta -> true plane orientation.
   * alpha: acute angle between plane and core axis (0-90).
   * beta: clockwise (looking down-hole) from the bottom-of-hole line to the lowest point
   * of the ellipse (point furthest down-hole). Right-hand rule: dip direction, dip.
   */
  function planeFromAlphaBeta(d, az, alpha, beta) {
    d = unit(d);
    const down = [0, 0, -1];
    let b = sub(down, mul(d, dot(down, d)));
    if (len(b) < 1e-6) {
      const h = [Math.sin(az * D2R), Math.cos(az * D2R), 0];
      b = sub(h, mul(d, dot(h, d)));
    }
    b = unit(b);
    const w = rotate(b, d, beta * D2R);
    const al = alpha * D2R;
    const q = add(mul(d, Math.cos(al)), mul(w, Math.sin(al)));
    const s = cross(d, w);
    let n = unit(cross(q, s));
    if (n[2] < 0) n = mul(n, -1);
    const dip = Math.acos(Math.min(1, Math.max(-1, n[2]))) * R2D;
    const dipDir = dip < 1e-9 ? 0 : norm360(Math.atan2(n[0], n[1]) * R2D);
    const strike = norm360(dipDir - 90);
    return { normal: n, dip, dipDir, strike };
  }

  // Unit vectors lying in the plane (for drawing).
  function planeBasis(dipDir, dip) {
    const dd = dipDir * D2R, dp = dip * D2R, st = norm360(dipDir - 90) * D2R;
    return {
      down: [Math.sin(dd) * Math.cos(dp), Math.cos(dd) * Math.cos(dp), -Math.sin(dp)],
      strike: [Math.sin(st), Math.cos(st), 0],
    };
  }

  function mcStep(v1, v2, L) {
    const c = Math.min(1, Math.max(-1, dot(v1, v2)));
    const delta = Math.acos(c);
    const rf = delta < 1e-8 ? 1 : (2 / delta) * Math.tan(delta / 2);
    return mul(add(v1, v2), (L / 2) * rf);
  }

  function slerpish(v1, v2, t) {
    const v = add(mul(v1, 1 - t), mul(v2, t));
    return len(v) < 1e-9 ? v1 : unit(v);
  }

  /*
   * Drillhole trace by minimum curvature.
   * stations: [{depth, az, inc}] with inc positive downwards.
   */
  function buildTrace(collar, stations) {
    const st = stations.slice().sort((a, b) => a.depth - b.depth).map((s) => ({
      depth: s.depth, az: s.az, inc: s.inc, v: dirVector(s.az, s.inc),
    }));
    if (!st.length) return null;
    if (st[0].depth > 0) st.unshift({ ...st[0], depth: 0 });
    const pos = [collar.slice()];
    for (let i = 1; i < st.length; i++) {
      pos.push(add(pos[i - 1], mcStep(st[i - 1].v, st[i].v, st[i].depth - st[i - 1].depth)));
    }
    function at(depth) {
      depth = Math.max(0, depth);
      let i = st.length - 1;
      for (let k = 0; k < st.length - 1; k++) {
        if (depth <= st[k + 1].depth) { i = k; break; }
      }
      const a = st[i];
      if (i === st.length - 1 && depth >= a.depth) {
        const L = depth - a.depth;
        return { pos: add(pos[i], mul(a.v, L)), dir: a.v, az: a.az, inc: a.inc };
      }
      const b = st[i + 1];
      const span = b.depth - a.depth;
      const t = span > 0 ? (depth - a.depth) / span : 0;
      const dir = slerpish(a.v, b.v, t);
      const p = add(pos[i], mcStep(a.v, dir, depth - a.depth));
      const horiz = Math.hypot(dir[0], dir[1]);
      return {
        pos: p, dir,
        az: horiz < 1e-9 ? a.az : norm360(Math.atan2(dir[0], dir[1]) * R2D),
        inc: -Math.asin(Math.max(-1, Math.min(1, dir[2]))) * R2D,
      };
    }
    return { stations: st, at };
  }

  window.Calc = { norm360, dirVector, planeFromAlphaBeta, planeBasis, buildTrace, D2R, R2D };
})();

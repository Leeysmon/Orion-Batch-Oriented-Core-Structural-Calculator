// Run: node tests/stereo.test.js
const assert = require('assert');
const S = require('../stereo.js');
const near = (a, b, tol = 1e-6, msg) => assert.ok(Math.abs(a - b) <= tol, `${msg || ''} expected ${b}, got ${a}`);
const nearAz = (a, b, tol = 1e-6, msg) => near(((a - b + 540) % 360) - 180, 0, tol, msg);
let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok -', name); };

t('strike -> dip direction (RHR)', () => {
  near(S.strikeToDipDir(0), 90); near(S.strikeToDipDir(270), 0); near(S.strikeToDipDir(350), 80);
  near(S.dipDirToStrike(S.strikeToDipDir(123)), 123);
});

t('pole of plane: plunge = 90 - dip, trend = dipDir + 180', () => {
  const p = S.poleFromPlane(30, 100); near(p.plunge, 60); near(p.trend, 280);
  near(S.poleFromPlane(45, 270).trend, 90);
});

t('trend/plunge <-> vector round trip, lower hemisphere', () => {
  [[0, 0], [90, 30], [225, 75], [359, 10]].forEach(([tr, pl]) => {
    const v = S.trendPlungeToVec(tr, pl); assert.ok(v[2] <= 0);
    const b = S.vecToTrendPlunge(v); nearAz(b.trend, tr, 1e-6); near(b.plunge, pl);
  });
  const e = S.trendPlungeToVec(90, 0); near(e[0], 1); near(e[1], 0, 1e-12);
  const nth = S.trendPlungeToVec(0, 0); near(nth[1], 1);
});

t('vector pointing up is flipped to the lower hemisphere', () => {
  near(S.vecToTrendPlunge([0, 0, 1]).plunge, 90);
});

t('plane strike 000 / dip 45 (dips east): pole in the west, great circle in the east', () => {
  const dd = S.strikeToDipDir(0); near(dd, 90);
  const pv = S.poleVec(45, dd); near(pv[0], -Math.SQRT1_2); near(pv[1], 0, 1e-12); near(pv[2], -Math.SQRT1_2);
  assert.ok(S.project(pv)[0] < 0);
  const gc = S.greatCircle(45, dd);
  assert.ok(S.project(gc[45])[0] > 0, 'great circle mid-point should lie east');
  gc.forEach((v) => near(v[0] * pv[0] + v[1] * pv[1] + v[2] * pv[2], 0, 1e-12, 'great circle perpendicular to pole'));
});

t('great circle: end points are the strike line on the primitive, all points lower hemisphere', () => {
  const gc = S.greatCircle(60, 200); const a = S.project(gc[0]), b = S.project(gc[gc.length - 1]);
  near(Math.hypot(...a), 1, 1e-9); near(Math.hypot(...b), 1, 1e-9);
  near(a[0] + b[0], 0, 1e-9); near(a[1] + b[1], 0, 1e-9);
  gc.forEach((v) => assert.ok(v[2] <= 1e-12));
});

t('equal-area projection radii', () => {
  near(Math.hypot(...S.project([0, 0, -1])), 0);
  near(Math.hypot(...S.project([1, 0, 0])), 1);
  near(Math.hypot(...S.project(S.trendPlungeToVec(0, 45))), Math.SQRT2 * Math.sin(Math.PI / 8));
  // a cap 60 deg from the nadir has half the hemisphere area, so r^2 = 0.5
  near(Math.hypot(...S.project(S.trendPlungeToVec(10, 30))) ** 2, 0.5, 1e-9);
});

t('project/unproject round trip', () => {
  const v = S.trendPlungeToVec(137, 42), p = S.project(v), u = S.unproject(p[0], p[1]);
  [0, 1, 2].forEach((i) => near(u[i], v[i], 1e-9));
  assert.strictEqual(S.unproject(1.1, 0), null);
});

t('mean plane of identical planes returns that plane', () => {
  const m = S.meanPlane([{ dip: 40, dipDir: 120 }, { dip: 40, dipDir: 120 }]);
  near(m.dip, 40); nearAz(m.dipDir, 120); nearAz(m.strike, 30);
  near(m.plunge, 50); nearAz(m.trend, 300); near(m.resultant, 1, 1e-9);
});

t('mean plane averages poles (symmetric pair about 100 -> dip dir 100)', () => {
  const m = S.meanPlane([{ dip: 50, dipDir: 90 }, { dip: 50, dipDir: 110 }]);
  nearAz(m.dipDir, 100, 0.5); assert.ok(m.dip > 49 && m.dip <= 50);
});

t('mean plane handles wrap-around at north (350 & 10)', () => {
  nearAz(S.meanPlane([{ dip: 30, dipDir: 350 }, { dip: 30, dipDir: 10 }]).dipDir, 0);
});

t('mean plane handles steep planes whose poles straddle the primitive', () => {
  const m = S.meanPlane([{ dip: 85, dipDir: 90 }, { dip: 85, dipDir: 270 }]);
  near(Math.min(m.strike % 180, 180 - (m.strike % 180)), 0, 0.001); near(m.dip, 90, 1e-6); // same strike, opposite dips -> vertical
});

t('mean plane is invariant to input order', () => {
  const a = [{ dip: 20, dipDir: 40 }, { dip: 35, dipDir: 70 }, { dip: 60, dipDir: 300 }];
  const m1 = S.meanPlane(a), m2 = S.meanPlane([a[2], a[0], a[1]]);
  near(m1.dip, m2.dip, 1e-9); nearAz(m1.dipDir, m2.dipDir, 1e-9);
});

t('density peaks at a tight cluster of poles and is NaN outside the net', () => {
  const poles = Array.from({ length: 10 }, (_, i) => S.poleVec(40 + (i % 3), 120 + (i % 4)));
  const d = S.density(poles, 61);
  let bi = -1; d.grid.forEach((v, i) => { if (v === v && (bi < 0 || v > d.grid[bi])) bi = i; });
  const pk = S.vecToTrendPlunge(S.unproject(-1 + (2 * (bi % 61)) / 60, 1 - (2 * Math.floor(bi / 61)) / 60));
  nearAz(pk.trend, 300, 8); near(pk.plunge, 50, 6);
  assert.ok(Number.isNaN(d.grid[0]));
});

t('contour segments surround a peak; none above the maximum', () => {
  const d = S.density([S.poleVec(30, 90), S.poleVec(32, 92), S.poleVec(28, 88)], 60);
  assert.ok(S.contourSegments(d.grid, d.size, d.max * 0.5).length > 10);
  assert.strictEqual(S.contourSegments(d.grid, d.size, d.max * 1.01).length, 0);
});

t('Wulff (equal-angle) projection: r = tan(theta/2); primitive stays radius 1', () => {
  near(Math.hypot(...S.project(S.trendPlungeToVec(0, 45), 'wulff')), Math.tan(Math.PI / 8));
  near(Math.hypot(...S.project([1, 0, 0], 'wulff')), 1);
});

t('net grid: every point is lower hemisphere; meridians meet at N and S; N-S axis small circle is a point-pair', () => {
  const g = S.netGrid(10, 90);
  assert.strictEqual(g.length, 34);
  g.forEach((ln) => ln.forEach((v) => { assert.ok(v[2] <= 1e-12); near(Math.hypot(...v), 1, 1e-12); }));
  const m = g[0]; near(m[0][1], -1, 1e-12); near(m[90][1], 1, 1e-12); // first meridian runs S -> N
  const p = S.project(m[90]); near(p[0], 0, 1e-12); near(p[1], 1, 1e-12); // north end sits at top of the net
});

t('blue ramp runs light blue -> deep blue', () => {
  assert.strictEqual(S.blueRamp(0), 'rgb(191,225,255)'); assert.strictEqual(S.blueRamp(1), 'rgb(8,48,140)');
});

console.log(`\n${n} tests passed`);

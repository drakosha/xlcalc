// What the scripts that measure two builds of xlcalc against each other in one process share: the clock they
// compute at, the builds' names, a quantile by its index (the same p95 as its median), the turn order the builds
// alternate in, and loading the two builds (with `synth`, their synth.js beside index.js).
export const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
export const names = ['base', 'branch']
export const at = (xs, q) => [...xs].sort((a, b) => a - b)[Math.min(xs.length - 1, Math.floor(xs.length * q))]
export const pair = (round) => (round % 2 ? [1, 0] : [0, 1])
export const load = (base, branch, synth = false) =>
  Promise.all([base, branch].map(async (r) => (synth ? { ...(await import(r + '/dist/index.js')), ...(await import(r + '/dist/synth.js')) } : await import(r + '/dist/index.js'))))

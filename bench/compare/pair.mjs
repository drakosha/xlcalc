// What the scripts that measure two builds of xlcalc against each other in one process share: the clock they
// compute at, the builds' names, the quantile the prints go by (`stats.mjs`'s `at`), the turn order the builds
// alternate in, and loading the two builds (with `synth`, their synth.js beside index.js).
import { at } from '../stats.mjs'

export { at }
export const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
export const names = ['base', 'branch']
export const pair = (round) => (round % 2 ? [1, 0] : [0, 1])
export const load = (base, branch, synth = false) =>
  Promise.all([base, branch].map(async (r) => (synth ? { ...(await import(r + '/dist/index.js')), ...(await import(r + '/dist/synth.js')) } : await import(r + '/dist/index.js'))))

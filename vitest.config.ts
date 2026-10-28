import { fileURLToPath } from 'node:url'
import { configDefaults, defineConfig } from 'vitest/config'

// FORMULA_TARGET=dist (`npm run test:dist`): tests import the built dist instead of src, and so do the
// child processes (test/fixtures/target.ts).
const dist = fileURLToPath(new URL('./dist/', import.meta.url))

/**
 * The files whose tests check how long something takes (a slice, a deadline, a budget): project `timed`, which
 * `npm test` runs after the others (`test/suites.mjs`), one file at a time, so that the suite's own load fails none.
 */
const TIMED = [
  'test/arrays-spill.test.ts',
  'test/async-inline.test.ts',
  'test/async-inline-long.test.ts',
  'test/async-recovery.test.ts',
  'test/budgets.test.ts',
  'test/end.test.ts',
  'test/limits-writes.test.ts',
  'test/regex-levels.test.ts',
  'test/session-structure.test.ts',
  'test/session-undo-async.test.ts',
  'test/spill-occupied.test.ts',
  'test/style-table.test.ts',
  'test/worker-browser.test.ts',
  'test/worker-pack.test.ts',
  'test/xlsx-budget.test.ts',
]

export default defineConfig({
  test: {
    testTimeout: 15_000,
    projects: [
      { extends: true, test: { name: 'main', include: ['test/**/*.test.ts', 'bench/**/*.test.ts'], exclude: [...configDefaults.exclude, ...TIMED] } },
      { extends: true, test: { name: 'timed', include: TIMED, fileParallelism: false } },
    ],
  },
  resolve: {
    alias: process.env.FORMULA_TARGET === 'dist' ? [{ find: /^\.\.\/src\/(.*)$/, replacement: `${dist}$1` }] : [],
  },
})

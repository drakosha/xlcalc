import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

// FORMULA_TARGET=dist (`npm run test:dist`): tests import the built dist instead of src, and so do the
// child processes (test/fixtures/target.ts).
const dist = fileURLToPath(new URL('./dist/', import.meta.url))

export default defineConfig({
  test: {
    include: ['test/**/*.test.ts'],
    testTimeout: 15_000,
  },
  resolve: {
    alias: process.env.FORMULA_TARGET === 'dist' ? [{ find: /^\.\.\/src\/(.*)$/, replacement: `${dist}$1` }] : [],
  },
})

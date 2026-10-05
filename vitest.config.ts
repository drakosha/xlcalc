import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

// FORMULA_TARGET=dist (`npm run test:dist`): тесты импортируют собранный dist вместо src, дочерние
// процессы тоже (test/fixtures/target.ts).
const dist = fileURLToPath(new URL('./dist/', import.meta.url))

export default defineConfig({
  test: {
    include: ['test/**/*.test.ts', 'bench/**/*.test.ts'],
    testTimeout: 15_000,
  },
  resolve: {
    alias: process.env.FORMULA_TARGET === 'dist' ? [{ find: /^\.\.\/src\/(.*)$/, replacement: `${dist}$1` }] : [],
  },
})

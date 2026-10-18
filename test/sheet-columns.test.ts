/**
 * A sheet's values stay kept by index however far apart its cells are: a column of 34,000 values with one cell far below
 * them, or a cell every 50,000 rows down to the sheet's last, leaves no array of the sheet's storage a dictionary of the
 * JavaScript engine (each read of such a column was a lookup by hash: an amount edit of the synthetic book of 34,000 rows
 * with a cell in row 100,000 of its data columns took 14.3 ms for 8.5), and the values read back as written. In a fresh
 * process with `--allow-natives-syntax` (`fixtures/columns-child.ts`).
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { expect, it } from 'vitest'

const child = fileURLToPath(new URL('./fixtures/columns-child.ts', import.meta.url))

it('keeps the values of a sheet with far apart cells by index', () => {
  const out = JSON.parse(execFileSync(process.execPath, ['--allow-natives-syntax', '--import', 'tsx', child], { encoding: 'utf8' })) as { dictionaries: number; read: unknown[] }
  expect(out.read).toEqual([1.5, 51_000, 'note', 50_001, 1_000_001, null])
  expect(out.dictionaries).toBe(0)
}, 60_000)

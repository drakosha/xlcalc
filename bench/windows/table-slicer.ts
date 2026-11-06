/**
 * Runs `table-slicer.ps1` on the machine and stores what Excel saved in `test/excel/win/table-slicer/`, with its answers
 * in `answers.json`: `XLWIN_HOST=… XLWIN_CODEPAGE=any npx tsx bench/windows/table-slicer.ts`.
 */
import { execFileSync } from 'node:child_process'
import { copyFileSync, mkdirSync, mkdtempSync, readdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { HOST, metaOf, ssh, withWindowsExcel } from './remote.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const OUT = join(HERE, '../../test/excel/win/table-slicer')
const WORK = 'C:\\xlcalc\\work\\table-slicer'

await withWindowsExcel('en-US', async (e) => {
  execFileSync('scp', [join(HERE, 'table-slicer.ps1'), `${HOST}:C:/xlcalc/work/table-slicer.ps1`])
  const cases = JSON.parse(ssh(`powershell -NoProfile -ExecutionPolicy Bypass -File C:\\xlcalc\\work\\table-slicer.ps1 -dir ${WORK}`, 300_000)) as unknown
  const local = mkdtempSync(join(tmpdir(), 'table-slicer-'))
  execFileSync('scp', ['-r', `${HOST}:C:/xlcalc/work/table-slicer`, local])
  ssh(`Remove-Item -Recurse -Force ${WORK}, C:\\xlcalc\\work\\table-slicer.ps1`)
  mkdirSync(OUT, { recursive: true })
  for (const f of readdirSync(join(local, 'table-slicer'))) copyFileSync(join(local, 'table-slicer', f), join(OUT, f))
  writeFileSync(join(OUT, 'answers.json'), `${JSON.stringify({ meta: metaOf(e, 'en-US'), cases }, null, 1)}\n`)
})

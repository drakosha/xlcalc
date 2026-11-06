/**
 * Runs `filter-criteria.ps1` on the machine and stores what Excel saved in `test/excel/win/filter-criteria-made/`, with its
 * answers in `answers.json`: `XLWIN_HOST=… XLWIN_CODEPAGE=any npx tsx bench/windows/filter-criteria.ts`.
 */
import { execFileSync } from 'node:child_process'
import { copyFileSync, mkdirSync, mkdtempSync, readdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { HOST, metaOf, ssh, withWindowsExcel } from './remote.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const OUT = join(HERE, '../../test/excel/win/filter-criteria-made')
const WORK = 'C:\\xlcalc\\work\\filter-criteria'

await withWindowsExcel('en-US', async (e) => {
  execFileSync('scp', [join(HERE, 'filter-criteria.ps1'), `${HOST}:C:/xlcalc/work/filter-criteria.ps1`])
  const cases = JSON.parse(ssh(`powershell -NoProfile -ExecutionPolicy Bypass -File C:\\xlcalc\\work\\filter-criteria.ps1 -dir ${WORK}`, 590_000)) as unknown
  const local = mkdtempSync(join(tmpdir(), 'filter-criteria-'))
  execFileSync('scp', ['-r', `${HOST}:C:/xlcalc/work/filter-criteria`, local])
  ssh(`Remove-Item -Recurse -Force ${WORK}, C:\\xlcalc\\work\\filter-criteria.ps1`)
  mkdirSync(OUT, { recursive: true })
  for (const f of readdirSync(join(local, 'filter-criteria'))) copyFileSync(join(local, 'filter-criteria', f), join(OUT, f))
  writeFileSync(join(OUT, 'answers.json'), `${JSON.stringify({ meta: metaOf(e, 'en-US'), cases }, null, 1)}\n`)
})

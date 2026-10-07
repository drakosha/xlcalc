/**
 * A file through Excel: copied into Excel's sandbox, opened, computed fully, saved as xlsx under another name and
 * closed (`excel-save.applescript`). The repair window ("We found a problem with some content") is closed with
 * No and the file reported as rejected; any other window stops the run and is left open.
 */
import { execFileSync, spawn } from 'node:child_process'
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { declineRepair, dialogText, excelDialogs } from '../excel.js'

const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')
const SCRIPT = fileURLToPath(new URL('../excel-save.applescript', import.meta.url))

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

export type Saved = { saved: Uint8Array } | { rejected: string }

export async function saveThroughExcel(data: Uint8Array, label: string): Promise<Saved> {
  const before = dialogText()
  if (before) throw new Error(`${label}: Excel has a window that is not a book: ${before}`)
  const dir = join(SANDBOX, `fxlsx-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    const out = join(dir, 'saved.xlsx')
    writeFileSync(file, data)
    const child = spawn('osascript', [SCRIPT, file, out, 'saved.xlsx'])
    let err = ''
    child.stderr.on('data', (d: Buffer) => (err += d.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const until = Date.now() + 180_000
    for (;;) {
      const code = await Promise.race([exited, sleep(1000).then(() => undefined)])
      if (code !== undefined) {
        const after = dialogText()
        if (after) throw new Error(`${label}: Excel window after saving: ${after}`)
        if (code !== 0) throw new Error(`${label}: osascript ${code}: ${err.trim().split('\n')[0]}`)
        return { saved: new Uint8Array(readFileSync(out)) }
      }
      const windows = excelDialogs()
      if (windows.length) {
        const text = dialogText()
        if (!declineRepair()) {
          child.kill()
          throw new Error(`${label}: Excel shows a window, stopped without clicking: ${text}`)
        }
        child.kill()
        await exited
        await sleep(1500)
        declineRepair()
        // as the harness's other steps do: a book left open after the refusal is closed unsaved
        const open = Number(execFileSync('osascript', ['-e', 'tell application "Microsoft Excel" to count workbooks'], { encoding: 'utf8' }).trim())
        if (open) execFileSync('osascript', ['-e', 'tell application "Microsoft Excel" to close every workbook saving no'])
        return { rejected: text }
      }
      if (Date.now() > until) {
        child.kill()
        throw new Error(`${label}: Excel did not answer in 180 s`)
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

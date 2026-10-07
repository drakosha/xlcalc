/**
 * A file through Excel: copied into Excel's sandbox and handed to an AppleScript (`excel-save.applescript` opens it,
 * computes it fully, saves it as xlsx under another name and closes it; `excel-colors.applescript` reads fill
 * colours). The repair window ("We found a problem with some content") is closed with No and the file reported
 * as rejected; any other window stops the run and is left open.
 */
import { execFileSync, spawn } from 'node:child_process'
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { declineRepair, dialogText, excelDialogs } from '../excel.js'

const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')
const script = (name: string) => fileURLToPath(new URL(`../${name}`, import.meta.url))

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

const osa = (s: string) => execFileSync('osascript', ['-e', s], { encoding: 'utf8' }).trim()

/** The script's output over the file at `book.xlsx` in the sandbox (`args` after its path), or the repair window's text. */
async function watched(name: string, data: Uint8Array, label: string, args: (dir: string) => string[], after: (dir: string, out: string) => Uint8Array | string): Promise<{ out: Uint8Array | string } | { rejected: string }> {
  const before = dialogText()
  if (before) throw new Error(`${label}: Excel has a window that is not a book: ${before}`)
  const dir = join(SANDBOX, `fxlsx-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    writeFileSync(file, data)
    const child = spawn('osascript', [script(name), file, ...args(dir)])
    let out = ''
    let err = ''
    child.stdout.on('data', (d: Buffer) => (out += d.toString('utf8')))
    child.stderr.on('data', (d: Buffer) => (err += d.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const until = Date.now() + 180_000
    for (;;) {
      const code = await Promise.race([exited, sleep(1000).then(() => undefined)])
      if (code !== undefined) {
        const windows = dialogText()
        if (windows) throw new Error(`${label}: Excel window after the script: ${windows}`)
        if (code !== 0) throw new Error(`${label}: osascript ${code}: ${err.trim().split('\n')[0]}`)
        return { out: after(dir, out) }
      }
      if (excelDialogs().length) {
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
        if (Number(osa('tell application "Microsoft Excel" to count workbooks'))) osa('tell application "Microsoft Excel" to close every workbook saving no')
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

export type Saved = { saved: Uint8Array } | { rejected: string }

/** The file as Excel saves it after opening and computing it. */
export async function saveThroughExcel(data: Uint8Array, label: string): Promise<Saved> {
  const r = await watched('excel-save.applescript', data, label, (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx'], (dir) => new Uint8Array(readFileSync(join(dir, 'saved.xlsx'))))
  return 'out' in r ? { saved: r.out as Uint8Array } : r
}

/** The fill colours Excel shows for cells of the file's first sheet, as `#RRGGBB`. */
export async function fillColors(data: Uint8Array, cells: string[]): Promise<string[]> {
  const r = await watched('excel-colors.applescript', data, 'colors', () => [cells.join(',')], (_, out) => out)
  if ('rejected' in r) throw new Error(`colors: Excel refused the book: ${r.rejected}`)
  return (JSON.parse(r.out as string) as number[][]).map((c) => `#${c.map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`)
}


/**
 * A file through Excel: copied into Excel's sandbox and handed to an AppleScript (`excel-save.applescript` opens it,
 * computes it fully, saves it as xlsx under another name and closes it; `excel-colors.applescript` reads fill
 * colours, `excel-formats.applescript` number formats). The repair window ("We found a problem with some content") is closed with No and the file reported
 * as rejected; any other window stops the run and is left open.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { fromRgb } from '../cond/cases.js'
import { runWatched, settleRejected } from '../excel.js'

const script = (name: string) => fileURLToPath(new URL(`../${name}`, import.meta.url))

/** The script's output over the file at `book.xlsx` in the sandbox (`args` after its path), or the repair window's text. */
function watched(name: string, data: Uint8Array, label: string, args: (dir: string) => string[], after: (dir: string, out: string) => Uint8Array | string): Promise<{ out: Uint8Array | string } | { rejected: string }> {
  return runWatched<{ out: Uint8Array | string } | { rejected: string }>({
    script: script(name),
    label,
    folder: 'fxlsx',
    book: (file) => writeFileSync(file, data),
    args,
    timeout: 180_000,
    poll: 1000,
    after: 'script',
    fatal: true,
    // a book left open after the refusal is closed unsaved, as the harness's other steps do
    repair: async (stop, text) => {
      await stop()
      await settleRejected()
      return { rejected: text }
    },
    done: (out, dir) => ({ out: after(dir, out) }),
  })
}

export type Saved = { saved: Uint8Array; shown?: Record<string, string> } | { rejected: string }

/**
 * The file as Excel saves it after opening and computing it, with `enter`'s formulas typed into ranges of its first
 * sheet before; and the formulas Excel shows (`formula2`) in the cells `shown` names of that sheet, by address.
 */
export async function saveThroughExcel(data: Uint8Array, label: string, enter: readonly (readonly [range: string, formula: string])[] = [], shown: readonly string[] = []): Promise<Saved> {
  const args = [enter.map(([r, f]) => `${r}\t${f}`).join('\n'), ...(shown.length ? [shown.join(',')] : [])]
  let text = ''
  const r = await watched('excel-save.applescript', data, label, (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', ...(args.join('') ? args : [])], (dir, out) => ((text = out), new Uint8Array(readFileSync(join(dir, 'saved.xlsx')))))
  if (!('out' in r)) return r
  return { saved: r.out as Uint8Array, ...(shown.length ? { shown: Object.fromEntries(text.trim().split('\n').map((l) => l.split('\t') as [string, string])) } : {}) }
}

/** The fill colours Excel shows for cells of the file's first sheet, as `#RRGGBB`. */
export async function fillColors(data: Uint8Array, cells: string[]): Promise<string[]> {
  const r = await watched('excel-colors.applescript', data, 'colors', () => [cells.join(',')], (_, out) => out)
  if ('rejected' in r) throw new Error(`colors: Excel refused the book: ${r.rejected}`)
  return (JSON.parse(r.out as string) as number[][]).map(fromRgb)
}

/** The fill (colour, pattern as Excel names it) and the font's colour Excel shows for cells of the file's first sheet. */
export async function cellFills(data: Uint8Array, cells: string[]): Promise<{ fill: string; pattern: string; font: string }[]> {
  const r = await watched('excel-fills.applescript', data, 'fills', () => [cells.join(',')], (_, out) => out)
  if ('rejected' in r) throw new Error(`fills: Excel refused the book: ${r.rejected}`)
  return (JSON.parse(r.out as string) as [number[], string, number[]][]).map(([fill, pattern, font]) => ({ fill: fromRgb(fill), pattern, font: fromRgb(font) }))
}

/** The number format codes Excel shows for cells of the file's first sheet. */
export async function numberFormats(data: Uint8Array, cells: string[]): Promise<string[]> {
  const r = await watched('excel-formats.applescript', data, 'formats', () => [cells.join(',')], (_, out) => out)
  if ('rejected' in r) throw new Error(`formats: Excel refused the book: ${r.rejected}`)
  return JSON.parse(r.out as string) as string[]
}

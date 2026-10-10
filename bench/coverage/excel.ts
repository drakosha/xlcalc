/**
 * Excel for Mac as the oracle: a book goes into Excel's sandbox (Excel opens it from there without asking for
 * access), `excel-read.applescript` opens it, recalculates, reads three columns that never hold an error and
 * closes it without saving. The interface language and the calculation locale come from Excel's launch
 * arguments (`launchArgs` in `locales.ts`: `-AppleLanguages (en) -AppleLocale en_US` for en-US, none for the
 * system's ru-RU); Excel's settings are not changed.
 */
import { execFileSync, spawn } from 'node:child_process'
import { copyFileSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { ErrCode } from './data.js'
import type { Val } from './formulajs.js'
import { launchArgs } from './locales.js'

const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')
const SCRIPT = fileURLToPath(new URL('./excel-read.applescript', import.meta.url))

/** ERROR.TYPE → error code. */
const ERROR_TYPE: Record<number, ErrCode> = { 1: 'NULL', 2: 'DIV0', 3: 'VALUE', 4: 'REF', 5: 'NAME', 6: 'NUM', 7: 'NA', 8: 'GETTING_DATA', 9: 'SPILL', 14: 'CALC' }

/** A cell read through its type code (`READ_COLUMNS` in `book.ts`: 0 number, -1 TRUE, -2 FALSE, -3 text, > 0 ERROR.TYPE), number and text. */
export function typedValue(type: number, num: unknown, text: unknown): Val {
  if (type === 0) return { t: 'n', v: Number(num) }
  if (type === -1) return { t: 'b', v: true }
  if (type === -2) return { t: 'b', v: false }
  if (type === -3) return { t: 's', v: String(text) }
  if (type > 0) return { t: 'e', v: ERROR_TYPE[type] ?? 'ERROR' }
  return { t: 'other', v: `type ${type}` }
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

const LOCK = join(homedir(), '.cache/xlcalc-excel.lock')

const alive = (pid: number) => {
  try {
    process.kill(pid, 0)
    return true
  } catch {
    return false
  }
}

/**
 * One Excel per machine, while runs from different worktrees go in parallel: a restart for another locale in
 * the middle of another run breaks both. The lock is a directory (mkdir is atomic) holding the owner's pid; a
 * lock left by a dead process is taken over. Waits up to 60 minutes.
 */
export async function withExcelLock<T>(fn: () => Promise<T>): Promise<T> {
  mkdirSync(join(homedir(), '.cache'), { recursive: true })
  const until = Date.now() + 60 * 60_000
  let said = false
  for (;;) {
    try {
      mkdirSync(LOCK)
      writeFileSync(join(LOCK, 'pid'), String(process.pid))
      break
    } catch {
      let owner = NaN
      try {
        owner = Number(readFileSync(join(LOCK, 'pid'), 'utf8'))
      } catch {}
      if (Number.isFinite(owner) && !alive(owner)) {
        rmSync(LOCK, { recursive: true, force: true })
        continue
      }
      if (Date.now() > until) throw new Error(`Excel has been busy with process ${owner} for over 60 minutes (${LOCK})`)
      if (!said) console.log(`Excel is busy with process ${owner}, waiting`)
      said = true
      await sleep(2000)
    }
  }
  const release = () => rmSync(LOCK, { recursive: true, force: true })
  process.once('exit', release)
  try {
    return await fn()
  } finally {
    release()
    process.removeListener('exit', release)
  }
}

/** Excel's process and its arguments; null when it is not running. */
export function excelProcess(): { pid: number; args: string } | null {
  const ps = execFileSync('ps', ['-axo', 'pid=,args='], { encoding: 'utf8' })
  for (const line of ps.split('\n')) {
    const m = /^\s*(\d+)\s+(\/Applications\/Microsoft Excel\.app\/Contents\/MacOS\/Microsoft Excel.*)$/.exec(line)
    if (m) return { pid: Number(m[1]), args: m[2] }
  }
  return null
}

const osa = (script: string, timeout = 60_000) => execFileSync('osascript', ['-e', script], { encoding: 'utf8', timeout }).trim()

/** AppleEvents go only to a running Excel: `tell` would otherwise start it without the arguments needed. */
function countWorkbooks(): number {
  if (!excelProcess()) throw new Error('Excel is not running')
  return Number(osa('tell application "Microsoft Excel" to count workbooks'))
}

/** Whether the running Excel was started for a locale: its interface language and region arguments are the locale's launch arguments. */
const startedFor = (args: string, tag: string) => (args.match(/-Apple(?:Languages|Locale) \S+/g) ?? []).join(' ') === launchArgs(tag).join(' ')

export function excelVersion(): string {
  return osa('tell application "Microsoft Excel" to get version')
}

/**
 * Excel in the locale wanted: if it runs in another, with no books open it quits and starts again (`open -a`,
 * for en-US with `--args -AppleLanguages (en) -AppleLocale en_US`). Open books: refused.
 */
export async function ensureExcel(tag: string): Promise<void> {
  const p = excelProcess()
  if (p && startedFor(p.args, tag)) return closeStartupBook()
  if (p) {
    // the empty book an Excel started for another locale opened goes; any other open book cancels the restart
    closeStartupBook()
    osa('tell application "Microsoft Excel" to quit')
    for (let i = 0; excelProcess(); i++) {
      if (i > 120) throw new Error('Excel did not quit in 60 s')
      await sleep(500)
    }
  }
  const args = launchArgs(tag)
  execFileSync('open', ['-a', 'Microsoft Excel', ...(args.length ? ['--args', ...args] : [])])
  for (let i = 0; ; i++) {
    if (i > 120) throw new Error('Excel did not start in 60 s')
    await sleep(500)
    const q = excelProcess()
    if (!q) continue
    try {
      countWorkbooks()
      break
    } catch {
      // not answering AppleEvents yet
    }
  }
  const q = excelProcess()!
  if (!startedFor(q.args, tag)) throw new Error(`Excel started in the wrong locale: ${q.args}`)
  closeStartupBook()
}

/**
 * Started without a document, Excel opens an empty new book ("Book1", "Книга1" in Russian, named in the interface's
 * language): a new book with no file and no edits is closed without saving. Any other open book: refused.
 */
function closeStartupBook() {
  if (countWorkbooks() === 1) {
    const [name, saved, path] = osa('tell application "Microsoft Excel" to get {name, saved, path} of workbook 1').split(/,\s*/)
    if (saved === 'true' && !path) osa(`tell application "Microsoft Excel" to close workbook "${name}" saving no`)
  }
  const n = countWorkbooks()
  if (n !== 0) throw new Error(`Excel has ${n} books open (not the empty startup book)`)
}

/**
 * Excel's windows that are not book windows (a modal dialog and the like): subrole and text, read through
 * System Events, nothing is clicked. System Events fails with -1719 when a book's window closes while it
 * lists them: then it asks again, three times at most.
 */
export function excelDialogs(): { subrole: string; text: string }[] {
  for (let i = 0; ; i++)
    try {
      return listDialogs()
    } catch (e) {
      if (i >= 2 || !String((e as { stderr?: string }).stderr).includes('(-1719)')) throw e
    }
}

/** The windows that are not books in one line (`subrole «text»; …`), empty when there are none. */
export const dialogText = (): string => excelDialogs().map((d) => `${d.subrole} «${d.text}»`).join('; ')

function listDialogs(): { subrole: string; text: string }[] {
  const script = `tell application "System Events" to tell process "Microsoft Excel"
  set out to ""
  repeat with w in windows
    set sr to ""
    try
      set sr to subrole of w
    end try
    if sr is not "AXStandardWindow" then
      set tx to ""
      repeat with t in (static texts of w)
        try
          set v to value of t
          if v is not missing value then set tx to tx & (v as text) & " "
        end try
      end repeat
      set out to out & sr & "\t" & tx & linefeed
    end if
  end repeat
  return out
end tell`
  const out = execFileSync('osascript', ['-e', script], { encoding: 'utf8', timeout: 30_000 })
  return out
    .split('\n')
    .filter((l) => l.trim())
    .map((l) => {
      const [subrole, text = ''] = l.split('\t')
      return { subrole, text: text.trim() }
    })
}

/** The read script's output (type codes, numbers, texts) as values by row. */
function parseRead(out: string, n: number): Val[] {
  const [types, nums, strs] = JSON.parse(out) as [number[][], number[][], string[][]]
  if (types.length !== n || nums.length !== n || strs.length !== n) throw new Error(`Excel gave ${types.length}/${nums.length}/${strs.length} rows of ${n}`)
  return types.map(([t], i) => typedValue(t, nums[i][0], strs[i][0]))
}

/** The values of a book's case formulas by row; the book is copied into the sandbox and removed after. A part is small, so the wait is short (3 minutes). */
export function runExcel(book: string, n: number): Val[] {
  if (countWorkbooks() !== 0) throw new Error('Excel has open books')
  const dir = join(SANDBOX, `fcov-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    copyFileSync(book, file)
    return parseRead(execFileSync('osascript', [SCRIPT, file, String(n)], { encoding: 'utf8', timeout: 180_000, maxBuffer: 1 << 28 }), n)
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

/** The window "We found a problem with some content": Excel does not read the whole book, usually over a formula it does not load. */
const REPAIR = /problem with some content|проблем\S* с содержим|ошибка в части содержимого/i

/** Closes the book repair window with No; other windows are left alone. */
export function declineRepair(): boolean {
  if (!excelDialogs().some((d) => REPAIR.test(d.text))) return false
  osa(`tell application "System Events" to tell process "Microsoft Excel"
  repeat with w in windows
    set tx to ""
    repeat with t in (static texts of w)
      try
        set v to value of t
        if v is not missing value then set tx to tx & (v as text)
      end try
    end repeat
    if tx contains "problem with some content" or tx contains "содержим" then
      repeat with b in (buttons of w)
        if name of b is in {"No", "Нет"} then
          click b
          -- the window is gone: the list of windows is stale
          return
        end if
      end repeat
    end if
  end repeat
end tell`)
  return true
}

/** After the repair window was declined and its script stopped: a second one declined, a book Excel left open closed. */
export async function settleRejected(): Promise<void> {
  await sleep(1500)
  declineRepair()
  if (countWorkbooks() !== 0) osa('tell application "Microsoft Excel" to close every workbook saving no')
}

/**
 * As runExcel, but watching for the repair window: it is closed and 'rejected' returned instead of waiting for the
 * AppleEvent timeout. With `formula2`, the case formulas as Excel 365 shows them too (`@` where it intersects).
 */
export async function runExcelWatched(book: string, n: number, formula2 = false): Promise<{ values: Val[]; formulas?: string[] } | 'rejected'> {
  const out = await runScriptWatched(SCRIPT, book, n, formula2 ? ['formula2'] : [])
  if (out === 'rejected') return out
  const values = parseRead(out, n)
  if (!formula2) return { values }
  const f = (JSON.parse(out) as unknown[][][])[3]
  return { values, formulas: f.map(([x]) => String(x)) }
}

/** The AppleScript `script` over a copy of the book in the sandbox (arguments: the path, `n`, `more`): its output, or 'rejected' as in runExcelWatched. */
export async function runScriptWatched(script: string, book: string, n: number, more: string[] = []): Promise<string | 'rejected'> {
  if (countWorkbooks() !== 0) throw new Error('Excel has open books')
  const dir = join(SANDBOX, `fcov-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    copyFileSync(book, file)
    const child = spawn('osascript', [script, file, String(n), ...more])
    let out = ''
    let err = ''
    child.stdout.on('data', (b: Buffer) => (out += b.toString('utf8')))
    child.stderr.on('data', (b: Buffer) => (err += b.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const until = Date.now() + 180_000
    for (;;) {
      const code = await Promise.race([exited, sleep(2000).then(() => undefined)])
      if (code !== undefined) {
        if (code !== 0) throw new Error(`osascript ${code}: ${err.trim().split('\n')[0]}`)
        return out
      }
      if (declineRepair()) {
        child.kill()
        await exited
        await settleRejected()
        return 'rejected'
      }
      if (Date.now() > until) {
        child.kill()
        await exited
        throw new Error('Excel did not answer in 180 s')
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

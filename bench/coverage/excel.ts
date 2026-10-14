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
import { alive, sleep } from './proc.js'
import type { BookInput } from '../../src/contract.js'
import { writeXlsx } from '../../src/xlsx.js'

/** Excel's sandbox: Excel opens a book from there without asking for access. */
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

const LOCK = join(homedir(), '.cache/xlcalc-excel.lock')

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

/** The repair window declined where System Events does not list it, on another Space: Excel activated first, which brings it forward. */
export async function declineElsewhere(): Promise<boolean> {
  if (declineRepair()) return true
  osa('tell application "Microsoft Excel" to activate')
  await sleep(2000)
  return declineRepair()
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

/**
 * Runs `items` in one book; a book Excel does not load is halved, and so on down to the items it refuses one by one
 * (`refused`). `run` answers whether Excel loaded the book; `label` names the halves (`3a`, `3ab`).
 */
export async function halving<T>(items: T[], label: string, run: (part: T[], label: string) => Promise<boolean>, refused: (item: T) => void): Promise<void> {
  if (await run(items, label)) return
  if (items.length === 1) return refused(items[0])
  const mid = items.length >> 1
  await halving(items.slice(0, mid), `${label}a`, run, refused)
  await halving(items.slice(mid), `${label}b`, run, refused)
}

/** How `runWatched` runs an AppleScript over a book in Excel's sandbox and what it does with the windows Excel shows meanwhile. */
export interface Watch<T> {
  script: string
  /** Begins the errors: a case's id, a part's label; none for a bare message. */
  label?: string
  /** The book's folder in the sandbox, before the pid: each step has its own (`fsave`). */
  folder: string
  /** Writes the book at the path given. */
  book: (file: string) => void | Promise<void>
  /** The script's arguments after the book's path; `dir` is the folder, for a file the script saves. */
  args: (dir: string) => string[]
  /** No answer in `timeout` ms stops the run; Excel's windows are looked at every `poll` ms. */
  timeout: number
  poll: number
  /** Before the run: no window that is not a book (unless false), and with `empty` no open book. */
  before?: false
  empty?: true
  /** After the script ends a window stops the run, named by what ran (`case`, `part`, `script`); unchecked when left out. */
  after?: string
  /** A window that is not the repair window stops the run (false: the script answers the windows it brings itself). */
  fatal: boolean
  /** The repair window, declined: the step's answer, or undefined to go on; `stop` kills the script and waits for it, `text` is what Excel's windows said. */
  repair: (stop: () => Promise<void>, text: string) => Promise<T | undefined>
  /** The answer from the script's output, taken before the folder goes. */
  done: (out: string, dir: string) => T | Promise<T>
  /** The script failed (`err` its error output): the step's answer, or undefined for the error. */
  failed?: (err: string, stop: () => Promise<void>) => Promise<T | undefined>
  /** No answer in time, the script killed: the step's answer, or undefined for the error. */
  late?: (stop: () => Promise<void>) => Promise<T | undefined>
}

/**
 * An AppleScript over a book in Excel's sandbox (the folder is removed after), watched while it runs: the repair window
 * ("We found a problem with some content") is declined and the step's `repair` answers; another window that is not a
 * book stops the run when `fatal`, nothing clicked; so does Excel not answering in time.
 */
/** The keyboard's script, which the steps that type into Excel run. */
export const KEYS = fileURLToPath(new URL('./excel-keys.applescript', import.meta.url))

/**
 * What `runWatched` takes for a part of a set run with the keyboard's help (`fit`, `rows`): the book written whole as
 * xlsx (`label` names it in errors), the saved file and the part's spec as arguments; typed on the keyboard, the script
 * answers a refused formula's window itself, and a book Excel would repair stops the run.
 */
export function keyedPart(label: string, script: string, folder: string, book: BookInput, spec: unknown, keys: boolean): Omit<Watch<never>, 'done'> {
  return {
    script,
    label,
    folder,
    book: async (file) => {
      const { data, issues } = await writeXlsx({ book }, { values: 'omit' })
      if (issues.length) throw new Error(`${label}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
      writeFileSync(file, data)
    },
    args: (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', JSON.stringify(spec), KEYS],
    timeout: 290_000,
    poll: 1000,
    after: 'part',
    fatal: !keys,
    repair: async (stop) => {
      await stop()
      await settleRejected()
      throw new Error(`${label}: Excel would repair the book`)
    },
  }
}

export async function runWatched<T>(w: Watch<T>): Promise<T> {
  const at = (m: string) => (w.label === undefined ? m : `${w.label}: ${m}`)
  if (w.before !== false) {
    const before = dialogText()
    if (before) throw new Error(at(`Excel has a window that is not a book: ${before}`))
  }
  if (w.empty && countWorkbooks() !== 0) throw new Error('Excel has open books')
  const dir = join(SANDBOX, `${w.folder}-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    await w.book(file)
    const child = spawn('osascript', [w.script, file, ...w.args(dir)])
    let out = ''
    let err = ''
    child.stdout.on('data', (b: Buffer) => (out += b.toString('utf8')))
    child.stderr.on('data', (b: Buffer) => (err += b.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const stop = async () => {
      child.kill()
      await exited
    }
    const until = Date.now() + w.timeout
    for (;;) {
      const code = await Promise.race([exited, sleep(w.poll).then(() => undefined)])
      if (code !== undefined) {
        const after = w.after === undefined ? '' : dialogText()
        if (after) throw new Error(at(`Excel window after the ${w.after}: ${after}`))
        if (code !== 0) {
          const r = await w.failed?.(err, stop)
          if (r !== undefined) return r
          throw new Error(at(`osascript ${code}: ${err.trim().split('\n')[0]}`))
        }
        return await w.done(out, dir)
      }
      const open = dialogText()
      if (open && declineRepair()) {
        const r = await w.repair(stop, open)
        if (r !== undefined) return r
      } else if (open && w.fatal) {
        child.kill()
        throw new Error(at(`Excel shows a window, stopped without clicking: ${open}`))
      }
      if (Date.now() > until) {
        child.kill()
        const r = await w.late?.(stop)
        if (r !== undefined) return r
        throw new Error(at(`Excel did not answer in ${w.timeout / 1000} s`))
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

/** The AppleScript `script` over a copy of the book in the sandbox (arguments: the path, `n`, `more`): its output, or 'rejected' as in runExcelWatched. */
export async function runScriptWatched(script: string, book: string, n: number, more: string[] = []): Promise<string | 'rejected'> {
  return runWatched<string>({
    script,
    folder: 'fcov',
    before: false,
    empty: true,
    book: (file) => copyFileSync(book, file),
    args: () => [String(n), ...more],
    timeout: 180_000,
    poll: 2000,
    fatal: false,
    repair: async (stop) => {
      await stop()
      await settleRejected()
      return 'rejected'
    },
    done: (out) => out,
    late: async (stop) => void (await stop()),
  })
}

/**
 * One case of a step that saves the book: `data` goes into the sandbox, the AppleScript `script` takes its path, the path
 * to save to, the saved file's name and `args`, and answers its output and the saved file, or the repair window's text
 * (declined). A window that is not a book stops the run, as does no answer in `timeoutMs`.
 */
export async function runSavingScript(id: string, script: string, data: Uint8Array, args: string[], timeoutMs = 180_000): Promise<{ out: string; saved: Uint8Array } | { rejected: string }> {
  return runWatched<{ out: string; saved: Uint8Array } | { rejected: string }>({
    script,
    label: id,
    folder: 'fsave',
    book: (file) => writeFileSync(file, data),
    args: (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', ...args],
    timeout: timeoutMs,
    poll: 1000,
    after: 'case',
    fatal: true,
    // the window declined once more, no book closed
    repair: async (stop) => {
      await stop()
      await sleep(1500)
      declineRepair()
      return { rejected: 'Excel did not accept the book (repair window declined)' }
    },
    // the book did not open (`wb` undefined): the repair window on another Space, which the watch does not see
    failed: async (err, stop) => {
      if (!/\bwb\b/.test(err) || !(await declineElsewhere())) return undefined
      await stop()
      await settleRejected()
      return { rejected: 'Excel did not accept the book (repair window declined)' }
    },
    done: (out, dir) => ({ out, saved: new Uint8Array(readFileSync(join(dir, 'saved.xlsx'))) }),
  })
}

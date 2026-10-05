/**
 * Сессия над книгой: правки атомарно одним пересчётом зависимых. Живая сессия (`live`) берёт часы и
 * новое зерно на каждый пересчёт; текущие часы (`clock`) пишутся в версию, и полный расчёт с ними
 * даёт тот же хэш.
 */
import type { BookInput, CellChange, CellEdit, Clock, Explanation, ExplainOptions, NameInput, Session, SessionOptions } from './contract.js'
import { applyClock, applyEdits, createState, explainCell, explainText, toCellValue, valueAt, type EngineState } from './engine.js'
import { formatAddress, parseAddress } from './model.js'
import { buildResult, clockSerial, outValue } from './result.js'
import { sheetKey } from './syntax.js'

function liveClock(): Clock {
  const seed = new Uint32Array(1)
  crypto.getRandomValues(seed)
  return { now: new Date().toISOString(), seed: seed[0] }
}

export function openSession(book: BookInput, opts: SessionOptions): Session {
  const live = 'live' in opts
  let clock: Clock = live ? liveClock() : { ...opts.clock }
  const externals = opts.externals
  let state: EngineState | null = createState(book, { now: clockSerial(clock), seed: clock.seed, externals })
  const st = () => {
    if (!state) throw new Error('session is closed')
    return state
  }
  const changes = (s: EngineState, keys: string[]): CellChange[] =>
    keys.map((k) => {
      const [i, r, c] = k.split(':').map(Number)
      return { sheet: s.sheetNames[i], cell: formatAddress(r, c), value: outValue(toCellValue(valueAt(s, i, r, c))) }
    })
  const tick = (s: EngineState) => {
    clock = liveClock()
    s.ctx.now = clockSerial(clock)
    s.ctx.seed = clock.seed
  }
  return {
    get clock() {
      return { ...clock }
    },
    setCells(edits: CellEdit[]) {
      const s = st()
      if (live) tick(s)
      return changes(s, applyEdits(s, edits, live))
    },
    setNames(names: NameInput[]) {
      const old = st()
      if (live) tick(old)
      const next = createState({ sheets: old.input.sheets, names }, { now: clockSerial(clock), seed: clock.seed, externals })
      // имена меняют граф целиком: книга собирается заново, изменения по всем ячейкам
      const keys: string[] = []
      next.input.sheets.forEach((_, i) => {
        const cols = Math.max(old.ctx.sheets[i].cols.length, next.ctx.sheets[i].cols.length)
        const rows = Math.max(old.ctx.sheets[i].rows, next.ctx.sheets[i].rows)
        for (let c = 0; c < cols; c++)
          for (let r = 0; r < rows; r++) {
            const a = toCellValue(valueAt(old, i, r, c))
            const b = toCellValue(valueAt(next, i, r, c))
            if (JSON.stringify(a) !== JSON.stringify(b)) keys.push(`${i}:${r}:${c}`)
          }
      })
      state = next
      return changes(next, keys)
    },
    setClock(next: Clock) {
      if (live) throw new Error('a live session takes its clock from the system')
      clock = { ...next }
      const s = st()
      return changes(s, applyClock(s, clockSerial(clock), clock.seed))
    },
    explain(sheet: string, cell: string, o: ExplainOptions = {}): Explanation {
      const s = st()
      const i = s.sheetIdx.get(sheetKey(sheet))
      if (i === undefined) throw new Error(`unknown sheet ${sheet}`)
      const { row, col } = parseAddress(cell)
      const e = explainCell(s, i, row, col, o.depth ?? 3, o.maxNodes ?? 200, o.expand ?? false)
      return { ...e, text: explainText(e.root) }
    },
    result() {
      const t0 = performance.now()
      return buildResult(st(), clock, t0)
    },
    close() {
      state = null
    },
  }
}

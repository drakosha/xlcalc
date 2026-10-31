/**
 * The answers to the calls a workbook makes outside itself (`Externals`): the functions of connections and the host's
 * own functions. A call is keyed by its text with its arguments' values (`BALANCE("2025-01")`, as `Externals.values`
 * keys it), taken from `values` or asked of `resolve` once and kept: by a calculation, or by a session until
 * `refreshConnections`. An answer on its way leaves its cell waiting (`#BUSY!`, a cube function's `#GETTING_DATA`); the
 * cells that waited are computed again once it arrives (`settled`). A cube function's member is kept by the cell that
 * answered it, and given to a cube function that reads that cell.
 */
import type { CellValue, ExternalAnswer, ExternalResolver } from './contract.js'
import type { Arg, Cx } from './args.js'
import { Arr } from './arrays/value.js'
import { callSignal } from './exec/budget.js'
import { formatAddress, placeOf } from './model.js'
import { printFormula, type Ast } from './syntax.js'
import { Areas, E, fromCellValue, isErr, Rng, type V } from './values.js'

/** An answer kept: a value or an array, a cube member with its caption, none (the function's own), or one on its way. */
type Answer = { value: V | Arr | undefined; member?: string } | { pending: Promise<void> }

/** An argument's value for the host: a range or an array as rows, a cube member as its name. */
type Given = CellValue | CellValue[][] | undefined

const cellValue = (v: V): CellValue => (isErr(v) ? { error: v.code } : v)

/** A value waiting for an answer from outside the workbook. */
const waits = (v: Given): boolean => v !== null && typeof v === 'object' && !Array.isArray(v) && (v.error === 'BUSY' || v.error === 'GETTING_DATA')

/** A value as a literal of a call's text. */
function literal(v: CellValue): Ast {
  if (v === null) return { k: 'missing' }
  if (typeof v === 'number') return { k: 'num', v }
  if (typeof v === 'string') return { k: 'str', v }
  if (typeof v === 'boolean') return { k: 'bool', v }
  return { k: 'err', code: v.error }
}

/** A call's text with its arguments' values, as `printFormula` writes a call of literals (`canonical`). */
export function callKey(name: string, args: readonly Given[]): string {
  const ast: Ast = {
    k: 'fn',
    name,
    args: args.map((a): Ast => (a === undefined ? { k: 'missing' } : Array.isArray(a) ? { k: 'arr', rows: a.map((row) => row.map((v) => literal(v) as Extract<Ast, { k: 'num' | 'str' | 'bool' | 'err' }>)) } : literal(a))),
  }
  return printFormula(ast)
}

/** A host's answer as the engine's value; an array of rows as an array, an empty one none. */
function taken(a: ExternalAnswer): { value: V | Arr | undefined; member?: string } {
  if (a === undefined) return { value: undefined }
  if (Array.isArray(a)) {
    const h = a.length
    const w = h ? a[0].length : 0
    if (!h || !w || a.some((row) => !Array.isArray(row) || row.length !== w)) return { value: E.VALUE }
    return { value: new Arr(h, w, h, w, a.flat().map(fromCellValue)) }
  }
  if (a !== null && typeof a === 'object' && 'member' in a) return { value: fromCellValue(a.value), member: String(a.member) }
  return { value: fromCellValue(a as CellValue) }
}

const isThenable = (x: unknown): x is PromiseLike<unknown> => x !== null && typeof x === 'object' && typeof (x as { then?: unknown }).then === 'function'

/** `m`'s entry at `p` set to `v` (none: deleted), the entry as the call started recorded in `was` at its first change. */
function change<V>(m: Map<number, V>, was: Map<number, V | undefined>, p: number, v: V | undefined) {
  if (!was.has(p)) was.set(p, m.get(p))
  if (v === undefined) m.delete(p)
  else m.set(p, v)
}

/** `m`'s entries changed since the call started put back (`change`). */
function putBack<V>(m: Map<number, V>, was: Map<number, V | undefined>) {
  for (const [p, v] of was) {
    if (v === undefined) m.delete(p)
    else m.set(p, v)
  }
  was.clear()
}

export class Outside {
  private readonly answers = new Map<string, Answer>()
  /** The member a cube function answered, by its cell (`placeOf`). */
  private readonly members = new Map<number, string>()
  /** The cells waiting for an answer on its way, by place, with the call's text. */
  readonly waiting = new Map<number, string>()
  // what the call running changed (`mark`), put back if it stops (`back`): the waits and members as it started, the answers it dropped
  private readonly wasWaiting = new Map<number, string | undefined>()
  private readonly wasMembers = new Map<number, string | undefined>()
  private readonly dropped = new Map<string, Answer>()
  /** The calls the cells waited for as a stopped call or a preview started: the cells computed show them waiting, though their answers arrived. */
  private held: ReadonlySet<string> | null = null
  constructor(
    readonly resolve: ExternalResolver | undefined,
    readonly values: Readonly<Record<string, CellValue>>,
  ) {}

  /** A call starts: what it changes is recorded from here, and the answers held for a state built again are taken as any. */
  mark() {
    this.wasWaiting.clear()
    this.wasMembers.clear()
    this.dropped.clear()
    this.held = null
  }

  /** A preview starts: the answers the cells wait for are held, so that it takes none of them (`back` after it). */
  hold() {
    this.held = new Set(this.waiting.values())
  }

  /**
   * The call stopped, or a preview ended: the waits, members and answers as the call started put back, and the answers
   * the cells waited for then held, so that the state built again shows them waiting and the next call takes them with
   * all they change; with `hold` false, as a preview inside a call ends, the call goes on taking them.
   */
  back(hold = true) {
    putBack(this.waiting, this.wasWaiting)
    putBack(this.members, this.wasMembers)
    for (const [key, a] of this.dropped) this.answers.set(key, a)
    this.dropped.clear()
    this.held = hold ? new Set(this.waiting.values()) : null
  }

  /** A cell computed again forgets what it waited for and the member it answered. */
  forget(s: number, r: number, c: number) {
    const p = placeOf(s, r, c)
    if (this.waiting.has(p)) change(this.waiting, this.wasWaiting, p, undefined)
    if (this.members.has(p)) change(this.members, this.wasMembers, p, undefined)
  }

  /** Arguments as the host takes them: a cell a cube function answered with a member gives that member. */
  private given(cx: Cx, args: readonly (Arg | undefined)[]): Given[] {
    return args.map((x): Given => {
      if (x === undefined) return undefined
      if (x instanceof Rng && x.r1 === x.r2 && x.c1 === x.c2) {
        const m = this.members.get(placeOf(x.s, x.r1, x.c1))
        if (m !== undefined) return m
      }
      if (x instanceof Areas) return { error: 'VALUE' }
      if (x instanceof Rng || x instanceof Arr) {
        const g = cx.grid(x)
        if (g.rows === 1 && g.cols === 1) return cellValue(g.v[0] ?? null)
        return Array.from({ length: g.h }, (_, i) => g.v.slice(i * g.w, (i + 1) * g.w).map((v) => cellValue(v ?? null)))
      }
      return cellValue(x as V)
    })
  }

  /**
   * The answer to `name`'s call from the cell at `s`, `r`, `c`: from `values`, kept, or asked of `resolve`; undefined
   * where there is none (the function's own answer), a waiting error where it is on its way.
   */
  ask(cx: Cx, name: string, args: readonly (Arg | undefined)[], sheet: string, s: number, r: number, c: number, cube: boolean): V | Arr | undefined {
    const values = this.given(cx, args)
    // an argument still waiting for its own answer: the call waits with it, computed again once that one arrives
    if (values.some((v) => (Array.isArray(v) ? v.some((row) => row.some(waits)) : waits(v)))) return cube ? E.GETTING_DATA : E.BUSY
    const key = callKey(name, values)
    if (Object.hasOwn(this.values, key)) return fromCellValue(this.values[key])
    if (!this.resolve) return undefined
    let a = this.answers.get(key)
    if (!a) {
      a = this.request(name, key, values, sheet, r, c)
      this.answers.set(key, a)
    }
    const p = placeOf(s, r, c)
    if ('pending' in a || this.held?.has(key)) {
      change(this.waiting, this.wasWaiting, p, key)
      return cube ? E.GETTING_DATA : E.BUSY
    }
    if (a.member !== undefined) change(this.members, this.wasMembers, p, a.member)
    return a.value
  }

  private request(name: string, key: string, args: Given[], sheet: string, r: number, c: number): Answer {
    const signal = callSignal()
    try {
      const out = this.resolve!({ name, key, args, sheet, cell: formatAddress(r, c), signal })
      if (!isThenable(out)) return taken(out as ExternalAnswer)
      // an answer lands only while its call is still the one kept: refreshConnections may have asked again since;
      // a call aborted while its answer was on its way asks again next time
      const asked: Answer = { pending: undefined as unknown as Promise<void> }
      const kept = () => this.answers.get(key) === asked
      asked.pending = Promise.resolve(out).then(
        (x) => void (kept() && this.answers.set(key, taken(x as ExternalAnswer))),
        () => void (kept() && (signal.aborted ? this.answers.delete(key) : this.answers.set(key, { value: undefined }))),
      )
      return asked
    } catch {
      return { value: undefined }
    }
  }

  /** Whether a waiting cell's answer is still on its way. */
  onTheWay(): boolean {
    for (const key of this.waiting.values()) {
      if (this.held?.has(key)) continue
      const a = this.answers.get(key)
      if (a && 'pending' in a) return true
    }
    return false
  }

  /** The places of the waiting cells whose answers arrived (or were dropped, to be asked again). */
  arrived(): number[] {
    const out: number[] = []
    for (const [p, key] of this.waiting) {
      if (this.held?.has(key)) continue
      const a = this.answers.get(key)
      if (!a || !('pending' in a)) out.push(p)
    }
    return out
  }

  /** The answers on their way now, settled. */
  async settled(): Promise<void> {
    await Promise.all([...this.answers.values()].flatMap((a) => ('pending' in a ? [a.pending] : [])))
  }

  /** The kept answers of `names` (every one without them) dropped, so the next calculation asks again. */
  drop(names?: readonly string[]) {
    for (const [key, a] of [...this.answers]) {
      if (names && !names.some((n) => key.startsWith(`${n}(`))) continue
      if (!this.dropped.has(key)) this.dropped.set(key, a)
      this.answers.delete(key)
    }
  }
}

/**
 * The places of the cells written on each sheet of a session's book, as `putInput` (`looks/sheet.ts`), the one way a
 * cell of a live book is written, notes them: what an index of a sheet's cells kept beside the book (`session/search.ts`)
 * reads to stay whole. Indexes read the log from where they last stopped; a log grown past `MAX` starts again in a new
 * epoch, and an index of an older one is built again. No imports: the looks write it, the session reads it.
 */

/** The places `r * MAX_COLS + c` written on a sheet since its log's epoch began. */
interface Log {
  epoch: number
  at: number[]
}

const LOGS = new WeakMap<object, Log>()

/** The places a log holds at most before it starts again. */
const MAX = 65_536

let epochs = 0

/** Notes the write of the cell at `place` on `sheet`. */
export function wrote(sheet: object, place: number): void {
  let log = LOGS.get(sheet)
  if (!log || log.at.length >= MAX) LOGS.set(sheet, (log = { epoch: ++epochs, at: [] }))
  log.at.push(place)
}

/** Where the log of `sheet` stands, for a reader that has read it up to here: its epoch and length (a log begun for it). */
export function writesMark(sheet: object): [epoch: number, read: number] {
  let log = LOGS.get(sheet)
  if (!log) LOGS.set(sheet, (log = { epoch: ++epochs, at: [] }))
  return [log.epoch, log.at.length]
}

/** The places written on `sheet` since `mark`, or null when its log started again since (the reader builds its index anew). */
export function writesSince(sheet: object, mark: readonly [number, number]): number[] | null {
  const log = LOGS.get(sheet)
  return log?.epoch === mark[0] ? log.at.slice(mark[1]) : null
}

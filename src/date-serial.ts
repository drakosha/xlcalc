/**
 * A date as its serial number in a book's date system and back, without a book (`dateToSerial`, `serialToDate`). The
 * text is what `CellInput.date` takes (`2026-01-31`, `2026-01-31T18:00`, with seconds and their fraction), read by the
 * same parse, with no time zone; Excel's phantom 1900-02-29 is serial 60 in the 1900 system. Excel has no date before
 * the system's day 0 or after 9999-12-31 (a cell shows `#` for one, `DATE` gives `#NUM!`), so neither helper makes one:
 * it is an `InputError`.
 */
import { EPOCH_1904, epochOf, MAX_SERIAL, ymd } from './dates.js'
import { InputError, what } from './input-error.js'
import { checkOneOf, checkString, object, onlyKeys } from './input/basic.js'
import { DAY_MS, dateToSerial as serialOfText } from './model.js'

/** The workbook's date system, as in `BookInput.dateSystem`: 1900 without one. */
export interface DateSystemOptions {
  dateSystem?: 1900 | 1904
}


/** The day 0 of the options' date system in the 1900 system. */
function epoch(opts: unknown): number {
  if (opts === undefined) return 0
  const o = object(opts, 'opts', 'options such as { dateSystem: 1904 }')
  onlyKeys(o, 'opts', ['dateSystem'])
  if (o.dateSystem !== undefined) checkOneOf(o.dateSystem, 'opts.dateSystem', [1900, 1904])
  return epochOf(o.dateSystem as 1900 | 1904 | undefined)
}

/** The last serial number of the date system whose day 0 is `e`: the end of 9999-12-31. */
const last = (e: number) => MAX_SERIAL - e + 1

/** Whether `serial` is a date of the date system whose day 0 is `e` (`epochOf`), from day 0 to 9999-12-31. */
export const isDateSerial = (serial: number, e: number) => serial >= 0 && serial < last(e)

const pad = (n: number, width = 2) => String(n).padStart(width, '0')

/**
 * The serial number of a date and time in the form `CellInput.date` accepts (`2026-01-31` is 46053, `2026-01-31T18:00`
 * is 46053.75), in `opts.dateSystem` (1900 without one; 1904 counts from 1904-01-01).
 */
export function dateToSerial(text: string, opts?: DateSystemOptions): number {
  checkString(text, 'text')
  const e = epoch(opts)
  let serial: number
  try {
    serial = serialOfText(text) - e
  } catch {
    throw new InputError('text', `invalid date ${JSON.stringify(text)}: expected an ISO date such as "2025-01-31" or "2025-01-31T18:00:00"`)
  }
  if (serial < 0) throw new InputError('text', `${JSON.stringify(text)} is before the ${e === EPOCH_1904 ? '1904' : '1900'} date system's day 0 (${e === EPOCH_1904 ? '1904-01-01' : '1900-01-00'}): Excel has no such date`)
  return serial
}

/**
 * The date and time of a serial number in `opts.dateSystem`, in the form `CellInput.date` accepts, without a time zone:
 * just the date for a whole number (`2026-01-31`), otherwise with the time to the millisecond (`2026-01-31T18:00`,
 * `2026-01-31T18:00:30.5`). Serial 0 of the 1900 system is `1900-01-00` and 60 is `1900-02-29`, as Excel displays them.
 */
export function serialToDate(serial: number, opts?: DateSystemOptions): string {
  const e = epoch(opts)
  if (typeof serial !== 'number' || !isDateSerial(serial, e))
    throw new InputError('serial', `expected the serial number of a date, from 0 to ${last(e) - 1} (9999-12-31) in the ${e === EPOCH_1904 ? '1904' : '1900'} date system, got ${what(serial)}`)
  let day = Math.floor(serial)
  let ms = Math.round((serial - day) * DAY_MS)
  if (ms === DAY_MS) [day, ms] = [day + 1, 0]
  // the last millisecond of 9999-12-31 rounds to no later day
  if (day === last(e)) [day, ms] = [day - 1, DAY_MS - 1]
  const { y, m, d } = ymd(day + e)
  const date = `${pad(y, 4)}-${pad(m)}-${pad(d)}`
  if (!ms) return date
  const [h, min, s, frac] = [Math.floor(ms / 3_600_000), Math.floor(ms / 60_000) % 60, Math.floor(ms / 1000) % 60, ms % 1000]
  const seconds = s || frac ? `:${pad(s)}${frac ? `.${pad(frac, 3).replace(/0+$/, '')}` : ''}` : ''
  return `${date}T${pad(h)}:${pad(min)}${seconds}`
}

/** Date and time functions, this engine's own; the implementations are in `./date/`. */
import type { FnDef } from '../args.js'
import { functions as calendar } from './date/calendar.js'
import { functions as daycount } from './date/daycount.js'
import { functions as time } from './date/time.js'
import { functions as workday } from './date/workday.js'

export const functions: Record<string, FnDef> = { ...calendar, ...daycount, ...time, ...workday }

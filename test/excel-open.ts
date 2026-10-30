/**
 * Open cases of the random check (the `fuzz` set) and of case sets: differences from Excel whose cause lies in engine files
 * outside the case work (the compiler, the runtime, the array layer), each with one line why and the engine's
 * answer pinned, as KNOWN does. The lead closes them; an entry whose case matches Excel fails the tests.
 */
import type { Val } from '../bench/coverage/formulajs.js'

/** SUMIF inside its own resized sum range: Excel finds the circular reference and leaves the formulas reading the cell uncomputed (empty), so its value is not read; the engine holds the cycle at 0. */
const SUMIF_OWN = "a circular reference through SUMIF's resized sum range: Excel leaves the formulas that read the cell uncomputed, its value unread; the engine holds the cycle at 0"

/** Excel for Mac 16.114 answered #NUM! in the first run of the set `indirect`, a run beside other workbooks' calculations; the same formulas in the book of `indirect-held` gave the engine's answer. */
const MAC_NUM = 'Excel for Mac 16.114 answered #NUM! in the first run of indirect, not in the book of indirect-held (the same formulas, again#1 to again#6): not repeated'

/** The harness wrote a spill reference's file form inside INDIRECT's quotes (`bench/coverage/book.ts`, fixed): the set indirect-held-3 sent it as typed, and Excel answers as the engine. */
const SENT = 'the harness sent INDIRECT("_xlfn.ANCHORARRAY(R!E1)") for INDIRECT("R!E1#"), which Excel answers #REF!; sent as typed (indirect-held-3), Excel answers as the engine'

/** FORECAST.ETS and its kin where they detect the season (seasonality 1, their default): Excel's detection is not reproduced yet, and the engine answers #N/A there. */
const ETS_SEASON = "FORECAST.ETS's detection of the season is not reproduced yet: the engine answers #N/A where Excel detects one"

/** Seasons of 7 to 9 points over 10: Excel takes none, the engine one; the longest season Excel takes is not settled yet. */
const ETS_LONGEST = 'a season of 7 to 9 points over 10 is none in Excel, a season here: the longest season Excel takes is not settled yet'

const NA: Val = { t: 'e', v: 'NA' }

/** The cases of each ETS set where the season is detected, ids separated by spaces. */
const ETS_DETECTED: Readonly<Record<string, string>> = {
  'ets-fit': 
    'holt-trend#17 holt-trend#18 holt-trend#19 holt-trend#20 holt-trend#21 holt-trend#22 holt-level#17 holt-level#18 holt-level#19 holt-level#20 holt-level#21 holt-level#22 holt-line#17 holt-line#18 holt-line#19 holt-line#20 holt-line#21 holt-line#22 holt-flat#17 holt-flat#18 holt-flat#19 holt-flat#20 holt-flat#21 holt-flat#22 holt-walk#17 holt-walk#18 holt-walk#19 holt-walk#20 holt-walk#21 holt-walk#22 holt-down#17 holt-down#18 holt-down#19 holt-down#20 holt-down#21 holt-down#22 holt-three#17 holt-three#18 holt-three#19 holt-three#20 holt-three#21 holt-three#22 holt-four#17 holt-four#18 holt-four#19 holt-four#20 holt-four#21 holt-four#22 aaa-q16#17 aaa-q16#18 aaa-q16#19 aaa-q16#20 aaa-q16#21 aaa-q16#22 aaa-q24#17 aaa-q24#18 aaa-q24#19 aaa-q24#20 aaa-q24#21 aaa-q24#22 aaa-m36#25 aaa-m36#26 aaa-m36#27 aaa-m36#28 aaa-m36#29 aaa-m36#30 aaa-q8#17 aaa-q8#18 aaa-q8#19 aaa-q8#20 aaa-q8#21 aaa-q8#22 aaa-q10#17 aaa-q10#18 aaa-q10#19 aaa-q10#20 aaa-q10#21 aaa-q10#22 aaa-t9#16 aaa-t9#17 aaa-t9#18 aaa-t9#19 aaa-t9#20 aaa-t9#21 aaa-w28#20 aaa-w28#21 aaa-w28#22 aaa-w28#23 aaa-w28#24 aaa-w28#25 aaa-s5#18 aaa-s5#19 aaa-s5#20 aaa-s5#21 aaa-s5#22 aaa-s5#23 detect-p2#1 detect-p2#2 detect-p3#1 detect-p3#2 detect-p4#1 detect-p4#2 detect-p5#1 detect-p5#2 detect-p6#1 detect-p6#2 detect-p7#1 detect-p7#2 detect-p8#1 detect-p8#2 detect-p9#1 detect-p9#2 detect-p10#1 detect-p10#2 detect-p11#1 detect-p11#2 detect-p12#1 detect-p12#2 detect-p12sine#1 detect-p12sine#2 detect-p7long#1 detect-p7long#2 detect-noise#1 detect-noise#2 detect-p4flat#1 detect-p4flat#2 order#2 dup#15 season#6 season#13 omit#1 omit#3 omit#5 omit#6 size#11 size#12',
  'ets-fit-2': 'detect-2-0#1 detect-2-0#2 detect-2-0#3 detect-2-1#1 detect-2-1#2 detect-2-1#3 detect-3-0#1 detect-3-0#2 detect-3-0#3 detect-3-1#1 detect-3-1#2 detect-3-1#3 detect-4-0#1 detect-4-0#2 detect-4-0#3 detect-4-1#1 detect-4-1#2 detect-4-1#3 detect-5-0#1 detect-5-0#2 detect-5-0#3 detect-5-1#1 detect-5-1#2 detect-5-1#3 detect-6-0#1 detect-6-0#2 detect-6-0#3 detect-6-1#1 detect-6-1#2 detect-6-1#3 detect-7-0#1 detect-7-0#2 detect-7-0#3 detect-7-1#1 detect-7-1#2 detect-7-1#3 detect-8-0#1 detect-8-0#2 detect-8-0#3 detect-8-1#1 detect-8-1#2 detect-8-1#3 detect-10-0#1 detect-10-0#2 detect-10-0#3 detect-10-1#1 detect-10-1#2 detect-10-1#3 detect-12-0#1 detect-12-0#2 detect-12-0#3 detect-12-1#1 detect-12-1#2 detect-12-1#3 detect-noise-0#1 detect-noise-1#1 detect-noise-2#1 detect-noise-3#1 detect-noise-4#1 detect-noise-5#1',
}

/** The open ETS cases of a set (`<set>:<id>`): the detected seasons and, in ets-fit-2, the longest seasons. */
function etsOpen(set: string): Record<string, { why: string; engine: Val }> {
  const open: Record<string, { why: string; engine: Val }> = Object.fromEntries(ETS_DETECTED[set].split(' ').map((id) => [`${set}:${id}`, { why: ETS_SEASON, engine: NA }]))
  if (set === 'ets-fit-2')
    for (const [id, v] of [['longest#1', 0.499], ['longest#2', 0.499], ['longest#3', 0.499], ['longest#11', 24.3344715333105], ['longest#12', 19.8838732168262]] as const)
      open[`${set}:${id}`] = { why: ETS_LONGEST, engine: { t: 'n', v } }
  return open
}

/** Open cases of Excel for Windows's sets, by `<locale>:<set>:<id>` as WIN_KNOWN keys them. */
export const WIN_OPEN: Readonly<Record<string, { why: string; engine: Val }>> = Object.fromEntries(
  ['ets-fit', 'ets-fit-2'].flatMap((set) => Object.entries(etsOpen(set)).map(([k, v]) => [`en-US:${k}`, v])),
)

export const OPEN: Readonly<Record<string, { why: string; engine: Val }>> = {
  'indirect:a1#32': { why: SENT, engine: { t: 'n', v: 6 } },
  'indirect-held:again#7': { why: SENT, engine: { t: 'n', v: 6 } },
  'indirect-held:again#8': { why: SENT, engine: { t: 'n', v: 3 } },
  'indirect-held:again#9': { why: SENT, engine: { t: 'n', v: 2 } },
  'indirect:a1#26': { why: MAC_NUM, engine: { t: 'e', v: 'DIV0' } },
  'indirect:offset#37': { why: MAC_NUM, engine: { t: 'n', v: 30 } },
  'indirect:arrays#1': { why: MAC_NUM, engine: { t: 'n', v: 30 } },
  'indirect:arrays#2': { why: MAC_NUM, engine: { t: 'n', v: 40 } },
  'indirect:arrays#3': { why: MAC_NUM, engine: { t: 'n', v: 10 } },
  'indirect:arrays#4': { why: MAC_NUM, engine: { t: 'n', v: 40 } },
  'indirect:arrays#5': { why: MAC_NUM, engine: { t: 'n', v: 9 } },
  'indirect:cell#26': { why: MAC_NUM, engine: { t: 'n', v: 5 } },
  'indirect:formulatext#4': { why: MAC_NUM, engine: { t: 'e', v: 'NA' } },
  'sumif-own-held:SUMIF#1': { why: SUMIF_OWN, engine: { t: 'n', v: 0 } },
  'sumif-own-held:SUMIF#2': { why: SUMIF_OWN, engine: { t: 'n', v: 0 } },
  'sumif-own-held:SUMIF#3': { why: SUMIF_OWN, engine: { t: 'n', v: 0 } },
  ...etsOpen('ets-fit-2'),
}

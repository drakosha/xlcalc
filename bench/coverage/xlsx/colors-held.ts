/**
 * Held out from the rounding of a positive tint (`tinted` in `src/xlsx/styles/color.ts`), which was picked from about 30 forms on the
 * 156 tinted fills of `colors.ts`: the theme of Office 2023, none of whose colours but black and white are in that set, at Excel's
 * picker tints and at others where `240 * (1 - t)` is far from a whole number, so that truncating and rounding that term part.
 * `colors-held.predicted.json` was committed before Excel ran the set; `npx tsx bench/coverage/run.ts xlsx colors-held` and
 * `npx tsx bench/windows/run.ts xlsx colors-held` read the colours into `xlsx/colors-held.json` of each Excel's answers.
 */
import { formatAddress } from '../../../src/model.js'
import { THEME, type ColorCase } from './colors.js'

const SCHEME: [string, string][] = [
  ['dk2', '0E2841'],
  ['lt2', 'E8E8E8'],
  ['accent1', '156082'],
  ['accent2', 'E97132'],
  ['accent3', '196B24'],
  ['accent4', '0F9ED5'],
  ['accent5', 'A02B93'],
  ['accent6', '4EA72E'],
  ['hlink', '467886'],
  ['folHlink', '96607D'],
]

export const THEME_HELD = SCHEME.reduce((theme, [name, hex]) => theme.replace(new RegExp(`(<a:${name}><a:srgbClr val=")[0-9A-F]{6}`), `$1${hex}`), THEME)

/** The picker's, then whole hundredths with a product of 240 near a whole number and far from one, both signs. */
const TINTS = [0.79998168889431442, 0.59999389629810485, 0.39997558519241921, -0.249977111117893, -0.499984740745262, 0.25, 0.33, 0.77, 0.15, 0.05, 0.123, 0.666, 0.01, 0.99, 0.5, 0.3, 0.7, -0.33, -0.77, -0.15, -0.05, -0.123, -0.666, -0.01, -0.99, -0.5]

export function colorHeldCases(): ColorCase[] {
  const colors = Array.from({ length: 12 }, (_, t) => TINTS.map((x) => `theme="${t}" tint="${x}"`)).flat()
  return colors.map((color, i) => ({ cell: formatAddress(i % 50, Math.floor(i / 50)), color }))
}

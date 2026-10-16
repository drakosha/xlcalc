/**
 * en-US's text of Excel's refusals of an outline command, in a module of its own: `EN_MESSAGES` holds it, and so does
 * `outlineEdit` of the `./a1` subpath, which takes no locale and none of the engine.
 */
import type { EnglishMessages, Lines } from './messages.js'

/** One line of `lines`: `row`, `column`. */
const line = (lines: Lines) => (lines === 'rows' ? 'row' : 'column')

export type OutlineMessageKey = 'outlineDeep' | 'notGrouped' | 'noSummary' | 'detailShown' | 'detailHidden'

export const OUTLINE_MESSAGES: Pick<EnglishMessages, OutlineMessageKey> = {
  outlineDeep: (p) => `the ${p.lines} ${p.span} are at the deepest outline level, 7, and are not grouped further`,
  notGrouped: (p) => `the ${p.lines} ${p.span} are in no group`,
  noSummary: (p) => `${line(p.lines)} ${p.span} is in no group with a summary ${line(p.lines)}`,
  detailShown: (p) => `the detail of ${line(p.lines)} ${p.span} is shown already`,
  detailHidden: (p) => `the detail of ${line(p.lines)} ${p.span} is hidden already`,
}

/**
 * Held out after `names-held`, where Excel took `XFD1` as LET's name and refused `R1` and `C5`, predicted by the engine
 * before Excel ran it: LET refuses a cell within Excel 2003's sheet (256 columns, IV, by 65,536 rows) and takes one past
 * it, written without `$`, as a name its calculation reads instead of the cell. The edges IV and IW, rows 65,536 and
 * 65,537, the sheet's last cell, `$` written, lower case, and the cell itself after the LET.
 */
import { type EntryCase, entryCases } from './cases.js'

const let1 = (name: string) => `=LET(${name},1,${name})`

const GENERAL: Record<string, string[]> = {
  edge: ['IV1', 'IW1', 'A65536', 'A65537', 'IV65536', 'IW65537', 'XFD1048576', 'AAA1', 'iw1'].map(let1),
  dollar: ['=LET($IW$1,1,1)', '=LET(IW$1,1,1)'],
  scope: ['=LET(IW1,5,IW1*2)', '=LET(IW1,5,IW1)+IW1', '=LET(iw1,2,IW1*3)'],
}

export const buildEntryCases = (): EntryCase[] => entryCases(GENERAL)

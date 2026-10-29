/**
 * ASC, DBCS, JIS and PHONETIC held out: full-width and half-width forms both ways (ASCII, the space, katakana with
 * their sound marks, the signs of U+FFE0–U+FFE6, characters with no other form), numbers, booleans, errors and cells;
 * PHONETIC of cells without readings (the workbook keeps none), ranges and a union.
 * \`PREDICTED\` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n, s } from '../vals.js'

/** W: texts in Japanese and in half-width katakana, a number, a boolean, a blank (A4) and an error. */
export const SHEETS: SheetCells = {
  W: { A1: '東京', B1: 'とうきょう', A2: 5, B2: 'ｶﾅ', A3: true, A5: { formula: 'NA()' }, A6: 'ｶﾞｷﾞ ab' },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // ASC: full-width letters, digits, punctuation, katakana (voiced ones in two half-width characters), signs
  ['ASC', 'ASC("ＡＢＣ１２３")', s("ABC123")],
  ['ASC', 'ASC("ガギグヴパ")', s("ｶﾞｷﾞｸﾞｳﾞﾊﾟ")],
  ['ASC', 'ASC("　")', s(" ")],
  ['ASC', 'ASC("￥￠￡￢￣￤￦")', s("¥¢£¬¯¦₩")],
  ['ASC', 'ASC("ｱｲｳ")', s("ｱｲｳ")],
  ['ASC', 'ASC("アイウ")', s("ｱｲｳ")],
  ['ASC', 'ASC("ａ－ｂ")', s("a-b")],
  ['ASC', 'ASC("Excel２０１６")', s("Excel2016")],
  ['ASC', 'ASC("゛゜")', s("ﾞﾟ")],
  ['ASC', 'ASC("ー・、。「」")', s("ｰ･､｡｢｣")],
  ['ASC', 'ASC("ァィゥェォャュョッ")', s("ｧｨｩｪｫｬｭｮｯ")],
  ['ASC', 'ASC("ヰヱヵヶ")', s("ヰヱヵヶ")],
  ['ASC', 'ASC("é")', s("é")],
  ['ASC', 'ASC("")', s("")],
  // DBCS and JIS: half-width forms to full width, a letter and its sound mark as one
  ['DBCS', 'DBCS("abc 123")', s("ａｂｃ　１２３")],
  ['DBCS', 'DBCS("~\\")', s("～＼")],
  ['DBCS', 'DBCS("ｶﾞｷﾞｸﾞ")', s("ガギグ")],
  ['DBCS', 'DBCS("ｳﾞ")', s("ヴ")],
  ['DBCS', 'DBCS("ﾊﾟﾋﾟﾌﾟ")', s("パピプ")],
  ['DBCS', 'DBCS("ｱﾞ")', s("ア゛")],
  ['DBCS', 'DBCS("ﾊﾞﾟ")', s("バ゜")],
  ['DBCS', 'DBCS("¥¢£¬¯¦₩")', s("￥￠￡￢￣￤￦")],
  ['DBCS', 'DBCS("é")', s("é")],
  ['DBCS', "DBCS(\"\"\"'\")", s("＂＇")],
  ['DBCS', 'DBCS("ｰ･､｡｢｣")', s("ー・、。「」")],
  ['DBCS', 'DBCS("ﾞ")', s("゛")],
  ['DBCS', 'DBCS("ﾡ")', s("ﾡ")],
  ['DBCS', 'DBCS("")', s("")],
  ['JIS', 'JIS("abc 123")', s("ａｂｃ　１２３")],
  ['JIS', 'JIS("~\\")', s("～＼")],
  ['JIS', 'JIS("ｶﾞｷﾞｸﾞ")', s("ガギグ")],
  ['JIS', 'JIS("ｳﾞ")', s("ヴ")],
  ['JIS', 'JIS("ﾊﾟﾋﾟﾌﾟ")', s("パピプ")],
  ['JIS', 'JIS("ｱﾞ")', s("ア゛")],
  ['JIS', 'JIS("ﾊﾞﾟ")', s("バ゜")],
  ['JIS', 'JIS("¥¢£¬¯¦₩")', s("￥￠￡￢￣￤￦")],
  ['JIS', 'JIS("é")', s("é")],
  ['JIS', "JIS(\"\"\"'\")", s("＂＇")],
  ['JIS', 'JIS("ｰ･､｡｢｣")', s("ー・、。「」")],
  ['JIS', 'JIS("ﾞ")', s("゛")],
  ['JIS', 'JIS("ﾡ")', s("ﾡ")],
  ['JIS', 'JIS("")', s("")],
  // values of other kinds and from cells
  ['ASC', 'ASC(1.5)', s("1.5")],
  ['ASC', 'ASC(TRUE)', s("TRUE")],
  ['ASC', 'ASC(NA())', e('NA')],
  ['ASC', 'ASC(W!A1)', s("東京")],
  ['ASC', 'LEN(ASC(W!A6))', n(7)],
  ['JIS', 'JIS(1.5)', s("１．５")],
  ['JIS', 'JIS(TRUE)', s("ＴＲＵＥ")],
  ['JIS', 'JIS(NA())', e('NA')],
  ['JIS', 'JIS(W!A1)', s("東京")],
  ['JIS', 'LEN(JIS(W!A6))', n(5)],
  ['DBCS', 'DBCS(1.5)', s("１．５")],
  ['DBCS', 'DBCS(TRUE)', s("ＴＲＵＥ")],
  ['DBCS', 'DBCS(NA())', e('NA')],
  ['DBCS', 'DBCS(W!A1)', s("東京")],
  ['DBCS', 'LEN(DBCS(W!A6))', n(5)],
  // PHONETIC: the cell's own text, joined over a range (Excel loads no PHONETIC of a value that is no reference)
  ['PHONETIC', 'PHONETIC(W!A1)', s("東京")],
  ['PHONETIC', 'PHONETIC(W!A2)', s("")],
  ['PHONETIC', 'PHONETIC(W!A3)', s("")],
  ['PHONETIC', 'PHONETIC(W!A4)', s("")],
  ['PHONETIC', 'PHONETIC(W!A5)', s("")],
  ['PHONETIC', 'PHONETIC(W!A1:A3)', s("東京")],
  ['PHONETIC', 'PHONETIC(W!A1:B2)', s("東京とうきょうｶﾅ")],
  ['PHONETIC', 'PHONETIC((W!A1,W!A2))', e('VALUE')],
  ['PHONETIC', 'PHONETIC(W!A6)', s("ｶﾞｷﾞ ab")],
]

/** The predictions Excel answered otherwise: in en-US both Excels give ASC's and DBCS's text as it is, know no JIS and give PHONETIC an error. */
export const MISSED = ['ASC#1', 'ASC#2', 'ASC#3', 'ASC#4', 'ASC#6', 'ASC#7', 'ASC#8', 'ASC#9', 'ASC#10', 'ASC#11', 'DBCS#1', 'DBCS#2', 'DBCS#3', 'DBCS#4', 'DBCS#5', 'DBCS#6', 'DBCS#7', 'DBCS#8', 'DBCS#10', 'DBCS#11', 'DBCS#12', 'JIS#1', 'JIS#2', 'JIS#3', 'JIS#4', 'JIS#5', 'JIS#6', 'JIS#7', 'JIS#8', 'JIS#9', 'JIS#10', 'JIS#11', 'JIS#12', 'JIS#13', 'JIS#14', 'JIS#15', 'JIS#16', 'JIS#17', 'JIS#18', 'JIS#19', 'DBCS#15', 'DBCS#16', 'DBCS#19', 'PHONETIC#1', 'PHONETIC#2', 'PHONETIC#3', 'PHONETIC#4', 'PHONETIC#5', 'PHONETIC#6', 'PHONETIC#7', 'PHONETIC#8', 'PHONETIC#9']

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

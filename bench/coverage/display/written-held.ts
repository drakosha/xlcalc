/**
 * Held out from `numfmt/refused.ts`, its last: codes a person or an agent writes, for the share of them the rule misses.
 * `REFUSED_PREDICTED` is the rule's own answer, committed before Excel for Windows ran the set; the rule is not fitted to it.
 */
import type { DisplayCase } from './cases.js'

export const CODES: string[] = ["\"Итого: \"0","Итог: 0","Итог 0","Итого 0","0 (шт)","(0)","0 [шт]","№ 0","0-0","+0","0 +","~0","≈0","0 ±","0 ч.","0,00 руб.","0 руб. 00 коп.","# ##0 руб.","#,##0.00 ₽;-#,##0.00 ₽","#,##0.00 ₽;[Red]-#,##0.00 ₽","0.0 %","0.00\"%\"","[Green]0","[Color10]0","[Blue]#,##0;[Red]-#,##0;0","dd.mm.yyyy hh:mm","dd mmm yy","yyyy-mm-dd","dd/mm/yyyy hh:mm:ss","hh:mm:ss.000","[mm]:ss","[ss].0","d-mmm","mmm-yy","q","Q1","0 Q","\"Q\"0","e","ee","yyyy e","g","ggge","b","bb","0_);(0)","#,##0_);[Red](#,##0)","0.00E+00","##0.0E+0","0.00e+00","#/#","$#,##0","$ #,##0","#,##0 $","€#,##0.00","[$₽]0","[$₽-ru-RU]0","[$-ru-RU]dd.mm","[DBNum1]0","[Hebrew]0","[ENG]0","0:0","0: 0","00:00","\"a\":0",":0","0:","шт: 0","0 :шт","0 шт.;0 шт.;0 шт.","0;0;0;0","[<1]0;0","[<1]0;0;0","[>=100]0;[<0]0;0","[Red]0;[Blue]0;[Green]0;[Yellow]@","[Red][<1]0","[<1][Red]0","dd.mm.yyyy;@","hh:mm;@","0.00;@","@;@;@","\"шт\"@","@\"шт\"","0\"шт\"","0\\ шт","0_ шт","0* шт","* 0","#,##0.00\\ \"₽\"","#,##0.00\\ ₽","#,##0\\ \"тыс.\";-#,##0\\ \"тыс.\"","0.0,,\"млн\"","0.0,, \"млн\"","0,,","#,##0.0,,","dd.mm","мм.гг","dd.мм","0 мм","0 см","0 км","0 л","0 шт","0 пач","0 упак","0 ед","0 ед.","0 ₽/шт","0 \"₽/шт\"","0 руб/шт","0 тыс. руб","0,0 млн","0.0 млн руб.","0.0 \"млн руб.\""]

export const REFUSED_PREDICTED: string[] = ["Итог: 0","0 [шт]","0.00e+00","[DBNum1]0","[Hebrew]0","[ENG]0","0:0","0: 0","00:00","\"a\":0",":0","0:","шт: 0","0 :шт","@;@;@","0 ₽/шт","0 руб/шт"]

export function buildDisplayCases(): DisplayCase[] {
  return CODES.map((format, i) => ({ id: `w${i + 1}`, cell: { value: 1234.5, format }, width: 40 }))
}

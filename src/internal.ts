/**
 * Внутренности движка для харнессов замера и адаптеров других движков поверх нашего разбора формул
 * (замер в the app, `tools/probes/wb1`: адаптер HyperFormula, сборка в QuickJS). Это не контракт:
 * состав и сигнатуры меняются без смены мажорной версии. Приложению нужен корень пакета или `./contract`.
 */
export { valuesHash } from './canon.js'
export { buildIndex, parseCrit, select, type CritColumn, type CritIndex, type CritValue } from './criteria.js'
export { createState, toCellValue, valueAt } from './engine.js'
export { FUNCTIONS, type FunctionSpec } from './functions.js'
export { dateToSerial, formatAddress, indexToCol, inputScalar, parseAddress } from './model.js'
export { clockSerial, outValue } from './result.js'
export { randAt } from './runtime.js'
export { parse, printFormula, sheetKey, validate, walk, type Ast } from './syntax.js'

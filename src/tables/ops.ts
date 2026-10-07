/** Operations on a book's tables (`restructure`): adding, renaming, deleting and converting a table. */
import type { BookInput, RestructureResult, StructureOp } from '../contract.js'
import { InputError } from '../input.js'
import type { ValuesIn } from '../structure/values.js'

export function tableOp(_book: BookInput, op: Extract<StructureOp, { op: 'addTable' | 'renameTable' | 'deleteTable' | 'convertTable' }>, _values: ValuesIn): RestructureResult {
  throw new InputError('op.op', `${op.op} is not available yet`)
}

/** The text's rotation: each kind written by `writeXlsx`, and set and changed by `writeKept` over a file. */
import type { BookInput } from '../../../src/contract.js'
import type { FieldsCase } from './cases.js'

const turned: BookInput = {
  styles: [{ alignment: { rotation: 45 } }, { alignment: { rotation: -30 } }, { alignment: { rotation: 90 } }, { alignment: { rotation: -90 } }, { alignment: { rotation: 'vertical' } }, { alignment: { horizontal: 'center', rotation: 15 } }],
  sheets: [{ name: 'S', cells: { A1: { value: 'up', style: 0 }, B1: { value: 'down', style: 1 }, C1: { value: 'up all', style: 2 }, D1: { value: 'down all', style: 3 }, E1: { value: 'stack', style: 4 }, F1: { value: 'c', style: 5 }, G1: { value: 'level' } } }],
}

export function buildFieldsCases(): FieldsCase[] {
  return [
    { id: 'written', book: turned },
    { id: 'kept', book: turned, kept: { patches: [{ sheet: 'S', range: 'G1', set: { style: { alignment: { rotation: 60 } } } }, { sheet: 'S', range: 'A1', set: { style: { alignment: { rotation: 'vertical' } } } }] } },
  ]
}

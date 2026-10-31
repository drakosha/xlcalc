import { expect, it } from 'vitest'
import { comEdit } from '../bench/windows/books.js'

it('sends the exact double instead of relying on the Windows decimal parser', () => {
  expect(comEdit(0.00743244648818)).toMatchObject({ kind: 'n', bits: '3f7e717c29893a99' })
  expect(comEdit(-0.00743244648818)).toMatchObject({ kind: 'n', bits: 'bf7e717c29893a99' })
  expect(comEdit({ formula: '1+2' })).toEqual({ kind: 'f2', value: '=1+2' })
})

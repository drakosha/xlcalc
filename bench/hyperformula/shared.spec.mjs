import assert from 'node:assert/strict'
import { test } from 'node:test'
import { close, hfInput, open, prepareEdit, reference, verify } from './shared.mjs'

const book = {
  sheets: [{ name: 'S', cells: {
    A1: { value: '=1+1' }, A2: { value: '007' }, A3: { value: 'TRUE' }, A4: { value: true },
    A5: { value: { date: '1900-02-29' } }, A6: { value: "'text" }, A7: { value: 7 },
    B1: { formula: 'SUM(A4:A7)' }, B2: { formula: '1/0' }, B3: { formula: 'IF(TRUE,A7,0)' }, B4: { formula: 'Rate*2' },
  } }], names: [{ name: 'Rate', ref: "'S'!$A$7" }],
}

for (const engine of ['xl-closures', 'xl-eval', 'hf-default', 'hf-index', 'hf-sparse', 'hf-auto']) {
  test(`${engine}: literals, explicit dates, booleans, errors, names and contiguous edits agree`, () => {
    const s = open(engine, book, hfInput(book))
    try {
      assert.equal(verify(engine, s, book, reference(book)).differences, 0)
      const edits = [{ sheet: 'S', cell: 'A7', input: { value: 3 } }]
      prepareEdit(engine, s, edits)()
      assert.equal(verify(engine, s, book, reference(book, edits)).differences, 0)
      const bulk = [{ sheet: 'S', cell: 'A1', input: { value: 'still text' } }, { sheet: 'S', cell: 'A2', input: { value: 2 } }, { sheet: 'S', cell: 'A3', input: { value: false } }, { sheet: 'S', cell: 'A4', input: { value: true } }, { sheet: 'S', cell: 'A5', input: { value: { date: '1900-02-29' } } }, { sheet: 'S', cell: 'A6', input: { value: "'text" } }, ...edits]
      prepareEdit(engine, s, bulk)()
      assert.equal(verify(engine, s, book, reference(book, bulk)).differences, 0)
    } finally { close(engine, s) }
  })
}

test('the parity check detects a deliberately changed input', () => {
  const s = open('hf-default', book, hfInput(book))
  try {
    prepareEdit('hf-default', s, [{ sheet: 'S', cell: 'A7', input: { value: 99 } }])()
    assert.ok(verify('hf-default', s, book, reference(book)).differences > 0)
  } finally { close('hf-default', s) }
})

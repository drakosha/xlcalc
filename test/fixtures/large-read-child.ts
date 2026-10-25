/**
 * A session's reads of a rectangle of more than 2^27 cells with `maxCells` raised over it, in a fresh process with a heap
 * of 4 GB (for `review-1021-arrays.test.ts`): `getValues` gives rows of values and `copyText` stops at its text's limit,
 * neither being one array. Prints what each gave as JSON.
 */
const { libUrl } = await import('./target.js')
const { CalcError, createSession } = (await import(libUrl('index.js'))) as typeof import('../../src/index.js')

const s = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B2: { value: 2 } } }] })
const tried = (f: () => unknown) => {
  try {
    return f()
  } catch (e) {
    return e instanceof CalcError ? { code: e.code, message: e.message } : { thrown: (e as Error).name }
  }
}
const values = tried(() => {
  const v = s.getValues('S', 'A1:XFD8193', { maxCells: 200_000_000 })
  return [v.length, v[0].length, v[0][0], v[1][1]]
})
const text = tried(() => s.copyText('S', 'B1:XFD8193', { maxCells: 200_000_000 }).length)
console.log(JSON.stringify({ values, text }))

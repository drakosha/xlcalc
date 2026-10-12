import { synthBook, mulberry32, S } from '../../dist/synth.js'
import { formatAddress } from '../../dist/model.js'

const input = (sheet, cell, value) => ({ sheet, cell, input: { value } })
const oneSheet = (cells) => ({ sheets: [{ name: 'S', cells }] })

function business(rows) {
  const { book } = synthBook({ seed: 1, rows })
  const cell = `D${Math.floor(rows / 2) + 2}`
  const amount = book.sheets[0].cells[cell].value
  return {
    book,
    operations: [
      { id: 'criterion', edits: (i) => [input(S.params, 'A10', i % 2 ? 'Чай*' : 'Аренда*')] },
      { id: 'data', edits: (i) => [input(S.data, cell, amount + (i % 2 ? 0 : 1))] },
    ],
  }
}

function prefixes(rows, statistics = false) {
  const cells = {}
  for (let r = 1; r <= rows; r++) {
    cells[`A${r}`] = { value: statistics ? (r % 97) + 1 : 1 }
    const range = `$A$1:A${r}`
    cells[`B${r}`] = { formula: statistics ? `AVERAGE(${range})+COUNTIF(${range},">0")` : `SUM(${range})` }
  }
  const cell = `A${Math.floor(rows / 2)}`
  const value = cells[cell].value
  return { book: oneSheet(cells), operations: [{ id: 'middle-input', edits: (i) => [input('S', cell, value + (i % 2 ? 0 : 1))] }] }
}

function lookups(rows = 34_000, queries = 1000) {
  const random = mulberry32(123)
  const keys = Array.from({ length: rows }, (_, i) => i + 1)
  for (let i = rows - 1; i > 0; i--) {
    const j = Math.floor(random() * (i + 1))
    ;[keys[i], keys[j]] = [keys[j], keys[i]]
  }
  const cells = {}
  for (let r = 1; r <= rows; r++) {
    cells[`A${r}`] = { value: keys[r - 1] }
    cells[`B${r}`] = { value: keys[r - 1] * 3 }
  }
  for (let r = 1; r <= queries; r++) {
    cells[`D${r}`] = { value: keys[rows - r] }
    cells[`E${r}`] = { formula: `VLOOKUP(D${r},$A$1:$B$${rows},2,FALSE)` }
  }
  return {
    book: oneSheet(cells), indexed: true,
    operations: [
      { id: 'lookup-key', edits: (i) => [input('S', 'D1', i % 2 ? keys[rows - 1] : keys[0])] },
      { id: 'lookup-data', edits: (i) => [input('S', 'B1', keys[0] * 3 + (i % 2 ? 0 : 1))] },
    ],
  }
}

function chain(rows = 10_000) {
  const cells = { A1: { value: 1 } }
  for (let r = 2; r <= rows; r++) cells[`A${r}`] = { formula: `A${r - 1}+1` }
  return { book: oneSheet(cells), operations: [{ id: 'root', edits: (i) => [input('S', 'A1', i % 2 ? 1 : 2)] }] }
}

function fanout(rows = 50_000) {
  const cells = { A1: { value: 1 }, C1: { formula: `SUM(B1:B${rows})` } }
  for (let r = 1; r <= rows; r++) cells[`B${r}`] = { formula: `$A$1*${r}` }
  return { book: oneSheet(cells), operations: [{ id: 'root', edits: (i) => [input('S', 'A1', i % 2 ? 1 : 2)] }] }
}

function bulk(rows = 2000) {
  const cells = {}
  for (let r = 1; r <= rows; r++) {
    cells[`A${r}`] = { value: r }
    cells[`B${r}`] = { formula: `A${r}*3+SUM($D$1:$D$20)` }
  }
  for (let r = 1; r <= 20; r++) cells[`D${r}`] = { value: r }
  return { book: oneSheet(cells), operations: [{ id: '2000-inputs', edits: (i) => Array.from({ length: rows }, (_, r) => input('S', `A${r + 1}`, r + 1 + (i % 2 ? 0 : 1))) }] }
}

function sparse(rows = 1000) {
  const cells = {}
  for (let i = 0; i < rows; i++) {
    const addr = formatAddress(i * 97, (i % 8) * 31)
    cells[addr] = { value: i + 1 }
    cells[formatAddress(i * 97, 310)] = { formula: `${addr}*3` }
  }
  return { book: oneSheet(cells), operations: [{ id: 'one-input', edits: (i) => [input('S', 'A1', i % 2 ? 1 : 2)] }] }
}

export const WORKLOADS = [
  { id: 'business-3400', title: 'SUMIFS business book, 3400 data rows', create: () => business(3400) },
  { id: 'prefix-sum', title: '5000 cumulative SUM formulas', create: () => prefixes(5000) },
  { id: 'prefix-statistics', title: '5000 cumulative AVERAGE + COUNTIF formulas', create: () => prefixes(5000, true) },
  { id: 'lookup', title: '1000 exact VLOOKUPs over 34000 unsorted rows', create: () => lookups() },
  { id: 'chain', title: '10000 cells in one dependency chain', create: () => chain() },
  { id: 'fanout', title: '50000 direct dependents and one SUM', create: () => fanout() },
  { id: 'bulk', title: '2000 inputs replaced in one operation', create: () => bulk() },
  { id: 'sparse', title: '2000 cells in a roughly 97000 x 311 sheet', create: () => sparse() },
  { id: 'business-34000', title: 'SUMIFS business book, 34000 data rows', create: () => business(34000), hfProbe: true },
]

export function workload(id) {
  const spec = WORKLOADS.find((s) => s.id === id)
  if (!spec) throw new Error(`Unknown workload ${id}`)
  return { ...spec, ...spec.create() }
}

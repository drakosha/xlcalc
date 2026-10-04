/**
 * Синтетическая книга по сиду (без реальных данных): лист данных 3 400 × 15 и шесть листов вокруг,
 * около 2 800 SUMIFS и 700 SUM, ссылки на данные тремя способами: целая колонка (`ПоМесяцам`,
 * `ПоНаправлениям`), ограниченный диапазон (`ПоПроектам`), имя (`ПоКатегориям`).
 */
import type { BookInput, CellEdit, CellInput } from '../../src/contract'
import { formatAddress, indexToCol } from '../../src/model'

export function mulberry32(seed: number): () => number {
  let a = seed >>> 0
  return () => {
    a = (a + 0x6d2b79f5) | 0
    let t = Math.imul(a ^ (a >>> 15), 1 | a)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4_294_967_296
  }
}

export interface SynthOptions {
  seed: number
  /** Строк данных (по умолчанию 3 400; короткая книга тестов 340). */
  rows?: number
}

export interface SynthBook {
  book: BookInput
  /** Ячейки правки класса А: справочник и критерии, не лист данных. */
  classA: { sheet: string; cell: string }[]
  dataRows: number
  items: string[]
}

export const S = {
  data: 'Данные',
  items: 'Статьи',
  params: 'Параметры',
  months: 'ПоМесяцам',
  projects: 'ПоПроектам',
  cats: 'ПоКатегориям',
  dirs: 'ПоНаправлениям',
} as const

const WORDS = ['Чай', 'Чаи', 'Ёлка', 'Елка', 'Аренда', 'Связь', 'Налог', 'Доход', 'Сырьё', 'Сырье']
const GROUPS = 8
const PER_GROUP = 5
const N_ITEMS = GROUPS * PER_GROUP
const N_CATS = 20
const N_PROJ = 15
const DIRS = ['приход', 'расход'] as const
const NOTES = ['=1+1', 'TRUE', '25%', '2025-01-01', "'с кавычкой", 'он сказал "да"', 'путь\\к\\файлу', 'обычный текст', '', '  пробелы  ']

const q = (s: string) => `'${s.replaceAll("'", "''")}'`
const iso = (y: number, m: number, d: number) => `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`

export function synthBook(opts: SynthOptions): SynthBook {
  const rnd = mulberry32(opts.seed)
  const n = opts.rows ?? 3_400
  const pick = <T>(xs: readonly T[]) => xs[Math.floor(rnd() * xs.length)]
  const items = Array.from({ length: N_ITEMS }, (_, i) => `${WORDS[i % WORDS.length]} ${String(Math.floor(i / WORDS.length) + 1)}`)
  const cats = Array.from({ length: N_CATS }, (_, i) => `Кат-${String(i + 1).padStart(2, '0')}`)
  const projs = Array.from({ length: N_PROJ }, (_, i) => `П${String(i + 1).padStart(2, '0')}`)

  // ---------- данные ----------
  const data: Record<string, CellInput> = {}
  const header = ['Дата', 'Статья', 'Категория', 'Сумма', 'Направление', 'Проект', 'Примечание', 'Кол-во', 'Цена', 'Срок', 'Флаг', 'Доля', 'Вес', 'Код', 'Остаток']
  header.forEach((h, c) => (data[formatAddress(0, c)] = { value: h }))
  for (let i = 0; i < n; i++) {
    const r = i + 1
    const month = 1 + Math.floor(rnd() * 12)
    const day = 1 + Math.floor(rnd() * 28)
    const big = rnd() < 0.01
    const kop = Math.floor(rnd() * (big ? 2e11 : 5e7)) - (big ? 0 : 1e7)
    const row: CellInput[] = [
      { value: { date: iso(2025, month, day) } },
      { value: pick(items) },
      { value: pick(cats) },
      { value: kop / 100 },
      { value: pick(DIRS) },
      { value: pick(projs) },
      { value: pick(NOTES) },
      { value: Math.floor(rnd() * 100) },
      { value: Math.round(rnd() * 1e6) / 100 },
      { value: { date: iso(2025 + Math.floor(rnd() * 2), 1 + Math.floor(rnd() * 12), 1 + Math.floor(rnd() * 28)) } },
      { value: rnd() < 0.5 },
      { value: Math.round(rnd() * 1000) / 1000 },
      { value: rnd() * 10 },
      { value: `K${Math.floor(rnd() * 1e6)}` },
      { value: Math.round((rnd() - 0.5) * 1e8) / 100 },
    ]
    row.forEach((cell, c) => {
      if (cell.value !== '') data[formatAddress(r, c)] = cell
    })
  }
  const last = n + 1
  const col = (c: string) => `${q(S.data)}!$${c}:$${c}`
  const bounded = (c: string) => `${q(S.data)}!$${c}$2:$${c}$${last}`

  // ---------- справочник статей ----------
  const itemsSheet: Record<string, CellInput> = { A1: { value: 'Статья' }, B1: { value: 'Группа' }, C1: { value: 'Название группы' }, E1: { value: 'Группа' }, F1: { value: 'Название' } }
  for (let g = 0; g < GROUPS; g++) {
    itemsSheet[`E${g + 2}`] = { value: g + 1 }
    itemsSheet[`F${g + 2}`] = { value: `Группа «${g + 1}»` }
  }
  items.forEach((it, i) => {
    itemsSheet[`A${i + 2}`] = { value: it }
    itemsSheet[`B${i + 2}`] = { value: Math.floor(i / PER_GROUP) + 1 }
    itemsSheet[`C${i + 2}`] = { formula: `XLOOKUP(B${i + 2},$E$2:$E$${GROUPS + 1},$F$2:$F$${GROUPS + 1},"нет")` }
  })

  // ---------- параметры ----------
  const P = q(S.params)
  const params: Record<string, CellInput> = {
    A1: { value: 'Начало' },
    B1: { value: { date: '2025-01-01' } },
    A2: { value: 'Конец' },
    B2: { formula: 'EOMONTH(B1,11)' },
    A3: { value: 'Приход' },
    B3: { value: DIRS[0] },
    A4: { value: 'Расход' },
    B4: { value: DIRS[1] },
    A5: { value: 'Начало DATE' },
    B5: { formula: 'DATE(2025,1,1)' },
    // хвост семантики: COUNTIFS с шаблоном, IF, IFERROR, ROUND, TEXT
    A10: { value: 'Чай*' },
    B10: { formula: `COUNTIFS(${col('B')},"Чай*")` },
    B11: { formula: `COUNTIFS(${col('B')},A10,${col('E')},B3)` },
    B12: { formula: `IFERROR(${q(S.months)}!N3/${q(S.months)}!N4,0)` },
    B13: { formula: 'ROUND(B12,2)' },
    B14: { formula: 'IF(B12>1,"больше","не больше")' },
    B15: { formula: 'TEXT(B1,"YYYY-MM-DD")' },
    B16: { formula: 'TEXT(B13,"0.00")' },
    B17: { formula: `IFERROR(1/(B3="нет"),-1)` },
    B18: { formula: `SUMIFS(${col('D')},${col('B')},"Чаи*")` },
  }

  const book: BookInput = {
    sheets: [
      { name: S.data, cells: data },
      { name: S.items, cells: itemsSheet },
      { name: S.params, cells: params },
    ],
    names: [
      { name: 'data_amt', ref: bounded('D') },
      { name: 'data_item', ref: bounded('B') },
      { name: 'data_cat', ref: bounded('C') },
      { name: 'data_dir', ref: bounded('E') },
    ],
  }

  /**
   * Лист свода: строки статей по группам с подытогом группы, колонки из `cols`; первая колонка
   * ссылается на справочник, последняя это SUM по строке.
   */
  const pivot = (name: string, headRows: number, cols: number, head: (c: number, r: number) => CellInput | undefined, cell: (r: number, c: number) => string) => {
    const cells: Record<string, CellInput> = { A1: { value: name } }
    for (let r = 0; r < headRows; r++)
      for (let c = 0; c < cols; c++) {
        const h = head(c, r)
        if (h) cells[`${indexToCol(c + 1)}${r + 1}`] = h
      }
    const total = indexToCol(cols + 1)
    let row = headRows + 1
    for (let g = 0; g < GROUPS; g++) {
      const first = row
      for (let k = 0; k < PER_GROUP; k++) {
        const i = g * PER_GROUP + k
        cells[`A${row}`] = { formula: `${q(S.items)}!A${i + 2}` }
        for (let c = 0; c < cols; c++) cells[`${indexToCol(c + 1)}${row}`] = { formula: cell(row, c) }
        cells[`${total}${row}`] = { formula: `SUM(B${row}:${indexToCol(cols)}${row})` }
        row++
      }
      cells[`A${row}`] = { value: `Итого группа ${g + 1}` }
      for (let c = 0; c <= cols; c++) {
        const L = indexToCol(c + 1)
        cells[`${L}${row}`] = { formula: `SUM(${L}${first}:${L}${row - 1})` }
      }
      row++
    }
    book.sheets.push({ name, cells })
  }

  // ПоМесяцам: целые колонки; строки 1–2 начало и конец месяца
  pivot(
    S.months,
    2,
    12,
    (c, r) => {
      const L = indexToCol(c + 1)
      if (r === 0) return { formula: c === 0 ? `${P}!B1` : `EOMONTH(${indexToCol(c)}1,0)+1` }
      return { formula: `EOMONTH(${L}1,0)` }
    },
    (r, c) => {
      const L = indexToCol(c + 1)
      return `SUMIFS(${col('D')},${col('B')},$A${r},${col('A')},">="&${L}$1,${col('A')},"<="&${L}$2)`
    },
  )
  // ПоПроектам: ограниченные диапазоны, проект в строке 1
  pivot(
    S.projects,
    1,
    N_PROJ,
    (c) => ({ value: projs[c] }),
    (r, c) => {
      const L = indexToCol(c + 1)
      return `SUMIFS(${bounded('D')},${bounded('B')},$A${r},${bounded('F')},${L}$1,${bounded('A')},">="&${P}!$B$1,${bounded('A')},"<="&${P}!$B$2,${bounded('E')},${P}!$B$4)`
    },
  )
  // ПоКатегориям: имена
  pivot(
    S.cats,
    1,
    N_CATS,
    (c) => ({ value: cats[c] }),
    (r, c) => `SUMIFS(data_amt,data_item,$A${r},data_cat,${indexToCol(c + 1)}$1,data_dir,${P}!$B$3)`,
  )
  // ПоНаправлениям: целые колонки; 12 месяцев × 2 направления, строки 1–3 начало, конец, направление
  pivot(
    S.dirs,
    3,
    24,
    (c, r) => {
      const m = Math.floor(c / 2)
      if (r === 0) return { formula: `EOMONTH(${P}!$B$1,${m - 1})+1` }
      if (r === 1) return { formula: `EOMONTH(${P}!$B$1,${m})` }
      return { formula: `${P}!$B$${c % 2 === 0 ? 3 : 4}` }
    },
    (r, c) => {
      const L = indexToCol(c + 1)
      return `SUMIFS(${col('D')},${col('B')},$A${r},${col('E')},${L}$3,${col('A')},">="&${L}$1,${col('A')},"<="&${L}$2)`
    },
  )

  const classA = [
    ...items.map((_, i) => ({ sheet: S.items, cell: `A${i + 2}` })),
    { sheet: S.params, cell: 'B1' },
    { sheet: S.params, cell: 'B3' },
    { sheet: S.params, cell: 'B4' },
  ]
  return { book, classA, dataRows: n, items }
}

/** 100 правок класса А по сиду: другое имя статьи, другое начало периода, другое направление. */
export function classAEdits(s: SynthBook, seed: number, count = 100): CellEdit[] {
  const rnd = mulberry32(seed)
  const out: CellEdit[] = []
  for (let i = 0; i < count; i++) {
    const t = s.classA[Math.floor(rnd() * s.classA.length)]
    let input: CellInput
    if (t.sheet === S.items) input = { value: rnd() < 0.8 ? s.items[Math.floor(rnd() * s.items.length)] : `Новая ${i}` }
    else if (t.cell === 'B1') input = { value: { date: iso(2025, 1 + Math.floor(rnd() * 3), 1) } }
    else input = { value: DIRS[Math.floor(rnd() * 2)] }
    out.push({ ...t, input })
  }
  return out
}

/** Правка класса Б: сумма в случайной строке листа данных. */
export function classBEdit(s: SynthBook, seed: number): CellEdit {
  const rnd = mulberry32(seed)
  const row = 2 + Math.floor(rnd() * s.dataRows)
  return { sheet: S.data, cell: `D${row}`, input: { value: Math.round(rnd() * 1e8) / 100 } }
}

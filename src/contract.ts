/**
 * Контракт движка формул `@mike/formula` (подпуть `./contract`; в the app трек «книга», `docs/workbook-design.md`, раздел 5a).
 * Не зависит от реализации: здесь только типы. Любой движок (сейчас компиляция в JS, вариант C
 * замера WB1) стоит за этим контрактом, тесты поведения пакета это его приёмка.
 *
 * Формула хранится в поле `formula` без `=`, отдельно от значения: строка из данных никогда не
 * становится формулой. Имена функций английские, разделитель аргументов запятая, десятичная точка.
 */

export type Scalar = number | string | boolean | null

export type JsonValue = Scalar | JsonValue[] | { [key: string]: JsonValue }

/** Ячейка на входе. Дата приходит явно (`{ date: '2025-01-31' }`) и считается серийным номером Excel. */
export interface CellInput {
  value?: Scalar | { date: string }
  /** Текст формулы без `=`. Если есть, `value` не используется. */
  formula?: string
  /** Формат показа, проходит насквозь. */
  format?: string
  /** Произвольные данные ячейки (например, id строк источника), проходят насквозь. */
  meta?: JsonValue
}

export interface SheetInput {
  name: string
  /** Ячейки по адресу A1 (`B12`). */
  cells: Record<string, CellInput>
}

export interface NameInput {
  /** Латиница, цифры, `_` и `.`; не адрес ячейки, не R1C1, не TRUE/FALSE. */
  name: string
  /** Ссылка с листом: `'Данные'!$D$2:$D$3401`, `'Данные'!$D:$D`. */
  ref: string
  /** Лист, на котором имя видно; локальное имя перекрывает глобальное. Без него имя глобальное. */
  scope?: string
}

export interface BookInput {
  sheets: SheetInput[]
  names?: NameInput[]
}

export type ErrorCode = 'DIV0' | 'NA' | 'NAME' | 'NUM' | 'REF' | 'VALUE' | 'CYCLE' | 'SPILL' | 'ERROR'

/** Значение после расчёта. Даты выходят серийными номерами, числа округлены до 15 значащих цифр. */
export type CellValue = Scalar | { error: ErrorCode }

/** Коды проверки формулы. `unsupported`: вычисление над массивом или внешняя функция не с литералами. */
export type CheckCode =
  | 'empty'
  | 'syntax'
  | 'too_long'
  | 'too_deep'
  | 'external_ref'
  | 'unknown_function'
  | 'unknown_sheet'
  | 'unknown_name'
  | 'unsupported'

/** Момент (ISO) и зерно случайности, от которых считаются TODAY, NOW, RAND, RANDBETWEEN. */
export interface Clock {
  now: string
  seed: number
}

export interface CellRef {
  sheet: string
  cell: string
}

/**
 * Внешние функции (функции пакета, срез после книги): их значения разрешаются заранее, одним
 * запросом, и передаются в расчёт. Ключ вызова это его канонический текст (`BALANCE("2025-01")`),
 * аргументы только литералы.
 */
export interface Externals {
  functions: string[]
  values?: Record<string, CellValue>
}

export interface ComputeOptions {
  clock: Clock
  externals?: Externals
}

export interface CellResult {
  value: CellValue
  /** Канонический текст формулы (лист всегда в кавычках, функции заглавными). */
  formula?: string
  format?: string
  meta?: JsonValue
}

/** Ячейка с ошибкой и её причина: неверная формула (с кодом проверки и местом) или ошибка расчёта. */
export interface CellIssue extends CellRef {
  error: ErrorCode
  reason: string
  check?: CheckCode
  /** Позиция в тексте формулы для `check`. */
  at?: number
}

export interface NameIssue {
  name: string
  scope?: string
  reason: string
}

export interface ComputeResult {
  sheets: { name: string; cells: Record<string, CellResult> }[]
  issues: CellIssue[]
  nameIssues: NameIssue[]
  /** Ячейки, которые зависят от волатильных функций (прямо или через другие ячейки). */
  volatile: CellRef[]
  /** Часы, с которыми посчитано (у живой сессии последние). */
  clock: Clock
  /** sha256 канонического текста значений вместе с версией конфигурации движка. */
  hash: string
  /** sha256 входа (листы, ячейки, имена; порядок не важен). */
  inputHash: string
  configVersion: string
  ms: number
}

export interface ValidateContext {
  /** Листы книги; без них листы не проверяются. */
  sheets?: string[]
  /** Имена книги; без них имена не проверяются. */
  names?: string[]
  externalFunctions?: string[]
}

export type ValidateResult =
  | { ok: true; canonical: string; functions: string[]; sheets: string[]; names: string[]; externals: string[] }
  | { ok: false; code: CheckCode; message: string; at: number }

export type NameCheck = { ok: true } | { ok: false; message: string }

/** Правка ячейки; `input: null` очищает ячейку. */
export interface CellEdit extends CellRef {
  input: CellInput | null
}

export interface CellChange extends CellRef {
  value: CellValue
}

export interface ExplainOptions {
  /** Глубина по предшественникам, по умолчанию 3. */
  depth?: number
  /** Потолок узлов, по умолчанию 200. */
  maxNodes?: number
  /** Раскрывать диапазоны в непустые ячейки (по умолчанию диапазон один узел). */
  expand?: boolean
}

export interface ExplainNode {
  /** `'Лист'!B3` или `'Лист'!D:D`. */
  ref: string
  value?: CellValue
  formula?: string
  /** Для диапазона: непустых ячеек. */
  cells?: number
  children: ExplainNode[]
}

export interface Explanation {
  root: ExplainNode
  nodes: number
  truncated: boolean
  /** Дерево текстом, для агента. */
  text: string
}

export type SessionOptions = ({ clock: Clock } | { live: true }) & { externals?: Externals }

/**
 * Сессия над книгой. Живая (`live`) берёт часы и новое зерно на каждый пересчёт; текущие часы
 * сессии (`clock`) записываются в версию. Не живая считает с заданными часами.
 */
export interface Session {
  readonly clock: Clock
  /** Атомарно, один пересчёт; изменившиеся значения. */
  setCells(edits: CellEdit[]): CellChange[]
  setNames(names: NameInput[]): CellChange[]
  /** Только у не живой сессии. */
  setClock(clock: Clock): CellChange[]
  explain(sheet: string, cell: string, opts?: ExplainOptions): Explanation
  result(): ComputeResult
  close(): void
}

export interface FormulaEngine {
  compute(book: BookInput, opts: ComputeOptions): ComputeResult
  validateFormula(text: string, ctx?: ValidateContext): ValidateResult
  validateName(name: string): NameCheck
  openSession(book: BookInput, opts: SessionOptions): Session
}

export const MAX_FORMULA_LENGTH = 8_192
export const MAX_FORMULA_DEPTH = 64

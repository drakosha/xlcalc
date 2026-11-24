/** Replace modeled print settings in the final parts of a workbook written over its source file. */
import { sheetKey } from '../../syntax/print.js'
import { BEFORE_NAMES, type KeptPartView } from '../keep/book.js'
import type { Edit } from '../keep/carriers.js'
import { edited } from '../keep/refusals.js'
import { WORKSHEET_CHILD_ORDER } from '../keep/sheet.js'
import { attrsOf, childrenOf, closing, opening, rootOf, setAttrs, split, type Placed } from '../keep/xml.js'
import { readText } from '../xml/text.js'
import type { WorksheetPrintParts } from '../write/book.js'
import { printParts } from './write.js'
import { PROPERTIES, SETUP } from './xml.js'

const PRINT_CHILDREN = ['printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks'] as const
const SHEET_PROPERTIES = ['tabColor', 'outlinePr', 'pageSetUpPr'] as const

function qualifier(name: string): string {
  return name.slice(0, name.indexOf(':') + 1)
}

/** Our generated text has escaped cell labels and header text, so literal tags are all serializer-owned. */
function qualify(xml: string, prefix: string): string {
  return prefix ? xml.replace(/<(\/?)([a-zA-Z][\w]*)/g, '<$1' + prefix + '$2') : xml
}

function editText(text: string, edits: Omit<Edit, 'part'>[]): string {
  return edited(text, edits.map((edit) => ({ part: '', ...edit })))
}

function replaceChildren(xml: string, parent: Placed, order: readonly string[], changes: ReadonlyMap<string, string>): string {
  const children = childrenOf(xml, parent)
  const insertions = new Map<number, string>()
  const changesInOrder = order.filter((key) => changes.has(key))
  if (parent.end === parent.open) {
    const body = changesInOrder.map((key) => changes.get(key) ?? '').join('')
    if (!body) return xml
    const tag = opening(xml.slice(parent.start, parent.open))
    return editText(xml, [{ start: parent.start, end: parent.end, text: `${tag}${body}</${parent.qname}>` }])
  }
  const edits: Omit<Edit, 'part'>[] = []
  for (const key of changesInOrder) {
    const text = changes.get(key) ?? ''
    const current = children.find((node) => node.local === key)
    if (current) {
      edits.push({ start: current.start, end: current.end, text })
      continue
    }
    if (!text) continue
    const later = children.find((node) => order.indexOf(node.local) > order.indexOf(key))
    const at = later?.start ?? xml.lastIndexOf('</', parent.end - 1)
    insertions.set(at, (insertions.get(at) ?? '') + text)
  }
  for (const [at, text] of insertions) edits.push({ start: at, end: at, text })
  return editText(xml, edits)
}

/** Keep private attributes such as the printer settings relationship when replacing the modeled page attributes. */
function replaceAttributes(current: string, replacement: string, keys: string[]): string {
  if (!current) return replacement
  const old = split(current)
  const root = replacement ? rootOf(replacement) : undefined
  const attrs = root ? attrsOf(replacement, root) : {}
  const tag = setAttrs(old.tag, Object.fromEntries(keys.map((key) => [key, attrs[key] ?? null])))
  const self = closing(tag)
  if (old.body) return `${opening(tag)}${old.body}</${old.name}>`
  return Object.keys(attrsOf(self, rootOf(self)!)).length ? self : ''
}

function replacePrintWorksheet(xml: string, print: WorksheetPrintParts): string {
  const root = rootOf(xml)!
  const children = childrenOf(xml, root)
  const prefix = qualifier(root.qname)
  const currentProperties = children.find((node) => node.local === 'sheetPr')
  const updates = new Map<string, string>()
  if (currentProperties) {
    const original = xml.slice(currentProperties.start, currentProperties.end)
    const properties = rootOf(original)!
    const oldPage = childrenOf(original, properties).find((node) => node.local === 'pageSetUpPr')
    const oldSetup = oldPage ? original.slice(oldPage.start, oldPage.end) : ''
    const nextSetup = replaceAttributes(oldSetup, qualify(print.properties, qualifier(properties.qname)), Object.keys(PROPERTIES))
    if (nextSetup !== oldSetup) {
      const patch = new Map([['pageSetUpPr', nextSetup]])
      updates.set('sheetPr', replaceChildren(original, properties, SHEET_PROPERTIES, patch))
    }
  } else if (print.properties) {
    const tag = prefix + 'sheetPr'
    updates.set('sheetPr', `<${tag}>${qualify(print.properties, prefix)}</${tag}>`)
  }

  const parts = '<settings>' + print.settings + '</settings>'
  const generated = rootOf(parts)!
  const newChildren = childrenOf(parts, generated)
  for (const key of PRINT_CHILDREN) {
    const fresh = newChildren.find((part) => part.local === key)
    const replacement = fresh ? qualify(parts.slice(fresh.start, fresh.end), prefix) : ''
    const old = children.find((node) => node.local === key)
    if (key === 'pageSetup' && old) {
      const preserve = replaceAttributes(xml.slice(old.start, old.end), replacement, Object.keys(SETUP))
      updates.set(key, preserve)
    } else updates.set(key, replacement)
  }
  return replaceChildren(xml, root, WORKSHEET_CHILD_ORDER, updates)
}

function replacePrintNames(
  xml: string, root: Placed, children: Placed[], names: string[], indices: ReadonlySet<number>,
): string {
  const current = children.find((node) => node.local === 'definedNames')
  const prefix = qualifier(current?.qname ?? root.qname)
  const additions = names.map((name) => qualify(name, prefix)).join('')
  if (!current) {
    if (!additions) return xml
    // AlternateContent for workbook properties can precede <sheets>. Anchor after the last known predecessor.
    const before = children.findLast((node) => BEFORE_NAMES.includes(node.local))
    const at = before?.end ?? root.open
    const tag = prefix + 'definedNames'
    return editText(xml, [{ start: at, end: at, text: `<${tag}>${additions}</${tag}>` }])
  }

  const original = xml.slice(current.start, current.end)
  const node = rootOf(original)!
  const removed: Omit<Edit, 'part'>[] = []
  for (const name of childrenOf(original, node)) {
    if (name.local !== 'definedName') continue
    const attrs = attrsOf(original, name)
    if (indices.has(Number(attrs.localSheetId)) && /^_xlnm\.(Print_Area|Print_Titles)$/i.test(readText(attrs.name ?? ''))) {
      removed.push({ start: name.start, end: name.end, text: '' })
    }
  }
  if (!additions && !removed.length) return xml
  const left = split(editText(original, removed))
  const body = left.body + additions
  return editText(xml, [{
    start: current.start, end: current.end,
    text: body.trim() ? `${opening(left.tag)}${body}</${left.name}>` : '',
  }])
}

export function replaceKeptPrint(view: KeptPartView, print: unknown): boolean {
  const { book, src: source, parts, sheetParts, id } = view
  const workbook = parts.text.get(source.wbPath.toLowerCase()) ?? source.wb
  const root = rootOf(workbook)!
  const children = childrenOf(workbook, root)
  const declarations = children.find((node) => node.local === 'sheets')!
  const place = new Map<string, number>()
  childrenOf(workbook, declarations).filter((node) => node.local === 'sheet').forEach((node, i) => {
    place.set(sheetKey(readText(attrsOf(workbook, node).name ?? '')), i)
  })
  const changes = printParts(book, print, place)
  let changed = false
  const indices = new Set<number>()
  const names: string[] = []
  for (const [i, sheet] of book.sheets.entries()) {
    const settings = changes.get(sheet.name)
    if (!settings) continue
    const position = place.get(sheetKey(sheet.name))
    if (position === undefined) throw new Error(`xlsx/print/keep: no workbook entry for ${JSON.stringify(sheet.name)}`)
    indices.add(position)
    names.push(...settings.names)
    const part = sheetParts[i]
    const original = parts.text.get(part.toLowerCase()) ?? source.sheets[id.source[i]!].xml
    const xml = replacePrintWorksheet(original, settings)
    if (xml !== original) {
      parts.set(part, xml)
      changed = true
    }
  }
  const nextWorkbook = replacePrintNames(workbook, root, children, names, indices)
  if (nextWorkbook === workbook) return changed
  parts.set(source.wbPath, nextWorkbook)
  return true
}

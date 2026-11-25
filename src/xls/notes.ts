import { formatAddress } from '../model.js'
import { escapeText } from '../xlsx/xml/text.js'
import type { Xml } from '../xlsb/xml.js'
import { combinedPayload, CONTINUE, Continued, unicode, type RecordInfo } from './records.js'
import type { Sheet, Workbook } from './workbook.js'

interface Note {
  row: number
  col: number
  visible: boolean
  id: number
  author: string
}

export function notes(source: Generator<RecordInfo>, workbook: Workbook, sheet: Sheet, xml: Xml, vml: Xml): boolean {
  const notes: Note[] = []
  const text = new Map<number, { characters: string; runs: [number, number][] }>()
  const encoding = workbook.encoding
  let old: { row: number; col: number; left: number; parts: Uint8Array[] } | undefined
  const finishOld = () => {
    if (!old || !encoding) return
    const { row, col, left, parts } = old
    const id = notes.length
    const bytes = combinedPayload(parts, 'Workbook note', 0x001c)
    const characters = encoding.text(bytes, bytes.remaining)
    notes.push({ row, col, visible: false, id, author: '' })
    text.set(id, { characters, runs: [] })
    if (left) {
      workbook.losses.add('comment', 'the note text ends before its declared length', 'Текст примечания короче указанной длины',
        sheet.name, formatAddress(row, col))
    }
    old = undefined
  }
  let objectId: number | undefined
  let record = source.next().value
  while (record) {
    const { type, payload } = record
    record = undefined
    if (type !== 0x001c) finishOld()
    if (type === 0x001c && encoding) {
      const row = payload.u16()
      const col = payload.u16()
      const length = payload.u16()
      if (row !== 0xffff) {
        finishOld()
        if (col <= 255) old = { row, col, left: length, parts: [] }
      }
      if (old) {
        const count = Math.min(old.left, length, payload.remaining)
        old.parts.push(payload.bytes(count))
        old.left -= count
      }
    } else if (type === 0x005d && !encoding) {
      // FtCmo starts Obj. Object type 25 is a note.
      const structure = payload.u16()
      payload.skip(2)
      const kind = payload.u16()
      objectId = structure === 0x15 && kind === 25 ? payload.u16() : undefined
    } else if (type === 0x01b6 && objectId !== undefined) {
      payload.skip(10)
      const length = payload.u16()
      const runsSize = payload.u16()
      const runs: [number, number][] = []
      let characters = ''
      if (length) {
        const next = () => {
          record = source.next().value
          if (record?.type !== CONTINUE) return undefined
          const part = record.payload
          record = undefined
          return part
        }
        const head = next() ?? payload.fail('note text needs a Continue record')
        const continued = new Continued(head, next)
        characters = continued.text(length, Boolean(continued.u8() & 1))
        for (let i = 0; i + 8 < runsSize; i += 8) {
          runs.push([continued.u16(), continued.u16()])
          continued.skip(4)
        }
      }
      text.set(objectId, { characters, runs })
      objectId = undefined
    } else if (type === 0x001c) {
      const row = payload.u16()
      const col = payload.u16()
      const flags = payload.u16()
      const id = payload.u16()
      const author = unicode(payload)
      if (col <= 255) notes.push({ row, col, visible: Boolean(flags & 2), id, author })
    }
    record ??= source.next().value
  }
  finishOld()
  if (!notes.length) return false
  const authors = [...new Set(notes.map((note) => note.author))]
  const authorIds = new Map(authors.map((author, id) => [author, id]))
  xml.add('<comments><authors>')
  for (const author of authors) xml.add(`<author>${escapeText(author)}</author>`)
  xml.add('</authors><commentList>')
  vml.add('<xml>')
  for (const note of notes) {
    const at = formatAddress(note.row, note.col)
    const value = text.get(note.id)
    if (!value) {
      workbook.losses.add('comment', 'a note has no text object', 'У примечания отсутствует текстовый объект', sheet.name, at)
      continue
    }
    xml.add(`<comment ref="${at}" authorId="${authorIds.get(note.author)}"><text>`)
    const plain = (part: string) => `<t xml:space="preserve">${escapeText(part)}</t>`
    if (value.runs.length) {
      for (const [i, [start, font]] of value.runs.entries()) {
        const end = value.runs[i + 1]?.[0] ?? value.characters.length
        if (end > start && end <= value.characters.length) xml.add(workbook.styles.run(font, plain(value.characters.slice(start, end))))
      }
    } else xml.add(plain(value.characters))
    xml.add('</text></comment>')
    if (note.visible) vml.add(`<shape><ClientData ObjectType="Note"><Visible/><Row>${note.row}</Row><Column>${note.col}</Column></ClientData></shape>`)
  }
  xml.add('</commentList></comments>')
  vml.add('</xml>')
  return true
}

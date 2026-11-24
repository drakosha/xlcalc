import { textElement } from '../xlsx/xml/text.js'
import type { Part } from './parts.js'
import { records, type Payload } from './records.js'

export function richString(payload: Payload, fonts: readonly string[], pending: () => void): string {
  const flags = payload.u8()
  const value = payload.wide()
  const runs: { start: number; font: number }[] = []
  if (flags & 1) {
    const count = payload.u32()
    if (count > 32767 || count * 4 > payload.remaining) payload.fail(`invalid rich text run count ${count}`)
    for (let i = 0; i < count; i++) {
      const start = payload.u16()
      const font = payload.u16()
      if (start > value.length || (i && start <= runs[i - 1].start)) payload.fail(`invalid rich text run start ${start}`)
      runs.push({ start, font })
    }
  }
  if (flags & 2) {
    payload.wide()
    const count = payload.u32()
    if (count > 32767) payload.fail(`invalid phonetic run count ${count}`)
    payload.skip(count * 10)
  }
  payload.done()
  if (!runs.length) return textElement(value)
  let xml = runs[0].start ? `<r>${textElement(value.slice(0, runs[0].start))}</r>` : ''
  for (let i = 0; i < runs.length; i++) {
    const { start, font } = runs[i]
    if (start === value.length) continue
    const style = fonts[font]
    if (style === undefined) pending()
    xml += `<r>${style ? `<rPr>${style}</rPr>` : ''}${textElement(value.slice(start, runs[i + 1]?.start))}</r>`
  }
  return xml
}

export function sharedStrings({ bytes, path, xml, context, pending }: Part): void {
  xml.add('<sst>')
  for (const { type, payload } of records(bytes, path)) {
    if (type !== 19) continue
    xml.add(`<si>${richString(payload, context.fonts, pending)}</si>`)
    context.strings++
  }
  xml.add('</sst>')
}

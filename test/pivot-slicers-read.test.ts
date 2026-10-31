/**
 * Slicers and timelines read into the workbook: a book Excel for Mac saved with a PivotTable
 * (`test/excel/keep/made/pivot.xlsx`), its parts of a slicer and a timeline over that PivotTable written by hand as the
 * x14 and x15 schemas give them, every field of the model; and the books Excel for Windows saved for the set slicer,
 * against the items and dates Excel showed.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { PivotValue } from '../src/index.js'
import { dateToSerial } from '../src/model.js'
import { readXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { pack } from './xlsx-fixtures.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'

/** A drawing's anchor of a frame showing `tag` (`sle:slicer`, `tsle:timeslicer`) named `name`, from column and row to column and row. */
const anchor = (tag: string, ns: string, name: string, [c1, r1, c2, r2]: number[], off = 0) =>
  `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${c1}</xdr:col><xdr:colOff>${off}</xdr:colOff><xdr:row>${r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="a14"><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${ns}"><${tag} xmlns:${tag.split(':')[0]}="${ns}" name="${name}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`

async function book(): Promise<Uint8Array> {
  const data = readFileSync(new URL('./excel/keep/made/pivot.xlsx', import.meta.url))
  const a = openArchive(data)
  const parts: Record<string, string | undefined> = {}
  for (const n of a.names) parts[n] = await a.text(n)
  const wbRels = 'xl/_rels/workbook.xml.rels'
  parts[wbRels] = parts[wbRels]!.replace('</Relationships>', `<Relationship Id="rIdSC" Type="${'http://schemas.microsoft.com/office/2007/relationships'}/slicerCache" Target="slicerCaches/slicerCache1.xml"/><Relationship Id="rIdTC" Type="http://schemas.microsoft.com/office/2011/relationships/timelineCache" Target="timelineCaches/timelineCache1.xml"/></Relationships>`)
  const cache = 'xl/pivotCache/pivotCacheDefinition1.xml'
  parts[cache] = parts[cache]!.replace('<x14:pivotCacheDefinition/>', '<x14:pivotCacheDefinition pivotCacheId="5"/>')
  const tab = /<sheet name="Pivot" sheetId="(\d+)"/.exec(parts['xl/workbook.xml']!)![1]
  const pivots = `<pivotTables><pivotTable tabId="${tab}" name="PivotTable1"/></pivotTables>`
  parts['xl/slicerCaches/slicerCache1.xml'] = `<slicerCacheDefinition xmlns="${X14}" name="Slicer_Region" sourceName="Region">${pivots}<data><tabular pivotCacheId="5" sortOrder="descending" crossFilter="none" showMissing="0"><items count="4"><i x="0" s="1"/><i x="1"/><i x="2" s="1" nd="1"/><i x="3" s="1"/></items></tabular></data></slicerCacheDefinition>`
  parts['xl/timelineCaches/timelineCache1.xml'] = `<timelineCacheDefinition xmlns="${X15}" name="NativeTimeline_Month" sourceName="Month">${pivots}<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="5" filterType="dateBetween"><selection startDate="2024-01-01T00:00:00" endDate="2024-03-31T00:00:00"/><bounds startDate="2024-01-01T00:00:00" endDate="2025-01-01T00:00:00"/></state></timelineCacheDefinition>`
  parts['xl/slicers/slicer1.xml'] = `<slicers xmlns="${X14}"><slicer name="Region" cache="Slicer_Region" caption="Regions" columnCount="2" rowHeight="241300" style="SlicerStyleLight2" showCaption="0"/></slicers>`
  parts['xl/timelines/timeline1.xml'] = `<timelines xmlns="${X15}"><timeline name="Month" cache="NativeTimeline_Month" caption="Month" level="2" selectionLevel="1" scrollPosition="2024-01-01T00:00:00" showHorizontalScrollbar="0"/></timelines>`
  parts['xl/drawings/drawing9.xml'] =
    `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    anchor('sle:slicer', 'http://schemas.microsoft.com/office/drawing/2010/slicer', 'Region', [4, 2, 6, 12], 9525) +
    anchor('tsle:timeslicer', 'http://schemas.microsoft.com/office/drawing/2012/timeslicer', 'Month', [7, 2, 12, 9]) +
    `</xdr:wsDr>`
  const sheetRels = 'xl/worksheets/_rels/sheet2.xml.rels'
  parts[sheetRels] = parts[sheetRels]!.replace('</Relationships>', `<Relationship Id="rIdS" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/><Relationship Id="rIdT" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="../timelines/timeline1.xml"/><Relationship Id="rIdD" Type="${REL}/drawing" Target="../drawings/drawing9.xml"/></Relationships>`)
  return pack(parts)
}

describe('slicers and timelines read', () => {
  it('the caches, the slicer and the timeline with their places', async () => {
    const r = await readXlsx(await book())
    expect(r.book.slicerCaches).toEqual([
      { name: 'Slicer_Region', field: 'Region', pivots: [{ sheet: 'Pivot', name: 'PivotTable1' }], cache: 0, items: [{ item: 0 }, { item: 1, selected: false }, { item: 2, noData: true }, { item: 3 }], sort: 'descending', showMissing: false, noData: { mark: false } },
    ])
    expect(r.book.timelineCaches).toEqual([{ name: 'NativeTimeline_Month', field: 'Month', pivots: [{ sheet: 'Pivot', name: 'PivotTable1' }], cache: 0, selection: { start: '2024-01-01T00:00:00', end: '2024-03-31T00:00:00' }, bounds: { start: '2024-01-01T00:00:00', end: '2025-01-01T00:00:00' } }])
    const pivot = r.book.sheets.find((s) => s.name === 'Pivot')!
    expect(pivot.slicers).toEqual([{ name: 'Region', cache: 'Slicer_Region', caption: 'Regions', showCaption: false, style: 'SlicerStyleLight2', columns: 2, rowHeight: 241300, place: { from: 'E3', to: 'G13', fromOffset: [9525, 0] } }])
    expect(pivot.timelines).toEqual([{ name: 'Month', cache: 'NativeTimeline_Month', caption: 'Month', showScrollbar: false, scrollPosition: '2024-01-01T00:00:00', level: 'months', selectionLevel: 'quarters', place: { from: 'H3', to: 'M10' } }])
    expect(r.issues.filter((i) => i.code === 'pivotTable')).toEqual([])
  })
})

describe('slicers and timelines of the books Excel for Windows saved (the set slicer)', () => {
  it("each slicer's items chosen and with data, each timeline's dates and level, as Excel showed them", async () => {
    const dir = new URL('./excel/win/pivot/', import.meta.url)
    const answers = (JSON.parse(readFileSync(new URL('slicer.json', dir), 'utf8')) as { answers: Record<string, { phases: { slicers?: Record<string, unknown> }[] }> }).answers
    const LEVELS = ['years', 'quarters', 'months', 'days']
    const day = (iso: string) => dateToSerial(iso.slice(0, 10))
    let n = 0
    for (const [id, a] of Object.entries(answers)) {
      const shown = a.phases.at(-1)!.slicers ?? {}
      const r = await readXlsx(readFileSync(new URL(`slicer/${id}.xlsx`, dir)))
      expect(r.issues.filter((i) => i.code === 'pivotTable'), id).toEqual([])
      for (const [name, s] of Object.entries(shown)) {
        n++
        if (Array.isArray(s)) {
          const c = r.book.slicerCaches!.find((x) => x.name === name)!
          const items = r.book.pivotCaches!.find((x) => x.id === c.cache)!.fields.find((f) => f.name === c.field)!.items!
          const label = (v: PivotValue) => (v === null ? '(blank)' : typeof v === 'boolean' ? String(v).toUpperCase() : String(v))
          const got = c.items!.map((i) => [label(items[i.item]), i.selected !== false, !i.noData])
          expect([id, name, [...got].sort()]).toEqual([id, name, [...s].sort()])
          expect(r.book.sheets.flatMap((x) => x.slicers ?? []).some((x) => x.cache === name), `${id} ${name}`).toBe(true)
        } else {
          const [from, to, level] = (s as { timeline: [number, number, number] }).timeline
          const c = r.book.timelineCaches!.find((x) => x.name === name)!
          const t = r.book.sheets.flatMap((x) => x.timelines ?? []).find((x) => x.cache === name)!
          expect([id, day(c.selection!.start), day(c.selection!.end), t.level ?? 'months']).toEqual([id, from, to, LEVELS[level]])
        }
      }
    }
    expect(n).toBeGreaterThan(15)
  })
})

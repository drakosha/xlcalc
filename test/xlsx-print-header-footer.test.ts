import { describe, expect, it } from 'vitest'
import { formatPrintHeaderFooter } from '../src/xlsx-print.js'

const context = {
  pageNumber: 7, pageCount: 10, sheetName: 'Summary', fileName: 'report.xlsx', filePath: 'C:\\Reports\\',
  date: '24.11.2026', time: '15:10',
}

describe('Excel print header and footer codes', () => {
  it('expands fields with independent formatting in three sections', () => {
    const got = formatPrintHeaderFooter('&L&"Arial,Bold"&12Q && A &A&C&D &T&L&I italic&RPage &P-2 of &N+1 (&Z&F)', context)
    expect(got).toEqual({
      left: [
        { text: 'Q & A Summary', style: { name: 'Arial', size: 12, bold: true, italic: false } },
        { text: ' italic', style: { name: 'Arial', size: 12, bold: true, italic: true } },
      ],
      center: [{ text: '24.11.2026 15:10', style: {} }],
      right: [{ text: 'Page 5 of 11 (C:\\Reports\\report.xlsx)', style: {} }],
    })
  })

  it('keeps pictures, theme references and mutually exclusive script styles', () => {
    const got = formatPrintHeaderFooter('&L&Kffcc00Gold&K05.-050&XUp&YStillUp&X&YDown&Y&"+"&EUnder&G', context)
    expect(got.left).toEqual([
      { text: 'Gold', style: { color: '#FFCC00' } },
      { text: 'UpStillUp', style: { color: { theme: 5, tint: -50 }, script: 'superscript' } },
      { text: 'Down', style: { color: { theme: 5, tint: -50 }, script: 'subscript' } },
      { text: 'Under', style: { color: { theme: 5, tint: -50 }, themeFont: 'heading', underline: 'double' } },
      { picture: true },
    ])
  })

  it('matches Excel offsets with punctuation and the lone terminal digit', () => {
    const got = formatPrintHeaderFooter('&L&n+12 / &N-12 / &P- &Rpage &P+1', context)
    expect(got.left).toEqual([{ text: '22 / -2 / 7 ', style: {} }])
    expect(got.right).toEqual([{ text: 'page 71', style: {} }])
    expect(formatPrintHeaderFooter('&P+12', context).center).toEqual([{ text: '82', style: {} }])
    expect(formatPrintHeaderFooter('&P-12', context).center).toEqual([{ text: '62', style: {} }])
  })

  it('keeps line breaks, literal signs and localized font faces', () => {
    const got = formatPrintHeaderFooter('Plain\n&Q&L&"Times New Roman,Курсив"&8X&P+ &S1&S&', context)
    expect(got.center).toEqual([{ text: 'Plain\n', style: {} }])
    expect(got.left).toEqual([
      { text: 'X7 ', style: { name: 'Times New Roman', face: 'Курсив', size: 8 } },
      { text: '1', style: { name: 'Times New Roman', face: 'Курсив', size: 8, strike: true } },
    ])
  })
})

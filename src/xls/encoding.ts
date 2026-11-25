import type { Payload } from '../xlsb/records.js'

interface Decoder {
  decode(data: Uint8Array): string
}

const encodings: Readonly<Partial<Record<number, string>>> = {
  866: 'ibm866', 874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 10000: 'macintosh',
  10007: 'x-mac-cyrillic', 28591: 'windows-1252', 32768: 'macintosh',
  32769: 'windows-1252', 65001: 'utf-8',
}

// High bytes of the Microsoft OEM mappings omitted by the WHATWG TextDecoder.
// https://www.unicode.org/Public/MAPPINGS/VENDORS/MICSFT/PC/
const oem: Readonly<Partial<Record<number, string>>> = {
  437: "ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒ" +
    "áíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐" +
    "└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀" +
    "αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■\u00a0",
  850: "ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜø£Ø×ƒ" +
    "áíóúñÑªº¿®¬½¼¡«»░▒▓│┤ÁÂÀ©╣║╗╝¢¥┐" +
    "└┴┬├─┼ãÃ╚╔╩╦╠═╬¤ðÐÊËÈıÍÎÏ┘┌█▄¦Ì▀" +
    "ÓßÔÒõÕµþÞÚÛÙýÝ¯´\u00ad±‗¾¶§÷¸°¨·¹³²■\u00a0",
  852: "ÇüéâäůćçłëŐőîŹÄĆÉĹĺôöĽľŚśÖÜŤťŁ×č" +
    "áíóúĄąŽžĘę¬źČş«»░▒▓│┤ÁÂĚŞ╣║╗╝Żż┐" +
    "└┴┬├─┼Ăă╚╔╩╦╠═╬¤đĐĎËďŇÍÎě┘┌█▄ŢŮ▀" +
    "ÓßÔŃńňŠšŔÚŕŰýÝţ´\u00ad˝˛ˇ˘§÷¸°¨˙űŘř■\u00a0",
}

export type LegacyVersion = 2 | 3 | 4 | 5

export class LegacyEncoding {
  private decoder?: Decoder
  private unsupported?: number

  constructor(readonly version: LegacyVersion) {}

  set(page: number): void {
    const high = oem[page]
    if (high !== undefined) {
      this.decoder = {
        decode(data) {
          let text = ''
          for (const byte of data) text += byte < 128 ? String.fromCharCode(byte) : high[byte - 128]
          return text
        },
      }
      this.unsupported = undefined
      return
    }
    const label = page >= 1250 && page <= 1258 ? 'windows-' + page : encodings[page]
    const TextDecoder = (globalThis as typeof globalThis & { TextDecoder?: new (label: string) => Decoder }).TextDecoder
    if (label && TextDecoder) {
      this.decoder = new TextDecoder(label)
      this.unsupported = undefined
    } else this.unsupported = page
  }

  text(payload: Payload, length: number): string {
    if (this.unsupported !== undefined) {
      return payload.fail(`unsupported BIFF${this.version} code page ${this.unsupported}`)
    }
    if (!this.decoder) this.set(1252)
    return this.decoder ? this.decoder.decode(payload.bytes(length)) : payload.narrow(length)
  }
}

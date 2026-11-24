/** The Excel paper-size table (MS-OI29500 §2.1.638), with the Japanese envelope sizes from Windows PageMediaSize. */
import { checkNumber } from '../../input/basic.js'
import type { PrintSetup, PrintSize } from './types.js'

type Size = readonly [width: number, height: number]
const INCHES: Readonly<Record<number, Size>> = {
  1: [8.5, 11], 2: [8.5, 11], 3: [11, 17], 4: [17, 11], 5: [8.5, 14], 6: [5.5, 8.5], 7: [7.25, 10.5],
  14: [8.5, 13], 16: [10, 14], 17: [11, 17], 18: [8.5, 11],
  19: [3.875, 8.875], 20: [4.125, 9.5], 21: [4.5, 10.375], 22: [4.75, 11], 23: [5, 11.5],
  24: [17, 22], 25: [22, 34], 26: [34, 44], 37: [3.875, 7.5], 38: [3.625, 6.5],
  39: [14.875, 11], 40: [8.5, 12], 41: [8.5, 13], 44: [9, 11], 45: [10, 11], 46: [15, 11],
  50: [9.275, 12], 51: [9.275, 15], 52: [11.69, 18], 53: [9.27, 12.69], 54: [8.275, 11],
  56: [9.275, 12], 59: [8.5, 12.69], 75: [11, 8.5], 90: [12, 11],
}

const MM: Readonly<Record<number, Size>> = {
  8: [297, 420], 9: [210, 297], 10: [210, 297], 11: [148, 210], 12: [250, 354], 13: [182, 257], 15: [215, 275],
  27: [110, 220], 28: [162, 229], 29: [324, 458], 30: [229, 324], 31: [114, 162], 32: [114, 229],
  33: [250, 353], 34: [176, 250], 35: [176, 125], 36: [110, 230], 42: [250, 353], 43: [100, 148], 47: [220, 220],
  55: [210, 297], 57: [227, 356], 58: [305, 487], 60: [210, 330], 61: [148, 210], 62: [182, 257],
  63: [322, 445], 64: [174, 235], 65: [201, 276], 66: [420, 594], 67: [297, 420], 68: [322, 445],
  69: [200, 148], 70: [105, 148], 71: [240, 332], 72: [216, 277], 73: [120, 235], 74: [90, 205],
  76: [420, 297], 77: [297, 210], 78: [210, 148], 79: [364, 257], 80: [257, 182], 81: [148, 100],
  82: [148, 200], 83: [148, 105], 84: [332, 240], 85: [277, 216], 86: [235, 120], 87: [205, 90],
  88: [128, 182], 89: [182, 128], 91: [105, 235], 92: [235, 105], 93: [146, 215], 94: [97, 151], 95: [97, 151],
  96: [102, 165], 97: [102, 176], 98: [125, 176], 99: [110, 208], 100: [110, 220], 101: [120, 230],
  102: [160, 230], 103: [120, 309], 104: [229, 324], 105: [324, 458], 106: [215, 146], 107: [151, 97],
  108: [151, 97], 109: [165, 102], 110: [176, 102], 111: [176, 125], 112: [208, 110], 113: [220, 110],
  114: [230, 120], 115: [230, 160], 116: [309, 120], 117: [324, 229], 118: [458, 324],
}

function stock(code: number): PrintSize | undefined {
  const inches = INCHES[code]
  if (inches) return { width: inches[0] * 72, height: inches[1] * 72 }
  const metric = MM[code]
  if (metric) return { width: metric[0] * 72 / 25.4, height: metric[1] * 72 / 25.4 }
}

/** The named paper a sheet asks for, or points from the host's printer selection for unspecified and custom sizes. */
export function printPaperSize(setup: PrintSetup | undefined, fallback?: PrintSize): PrintSize | undefined {
  const paper = stock(setup?.paperSize ?? 0) ?? fallback
  if (!paper) return undefined
  const { width, height } = paper
  checkNumber(width, 'paper.width', Number.MIN_VALUE, Number.MAX_SAFE_INTEGER)
  checkNumber(height, 'paper.height', Number.MIN_VALUE, Number.MAX_SAFE_INTEGER)
  const orientation = setup?.orientation
  const swap = orientation === 'landscape' ? width < height : orientation === 'portrait' && width > height
  return swap ? { width: height, height: width } : { width, height }
}

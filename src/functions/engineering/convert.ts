/** CONVERT with Excel's units, aliases, prefixes and factors. This engine's own. */
import { E, type V } from '../../values.js'
import type { FnDef } from '../../args.js'

interface Unit {
  /** The quantity: units convert only within one. */
  q: string
  /** Value in the quantity's base unit (g, m, sec, Pa, N, J, W, T, m3, m2, bit, m/s; temperatures apart). */
  f: number
  /** Whether a decimal prefix applies, and to which power (m2 takes it squared). */
  p: 0 | 1 | 2 | 3
  /** Whether a binary prefix (ki … Yi) applies. */
  bin?: boolean
}

const IN = 0.0254
const FT = 0.3048
const LBF = 4.4482216152605
const PICA = IN / 72
const TSP = 4.92892159375e-6
const LY = 9460730472580800
const MI = 1609.344
const NMI = 1852

/** [quantity, prefix power, factor, symbols]; a power of 0 takes no prefix. */
const TABLE: [string, 0 | 1 | 2 | 3, number, string[]][] = [
  ['mass', 1, 1, ['g']],
  ['mass', 0, (LBF / FT) * 1000, ['sg']],
  ['mass', 0, 453.59237, ['lbm']],
  ['mass', 1, 1.660538782e-24, ['u']],
  ['mass', 0, 28.349523125, ['ozm']],
  ['mass', 0, 0.06479891, ['grain']],
  ['mass', 0, 45359.237, ['cwt', 'shweight']],
  ['mass', 0, 50802.34544, ['uk_cwt', 'lcwt', 'hweight']],
  ['mass', 0, 6350.29318, ['stone']],
  ['mass', 0, 907184.74, ['ton']],
  ['mass', 0, 1016046.9088, ['uk_ton', 'LTON', 'brton']],
  ['length', 1, 1, ['m']],
  ['length', 0, MI, ['mi']],
  ['length', 0, NMI, ['Nmi']],
  ['length', 0, IN, ['in']],
  ['length', 0, FT, ['ft']],
  ['length', 0, 0.9144, ['yd']],
  ['length', 1, 1e-10, ['ang']],
  ['length', 0, 1.143, ['ell']],
  ['length', 1, LY, ['ly']],
  ['length', 1, 30856775812815532, ['parsec', 'pc']],
  ['length', 0, PICA, ['Picapt', 'Pica']],
  ['length', 0, IN / 6, ['pica']],
  ['length', 0, (5280 * 1200) / 3937, ['survey_mi']],
  ['time', 0, 31557600, ['yr']],
  ['time', 0, 86400, ['day', 'd']],
  ['time', 0, 3600, ['hr']],
  ['time', 0, 60, ['mn', 'min']],
  ['time', 1, 1, ['sec', 's']],
  ['pressure', 1, 1, ['Pa', 'p']],
  ['pressure', 1, 101325, ['atm', 'at']],
  ['pressure', 1, 133.322, ['mmHg']],
  ['pressure', 0, LBF / (IN * IN), ['psi']],
  ['pressure', 0, 101325 / 760, ['Torr']],
  ['force', 1, 1, ['N']],
  ['force', 1, 1e-5, ['dyn', 'dy']],
  ['force', 0, LBF, ['lbf']],
  ['force', 1, 0.00980665, ['pond']],
  ['energy', 1, 1, ['J']],
  ['energy', 1, 1e-7, ['e']],
  ['energy', 1, 4.184, ['c']],
  ['energy', 1, 4.1868, ['cal']],
  ['energy', 1, 1.602176487e-19, ['eV', 'ev']],
  ['energy', 0, 550 * LBF * FT * 3600, ['HPh', 'hh']],
  ['energy', 1, 3600, ['Wh', 'wh']],
  ['energy', 0, LBF * FT, ['flb']],
  ['energy', 0, 1055.05585262, ['BTU', 'btu']],
  ['power', 0, 550 * LBF * FT, ['HP', 'h']],
  ['power', 0, 735.49875, ['PS']],
  ['power', 1, 1, ['W', 'w']],
  ['magnetism', 1, 1, ['T']],
  ['magnetism', 1, 1e-4, ['ga']],
  ['volume', 0, TSP, ['tsp']],
  ['volume', 0, 5e-6, ['tspm']],
  ['volume', 0, TSP * 3, ['tbs']],
  ['volume', 0, TSP * 6, ['oz']],
  ['volume', 0, TSP * 48, ['cup']],
  ['volume', 0, TSP * 96, ['pt', 'us_pt']],
  // Excel takes a prefix on the U.K. pint (CONVERT(1,"kuk_pt","l") is 568.26125)
  ['volume', 1, 0.00056826125, ['uk_pt']],
  ['volume', 0, TSP * 192, ['qt']],
  ['volume', 0, 0.0011365225, ['uk_qt']],
  ['volume', 0, TSP * 768, ['gal']],
  ['volume', 0, 0.00454609, ['uk_gal']],
  ['volume', 1, 0.001, ['l', 'L', 'lt']],
  ['volume', 3, 1e-30, ['ang3', 'ang^3']],
  ['volume', 0, 0.158987294928, ['barrel']],
  ['volume', 0, 0.03523907016688, ['bushel']],
  ['volume', 0, FT ** 3, ['ft3', 'ft^3']],
  ['volume', 0, IN ** 3, ['in3', 'in^3']],
  ['volume', 0, LY ** 3, ['ly3', 'ly^3']],
  ['volume', 3, 1, ['m3', 'm^3']],
  ['volume', 0, MI ** 3, ['mi3', 'mi^3']],
  ['volume', 0, 0.9144 ** 3, ['yd3', 'yd^3']],
  ['volume', 0, NMI ** 3, ['Nmi3', 'Nmi^3']],
  ['volume', 0, PICA ** 3, ['Picapt3', 'Picapt^3', 'Pica3', 'Pica^3']],
  ['volume', 0, 100 * FT ** 3, ['GRT', 'regton']],
  ['volume', 0, 40 * FT ** 3, ['MTON']],
  ['area', 0, 4046.8564224, ['uk_acre']],
  ['area', 0, 160 * ((5280 * 1200) / 3937 / 320) ** 2, ['us_acre']],
  ['area', 2, 1e-20, ['ang2', 'ang^2']],
  ['area', 1, 100, ['ar']],
  ['area', 0, FT ** 2, ['ft2', 'ft^2']],
  ['area', 0, 10000, ['ha']],
  ['area', 0, IN ** 2, ['in2', 'in^2']],
  ['area', 0, LY ** 2, ['ly2', 'ly^2']],
  ['area', 2, 1, ['m2', 'm^2']],
  ['area', 0, 2500, ['Morgen']],
  ['area', 0, MI ** 2, ['mi2', 'mi^2']],
  ['area', 0, NMI ** 2, ['Nmi2', 'Nmi^2']],
  ['area', 0, PICA ** 2, ['Picapt2', 'Pica2', 'Pica^2', 'Picapt^2']],
  ['area', 0, 0.9144 ** 2, ['yd2', 'yd^2']],
  ['information', 1, 1, ['bit']],
  ['information', 1, 8, ['byte']],
  ['speed', 0, (6080 * FT) / 3600, ['admkn']],
  ['speed', 0, NMI / 3600, ['kn']],
  ['speed', 1, 1 / 3600, ['m/h', 'm/hr']],
  ['speed', 1, 1, ['m/s', 'm/sec']],
  // Excel takes a prefix on miles per hour (CONVERT(1,"kmph","m/s") is 447.04)
  ['speed', 1, 0.44704, ['mph']],
  // temperatures: the factor is unused, see KELVIN
  ['temperature', 1, 1, ['K', 'kel']],
  ['temperature', 0, 1, ['C', 'cel', 'F', 'fah', 'Rank', 'Reau']],
]

const UNITS = new Map<string, Unit>()
for (const [q, p, f, symbols] of TABLE) for (const s of symbols) UNITS.set(s, { q, f, p, bin: q === 'information' })

/** Temperatures to and from kelvin. */
const KELVIN: Record<string, [(x: number) => number, (k: number) => number]> = {
  K: [(x) => x, (k) => k],
  C: [(x) => x + 273.15, (k) => k - 273.15],
  F: [(x) => ((x - 32) * 5) / 9 + 273.15, (k) => ((k - 273.15) * 9) / 5 + 32],
  Rank: [(x) => (x * 5) / 9, (k) => (k * 9) / 5],
  Reau: [(x) => (x * 5) / 4 + 273.15, (k) => ((k - 273.15) * 4) / 5],
}
const TEMP_ALIAS: Record<string, string> = { kel: 'K', cel: 'C', fah: 'F' }

const DECIMAL: Record<string, number> = {
  Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 1e1, e: 1e1,
  d: 1e-1, c: 1e-2, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24,
}
const BINARY: Record<string, number> = { Yi: 2 ** 80, Zi: 2 ** 70, Ei: 2 ** 60, Pi: 2 ** 50, Ti: 2 ** 40, Gi: 2 ** 30, Mi: 2 ** 20, ki: 2 ** 10 }

interface Found {
  unit: Unit
  symbol: string
  /** The prefix multiplier, raised to the unit's power. */
  scale: number
}

/** A unit symbol, exactly or after a binary or decimal prefix (`da` before `d`). */
function find(text: string): Found | null {
  const exact = UNITS.get(text)
  if (exact) return { unit: exact, symbol: text, scale: 1 }
  const bin = BINARY[text.slice(0, 2)]
  const afterBin = UNITS.get(text.slice(2))
  if (bin && afterBin?.bin) return { unit: afterBin, symbol: text.slice(2), scale: bin }
  for (const len of [2, 1]) {
    const mult = DECIMAL[text.slice(0, len)]
    const unit = UNITS.get(text.slice(len))
    if (mult && unit?.p) return { unit, symbol: text.slice(len), scale: mult ** unit.p }
  }
  return null
}

function convert(x: number, from: V, to: V): V {
  if (typeof from === 'boolean' || typeof to === 'boolean') return E.VALUE
  if (typeof from !== 'string' || typeof to !== 'string') return E.NA
  const a = find(from)
  const b = find(to)
  if (!a || !b || a.unit.q !== b.unit.q) return E.NA
  if (a.unit.q === 'temperature') {
    const k = KELVIN[TEMP_ALIAS[a.symbol] ?? a.symbol][0](x * a.scale)
    return KELVIN[TEMP_ALIAS[b.symbol] ?? b.symbol][1](k) / b.scale
  }
  return (x * a.unit.f * a.scale) / (b.unit.f * b.scale)
}

export const functions: Record<string, FnDef> = {
  CONVERT: { args: 'x v v', fn: convert },
}

import { lowerBound, upperBound } from '../sorted.js'

/** Positions in a stripe count the visible lines. Its borders still use worksheet coordinates. */
export interface StripeLines {
  before(index: number): number
  at(ordinal: number): number
}

export const allLines: StripeLines = { before: (index) => index, at: (ordinal) => ordinal }

export function visibleLines(spans: [number, number][]): StripeLines {
  if (!spans.length) return allLines
  const hidden = [0]
  const starts: number[] = []
  for (const [lo, hi] of spans) {
    const skipped = hidden.at(-1)!
    starts.push(lo - skipped)
    hidden.push(skipped + hi - lo + 1)
  }
  return {
    before(index) {
      const i = lowerBound(spans, index, ([, end]) => end)
      const span = spans[i]
      const partial = span ? Math.max(0, index - span[0]) : 0
      return index - hidden[i] - partial
    },
    at(ordinal) {
      return ordinal + hidden[upperBound(starts, ordinal)]
    },
  }
}

/** Neumaier summation, shared by statistical moments and cached range means. */
export class CompensatedSum {
  constructor(private sum = 0, private correction = 0) {}

  add(x: number): void {
    const next = this.sum + x
    this.correction += Math.abs(this.sum) >= Math.abs(x) ? (this.sum - next) + x : (x - next) + this.sum
    this.sum = next
  }

  get value(): number { return this.sum + this.correction }
  copy(): CompensatedSum { return new CompensatedSum(this.sum, this.correction) }
}

/** Undefined asks the caller for its exact-sum fallback after an intermediate overflow. */
export function meanOf(sum: CompensatedSum, count: number, first: number, constant: boolean): number | undefined {
  if (constant) return first
  const value = sum.value
  return Number.isFinite(value) ? value / count : undefined
}

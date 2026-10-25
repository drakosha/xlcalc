// Percentiles over samples, in one place for the benches and the test fixtures that measure times: the median and the
// 95th percentile as the benchmark reports take them, and `at`, the quantile by its index in the samples, which the
// pair scripts take their p50 and p95 by (another index rule than `p95`'s: they differ where 0.95n is a whole number).
export const median = (values) => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)]
export const p95 = (values) => [...values].sort((a, b) => a - b)[Math.ceil(values.length * 0.95) - 1]
export const at = (xs, q) => [...xs].sort((a, b) => a - b)[Math.min(xs.length - 1, Math.floor(xs.length * q))]

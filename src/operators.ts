/**
 * How a value compares with the bounds, by the order of Excel's comparison operators (`a`, `b`: the value against each
 * bound), as a conditional format's cellIs rule and a data validation rule of a compared type check it.
 */
export const COMPARE: Record<string, (a: number, b: number) => boolean> = {
  // bounds in the wrong order take nothing (the set `entry-held`, column C), where a conditional format's take both ways
  between: (a, b) => a >= 0 && b <= 0,
  notBetween: (a, b) => a < 0 || b > 0,
  equal: (a) => a === 0,
  notEqual: (a) => a !== 0,
  greaterThan: (a) => a > 0,
  lessThan: (a) => a < 0,
  greaterThanOrEqual: (a) => a >= 0,
  lessThanOrEqual: (a) => a <= 0,
}

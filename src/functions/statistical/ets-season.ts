/**
 * The season FORECAST.ETS and its kin detect where seasonality is 1 (their default) and FORECAST.ETS.SEASONALITY
 * reports. Excel's detection is not reproduced yet: `detectSeason` answers `UNDETECTED`, and the functions answer #N/A
 * there (`docs/limitations.md`).
 */

/** What `detectSeason` answers while the detection is not reproduced. */
export const UNDETECTED = -1

/** The length of the season of a series, 0 for none. */
export function detectSeason(_y: readonly number[]): number {
  return UNDETECTED
}

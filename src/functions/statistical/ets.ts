/**
 * FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY and FORECAST.ETS.STAT: the series read off its timeline
 * as Excel reads it (`ets-series.ts`), fitted by the model of `ets-model.ts`.
 */
import type { FnDef, Grid } from '../../args.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import { sumSquares } from '../common.js'
import { etsFit, etsForecast, etsSpread, type EtsFit } from './ets-model.js'
import { detectSeason, MAX_SEASON } from './ets-season.js'
import { horizonOf, seriesOf, type Series } from './ets-series.js'
import { normInv } from './special.js'

const whole = (x: number) => x === Math.trunc(x)

interface Model {
  s: Series
  m: number
  fit: EtsFit
}

/** The series and its model by the arguments shared by the four functions. */
function modelOf(values: Grid, timeline: Grid, season = 1, completion = 1, aggregation = 1): Model | CErr {
  if (completion !== 0 && completion !== 1) return E.NUM
  if (!(aggregation >= 1 && aggregation <= 7) || !whole(aggregation)) return E.NUM
  if (!(season >= 0 && season <= MAX_SEASON) || !whole(season)) return E.NUM
  const s = seriesOf(values, timeline, completion === 1, aggregation)
  if (isErr(s)) return s
  // a season given is none from 70 % of the series' points on
  const m = season === 1 ? detectSeason(s.y) : season < Math.trunc((70 * s.y.length) / 100) ? season : 0
  return { s, m, fit: etsFit(s.y, m) }
}

function forecastEts(target: number, values: Grid, timeline: Grid, season?: number, completion?: number, aggregation?: number): V {
  const md = modelOf(values, timeline, season, completion, aggregation)
  if (isErr(md)) return md
  const h = horizonOf(md.s, target)
  return isErr(h) ? h : etsForecast(md.fit, md.s.y.length, h)
}

/** The half width of the prediction interval at a level, from the one-step errors of the whole series. */
function confint(target: number, values: Grid, timeline: Grid, level = 0.95, season?: number, completion?: number, aggregation?: number): V {
  if (!(level >= 0 && level < 1)) return E.NUM
  const md = modelOf(values, timeline, season, completion, aggregation)
  if (isErr(md)) return md
  const h = horizonOf(md.s, target)
  if (isErr(h)) return h
  if (h <= 0) return E.NUM
  const { fit } = md
  const mse = sumSquares(fit.errors) / fit.errors.length
  return normInv((1 + level) / 2) * Math.sqrt(mse * etsSpread(fit, md.m, h))
}

function seasonality(values: Grid, timeline: Grid, completion?: number, aggregation?: number): V {
  const md = modelOf(values, timeline, 1, completion, aggregation)
  return isErr(md) ? md : md.m
}

function stat(values: Grid, timeline: Grid, type: number, season?: number, completion?: number, aggregation?: number): V {
  if (!(type >= 1 && type <= 8) || !whole(type)) return E.NUM
  const md = modelOf(values, timeline, season, completion, aggregation)
  if (isErr(md)) return md
  const { fit } = md
  return [fit.alpha, fit.beta, fit.gamma, fit.mase, fit.smape, fit.mae, fit.rmse, md.s.step][type - 1]
}

export const functions: Record<string, FnDef> = {
  'FORECAST.ETS': { args: 'n C C n? n? n?', fn: forecastEts },
  'FORECAST.ETS.CONFINT': { args: 'n C C n? n? n? n?', fn: confint },
  'FORECAST.ETS.SEASONALITY': { args: 'C C n? n?', fn: seasonality },
  'FORECAST.ETS.STAT': { args: 'C C n n? n? n?', fn: stat },
}

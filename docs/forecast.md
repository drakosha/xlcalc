# Forecasting with exponential smoothing

`FORECAST.ETS` forecasts a value of a timeline from the values before it, as Excel does: a level, a trend and a season
learnt from the series, each weighted toward the latest points. `FORECAST.LINEAR` (and `FORECAST`) fits one straight line
through all the points and reads the target off it; `FORECAST.ETS` follows the series as it changes and repeats its
season, so the two agree only on an exact line.

```js
compute({ sheets: [{ name: 'S', cells: {
  A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, A6: 6, A7: 7, A8: 8,
  B1: 12, B2: 15, B3: 11, B4: 14, B5: 18, B6: 13, B7: 16, B8: 19,
  C1: { formula: 'FORECAST.ETS(9,B1:B8,A1:A8)' },
  C2: { formula: 'FORECAST.ETS.CONFINT(9,B1:B8,A1:A8)' },
  C3: { formula: 'FORECAST.ETS.SEASONALITY(B1:B8,A1:A8)' },
} }] })
```

The four functions are in the group `statistics` ([Groups of functions](function-groups.md)).

## The functions

- `FORECAST.ETS(target_date, values, timeline, [seasonality], [data_completion], [aggregation])`: the forecast at
  `target_date`.
- `FORECAST.ETS.CONFINT(target_date, values, timeline, [confidence_level], [seasonality], [data_completion],
  [aggregation])`: the half width of the interval around that forecast, at a level of 0.95 by default (`0 <= level < 1`).
  The interval widens with the distance from the last point.
- `FORECAST.ETS.SEASONALITY(values, timeline, [data_completion], [aggregation])`: the length of the season the functions
  detect, 0 for none.
- `FORECAST.ETS.STAT(values, timeline, statistic_type, [seasonality], [data_completion], [aggregation])`: a statistic of
  the model: 1 alpha, 2 beta, 3 gamma (the smoothing of the level, the trend and the season), 4 MASE, 5 SMAPE, 6 MAE,
  7 RMSE (the accuracy of the one-step forecasts over the last points of the series), 8 the step of the timeline.

## The timeline

- `values` and `timeline` have as many cells as each other (`#N/A` otherwise). A point whose time is not a number is
  left out; an error in either is the answer.
- The step is the most frequent difference between successive times; two equally frequent steps are `#NUM!`. Times off
  the step's grid go to the nearest slot counted back from the last point, a point halfway between two to the later one.
- Dates on one day of every month, or of every third month, make a calendar timeline (a day a short month lacks is its
  last day): its targets are counted in months, so the first of the next month is one step ahead whatever the month's
  length, and a date between two steps lies on the line between them.
- `aggregation` combines the values of one slot: 1 `AVERAGE` (the default), 2 `COUNT`, 3 `COUNTA` (text counts), 4 `MAX`,
  5 `MEDIAN`, 6 `MIN`, 7 `SUM`.
- A slot without a value is completed by the line between its neighbours (`data_completion` 1, the default) or taken as
  0 (`data_completion` 0). Text, a boolean and a blank among the values count as missing. More than 30 % of the slots
  missing is `#NUM!`. When completing, missing values at either end are no part of the series.
- At least two points are needed: one is `#DIV/0!`, two at one time `#NUM!`.

## The season

- `seasonality` 1, the default, detects the season; 0 means none; 2 to 8,784 give its length. A season of 70 % of the
  series' points or more is taken as none (`FORECAST.ETS(11,B1:B10,A1:A10,7)` forecasts without one). Other values are
  `#NUM!`, fractions too.
- The detection looks for a period in the series' spectrum and in its autocorrelation, so a pattern that repeats over
  several periods is found, and noise usually gives 0.

## Targets

- A target is any time from the series' first step on, before the last point too. A fraction of a step lies on the line
  between the two whole steps around it; a target within 0.001 of a whole step is that step. Before the series it is
  `#NUM!`, and the interval of a target at or before the last point too.

## Checked against Excel

The functions match Excel 16.0.20430 for Windows and Excel 16.114 for Mac on 4,037 cases of five sets, two of them held
out and predicted before Excel ran them: the parameters and the statistics to the bit, the forecasts and intervals to
their last digits. The step of quarterly dates from January is a known difference ([Limitations](limitations.md)).

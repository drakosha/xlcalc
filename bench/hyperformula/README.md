# xlcalc and HyperFormula speed comparison

This optional benchmark compares the calculation engines on deterministic synthetic books. It adds no dependency to the xlcalc package and changes no engine implementation.

## Run

From the repository root:

```sh
npm run build
npm ci --prefix bench/hyperformula --ignore-scripts
node --test bench/hyperformula/shared.spec.mjs
node bench/hyperformula/run.mjs
```

Select workloads or engines, or keep output elsewhere:

```sh
node bench/hyperformula/run.mjs --case prefix-sum,lookup --out /tmp/xlcalc-hf-results
node bench/hyperformula/run.mjs --case sparse --engine hf-sparse,hf-auto
```

The output is `results/raw.json` (samples, verification, configuration, machine and revision) and `results/summary.csv`. A timeout or memory limit is recorded separately; mismatches and other errors also cause exit code 1. The benchmark does not assert absolute timing thresholds.

## Protocol

- HyperFormula is pinned to 3.5.0. xlcalc is compiled from the checkout being measured. The recorded engine revision is `c3733f6`.
- Each workload/engine runs sequentially in its own Node process, with a 1024 MiB V8 heap limit and a 30-second limit on an individual phase. A stopped child cannot take down the runner.
- Opening means a fresh session, including parsing, compilation, graph construction and calculation. There is one warmup and five measured sessions in the same process: this measures fresh sessions with a warmed JavaScript runtime, not cold process startup.
- Each editing workload has two warmup edits and twenty measured edits. Median, p95 and the first edit are recorded. Edits alternate between two states; both return shapes are included in the timed calls. A contiguous column is written to HF with one `setCellContents`, not thousands of scalar calls.
- Input generation and conversion, formatting, result extraction, hashes and verification are outside the timer. xlcalc uses `display: false`. Both wall time and process CPU time are recorded; CPU time can exceed wall time because V8 also works on background threads.
- Explicit GC precedes each fresh session. Retained heap is measured after GC with the final session held, above the shared inputs and the verifier's reference. It is not peak RSS, total application memory or a browser memory measurement.
- Every populated cell is compared after every fresh session and every edit. Formula values come from an independent full `compute`; literals come from the input and the edit overlay. Numeric tolerance is `1e-9 * max(1, abs(a), abs(b))`; strings, booleans, blanks and error codes match exactly. This checks agreement, not correctness against Excel. Mismatches do not qualify for a speed ranking.
- `HF_OPTIONS` aligns dates, parsing, locale and precision with the input contract. Text is explicitly escaped, date objects become Excel serials, and TRUE/FALSE constants are supplied as named expressions. Relative named references and Excel features unsupported by HF are not benchmarked.
- `hf-default` uses default performance options. The installed HF 3.5.0 defaults to `AlwaysDense`, confirmed from `Config.mjs` and the live configuration. `hf-index` enables `useColumnIndex` for the lookup book. `hf-sparse` and `hf-auto` separately test sparse mapping and automatic selection at threshold 0.5.
- The 34,000-row HF business case is a single bounded cold probe. It has no warmup or editing samples if opening cannot finish. Do not infer its output parity from a timeout.

The reference preparation runs xlcalc before the measurement, and verification runs between edits. Both can influence runtime warmup and GC. Samples were collected on an active host; load averages are saved. The recorded report combines the initial prefix/lookup run, the remaining workloads and the sparse-policy controls, each without concurrent benchmark children.

## Recorded results

Node 26.7.0, Apple M5 Pro, 48 GiB RAM. Times are milliseconds, medians. xlcalc below is `closures`; `eval`, edit p95, first edits and CPU samples are in the raw report.

| Workload | xlcalc open | HF open | Edit | xlcalc edit | HF edit |
|---|---:|---:|---|---:|---:|
| SUMIFS business book, 3,400 data rows | 180 | 6,062 | one criterion | 0.28 | 2.97 |
| Same book | 180 | 6,062 | one data amount | 25.6 | 4,523 |
| 5,000 cumulative SUMs | 34.2 | 46.0 | middle input | 6.63 | 4.19 |
| 5,000 cumulative AVERAGE + COUNTIF | 92.9 | 114 | middle input | 23.8 | 14.1 |
| 1,000 VLOOKUPs over 34,000 unsorted rows, HF default | 55.6 | 2,189 | one lookup key | 0.19 | 1.20 |
| Same book, HF column index | 55.6 | 1,270 | one lookup key | 0.19 | 0.12 |
| Same book, HF column index | 55.6 | 1,270 | one table value | 3.20 | 2.64 |
| 10,000-cell dependency chain | 62.0 | 47.5 | root input | 15.1 | 7.29 |
| 50,000 direct dependents and one SUM | 248 | 527 | root input | 95.1 | 59.3 |
| 2,000 inputs replaced atomically | 27.7 | 42.9 | all inputs | 22.0 | 4.34 |
| 2,000 cells in a roughly 97,000 x 311 sheet, HF automatic mapping | 9.74 | 23.7 | one input | 0.116 | 0.035 |
| SUMIFS business book, 34,000 data rows | 1,010 | >30,000 (probe timeout) | one data amount | 153 | not measured |

The automatic HF mapping selected sparse storage. On that sparse book, retained heap was 1.5 MiB for xlcalc, 235.9 MiB for HF's default dense mapping, and 0.9 MiB for HF's sparse or automatic mapping.

There were 30 workload/engine variants: 29 completed with matching values and one HF opening timeout. The successful variants made 80,622,748 cell comparisons, including repeated comparisons of the same cells; the largest relative numeric difference was below `5e-15`.

These results locate workload-dependent differences. They do not establish a universal winner, browser responsiveness, or Excel compatibility on arbitrary workbooks. The largest remaining editing differences in this set are the batch of 2,000 inputs, the dependency chain and the fanout; no profiler diagnosis or engine fix is part of this benchmark.

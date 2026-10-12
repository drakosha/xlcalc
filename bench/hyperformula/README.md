# xlcalc and HyperFormula speed comparison

This optional benchmark compares the calculation engines on deterministic synthetic books. It adds no dependency to the xlcalc package. The `perf/hf-parity` branch also contains three engine optimizations measured against master 4d15ecb (xlcalc 0.21.0).

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
node --expose-gc bench/hyperformula/profile.mjs chain xl-eval 160 /tmp/xlcalc-cpu-profiles
node bench/hyperformula/study.mjs
```

The output is `results/raw.json` (samples, verification, configuration, machine and revision) and `results/summary.csv`. A timeout or memory limit is recorded separately; mismatches and other errors also cause exit code 1. The benchmark does not assert absolute timing thresholds.

## Protocol

- HyperFormula is pinned to 3.5.0. xlcalc is compiled from the checkout being measured. The baseline engine is master `4d15ecb`, recorded at `364c4f2` after adding the same benchmark. The optimized engine is recorded at `b8d1ea0`. `xl-eval` uses `execution: 'eval'`, which compiles a batch with `new Function` and retains a function per formula.
- Each workload/engine runs sequentially in its own Node process, with a 1024 MiB V8 heap limit and a 30-second limit on an individual phase. A stopped child cannot take down the runner.
- Opening means a fresh session, including parsing, compilation, graph construction and calculation. There is one warmup and five measured sessions in the same process: this measures fresh sessions with a warmed JavaScript runtime, not cold process startup.
- Each editing workload has two warmup edits and twenty measured edits. Median, p95 and the first edit are recorded. Edits alternate between two states; both return shapes are included in the timed calls. A contiguous column is written to HF with one `setCellContents`, not thousands of scalar calls.
- Input generation and conversion, formatting, result extraction, hashes and verification are outside the timer. xlcalc uses `display: false`. Both wall time and process CPU time are recorded; CPU time can exceed wall time because V8 also works on background threads.
- Explicit GC precedes each fresh session. Retained heap is measured after GC with the final session held, above the shared inputs and the verifier's reference. It is not peak RSS, total application memory or a browser memory measurement.
- Every populated cell is compared after every fresh session and every edit. Formula values come from an independent full `compute`; literals come from the input and the edit overlay. Numeric tolerance is `1e-9 * max(1, abs(a), abs(b))`; strings, booleans, blanks and error codes match exactly. This checks agreement, not correctness against Excel. Mismatches do not qualify for a speed ranking.
- `HF_OPTIONS` aligns dates, parsing, locale and precision with the input contract. Text is explicitly escaped, date objects become Excel serials, and TRUE/FALSE constants are supplied as named expressions. Relative named references and Excel features unsupported by HF are not benchmarked.
- `hf-default` uses default performance options. The installed HF 3.5.0 defaults to `AlwaysDense`, confirmed from `Config.mjs` and the live configuration. `hf-index` enables `useColumnIndex` for the lookup book. `hf-sparse` and `hf-auto` separately test sparse mapping and automatic selection at threshold 0.5.
- The 34,000-row HF business case is a single bounded cold probe. It has no warmup or editing samples if opening cannot finish. Do not infer its output parity from a timeout.

The reference preparation runs xlcalc before the measurement, and verification runs between edits. Both can influence runtime warmup and GC. Samples were collected on an active host; load averages are saved. Baseline and optimized reports are separate complete runs of the same nine workloads, with no overlapping benchmark children or profiling processes. Their raw reports are in `results-baseline/` and `results/`.

`profile.mjs` warms thirty edits, profiles repeated edits alone with the [Node inspector CPU profiler](https://nodejs.org/api/inspector.html), then verifies every cell against a full calculation. Input preparation is included in its diagnostic profile; these durations are not the benchmark's editing samples. It writes a `.cpuprofile` and a JSON summary to the supplied directory.

## Recorded results

Node 26.7.0, Apple M5 Pro, 48 GiB RAM. Times are milliseconds, medians. Both execution modes are shown explicitly. Opening includes parsing and compilation; editing includes returned changes.

| Workload | xl closures open | xl eval open | HF default open |
|---|---:|---:|---:|
| SUMIFS business, 3,400 rows | 114 | 131 | 3843 |
| 5,000 cumulative SUMs | 33.14 | 37.80 | 45.88 |
| 5,000 cumulative AVERAGE + COUNTIF | 55.07 | 60.67 | 80.54 |
| 1,000 VLOOKUPs over 34,000 rows | 55.70 | 57.37 | 2205 |
| 10,000-cell dependency chain | 41.07 | 49.56 | 32.98 |
| 50,000 dependents and one SUM | 166 | 232 | 367 |
| 2,000 inputs replaced atomically | 18.70 | 20.65 | 28.55 |
| 2,000 sparse cells | 6.83 | 8.57 | 51.04 |
| SUMIFS business, 34,000 rows | 745 | 724 | >30,000 (timeout) |

HF with a column index opens the lookup book in 1239 ms. Automatic sparse mapping opens the sparse book in 19.64 ms.

| Workload and edit | xl closures | xl eval | HF default | HF tuned |
|---|---:|---:|---:|---:|
| SUMIFS business, 3,400 rows, criterion | 0.215 | 0.246 | 0.966 |  |
| SUMIFS business, 3,400 rows, data | 18.21 | 17.67 | 2901 |  |
| 5,000 cumulative SUMs, middle-input | 4.24 | 4.30 | 4.25 |  |
| 5,000 cumulative AVERAGE + COUNTIF, middle-input | 9.39 | 9.55 | 9.89 |  |
| 1,000 VLOOKUPs over 34,000 rows, lookup-key | 0.194 | 0.187 | 1.09 | 0.132 |
| 1,000 VLOOKUPs over 34,000 rows, lookup-data | 3.24 | 3.31 | 868 | 2.26 |
| 10,000-cell dependency chain, root | 5.23 | 5.04 | 5.47 |  |
| 50,000 dependents and one SUM, root | 41.34 | 42.32 | 44.54 |  |
| 2,000 inputs replaced atomically, 2000-inputs | 6.45 | 5.76 | 3.40 |  |
| 2,000 sparse cells, one-input | 0.084 | 0.091 | 0.026 | 0.084 |
| SUMIFS business, 34,000 rows, criterion | 0.467 | 0.470 | not measured |  |
| SUMIFS business, 34,000 rows, data | 104 | 103 | not measured |  |

The column index is the tuned lookup case; automatic mapping is the tuned sparse case. The sparse edit is below 0.1 ms and varies substantially between runs.

### Changes against the same benchmark on master

| Editing workload | closures before | closures after | eval before | eval after |
|---|---:|---:|---:|---:|
| 5,000 cumulative SUMs | 6.68 | 4.24 | 7.05 | 4.30 |
| 5,000 cumulative AVERAGE + COUNTIF | 16.07 | 9.39 | 15.67 | 9.55 |
| 10,000-cell dependency chain | 8.63 | 5.23 | 9.18 | 5.04 |
| 50,000 dependents and one SUM | 60.41 | 41.34 | 67.25 | 42.32 |
| 2,000 inputs replaced atomically | 14.33 | 6.45 | 14.57 | 5.76 |

The final paired run reaches median editing parity on prefixes, the chain and fanout. An earlier control still had eval at 5.01 ms versus HF at 4.21 ms on the chain, so the small remaining chain difference is not settled. The 2,000-input batch remains about 1.7 to 1.9 times slower; lookup edits remain about 1.4 to 1.5 times slower than indexed HF. Fresh chain construction also remains slower.

Business data edits did not improve: 3,400 rows changed from 17.35 to 18.21 ms in closures and 16.99 to 17.67 ms in eval; 34,000 rows changed from 100.82 to 104.05 ms and 99.44 to 103.41 ms. These roughly 3 to 5 percent differences are from an active host, and are a reason to review the cache tradeoff on application workbooks.

### Implementation and cost

- Consecutive column edits query their exact combined area once. Ascending and descending batches are covered; gaps and other columns stay separate. The listener probe test failed at 393,472 coordinate reads before the change and passes below its 16,384 limit.
- Internal cell keys are decoded with slices instead of split/map; edit bookkeeping creates sets only when needed, and books without spills skip spill lookups. Ordinary names uppercase directly; exceptional dotless i and long s preserve their existing treatment.
- Prefix checkpoints are 64 values apart instead of 256. Full folds still use the 256-cell threshold. Short ranges leave a mark on first use and retain a fold on repeated use. Their tests failed at 132,099 prefix reads and 5,120 shared short-range reads before the change, and pass afterward, including invalidation when the short range changes.

The controlled checkpoint comparison tried 256, 128, 64 and 32 while keeping the full-fold threshold at 256 and excluding short-range memoization. Eval AVERAGE + COUNTIF edits were 14.20, 10.70, 9.00 and 7.70 ms. Choosing 64 reaches the target with fewer checkpoints than 32. A long isolated range can retain about four times as many checkpoint records as before; the existing cache budget still bounds them. See `results/checkpoint-comparison.json` for samples and heap measurements.

On the 34,000-row book retained heap increased from 51.77 to 52.04 MiB in closures and 52.05 to 52.36 MiB in eval, above the shared input and verifier reference. These figures exclude peak RSS. The sparse book retains 1.54 MiB in closures, 235.94 MiB in HF default dense mode, and 0.86 MiB in HF automatic sparse mode.

Both reports completed 29 of 30 variants with matching values; HF timed out opening the 34,000-row business book. The optimized report made 80,622,748 repeated cell comparisons, with maximum relative numeric difference below 5e-15. Every workload hash and input hash match the baseline. The synthetic hash prefixes remain `101578a9c6b600b1` for 3,400 rows and `e64d5f4d0f0ac632` for 34,000.

`npm run typecheck`, build, full source and dist suites passed. Each full suite ran 20,220 passing tests with 42 skipped. The benchmark adapter has seven passing tests; jscpd 4 found zero source clones with the requested 8-line/70-token thresholds. Public contracts, validation behavior and CHANGELOG are unchanged.

### Further work from HyperFormula 3.5.0

The following are further design ideas; API proposals remain unimplemented. The low-memory round below implements a conservative unchanged-result cutoff. Findings combine the published 3.5.0 package with the official documentation.

| Mechanism | Evidence in HF | xlcalc opportunity and cost |
|---|---|---|
| Stop propagation at unchanged formula results | `es/Evaluator.mjs: recomputeVertex` and `es/DependencyGraph/TopSort.mjs: postprocess` | Our unchanged-output probe evaluates 4,999 formulas in either mode versus one in HF. Propagation must also account for changed spill areas, cycles, clock dependencies and nested SUBTOTAL marks; scalar equality alone is insufficient. |
| Shared relative formula templates | [Repetitive ASTs and relative addressing](https://hyperformula.handsontable.com/docs/guide/key-concepts.html#repetitive-asts), `es/parser/ParserWithCaching.mjs` | Our prepare/compile path processes each formula separately. Template kernels could reduce opening time and compiled storage, but must retain cell-specific names, table context, array rules and runtime addresses. Measure both modes before adopting. |
| Incrementally updated lookup indexes | [Column indexes](https://hyperformula.handsontable.com/docs/guide/performance.html#vlookup-match), `es/Lookup/ColumnIndex.mjs: change` | Our exact lookup index is keyed by immutable flat storage; a changed return column invalidates that storage. A key-column index plus direct payload reads could avoid rebuilding unaffected keys, at the cost of index maintenance and memory. |
| Deferred reference transformations | [Lazy CRUD](https://hyperformula.handsontable.com/docs/guide/key-concepts.html#laziness-of-crud-operations), `es/LazilyTransformingAstService.mjs` | Could reduce structure-operation rebuilds. Versioned transforms and compaction add complexity to undo, serialization and inspection. This is a separate architectural change. |

The unchanged-output and dependency probes are reproducible with `study.mjs`; results are in `results/hf-architecture-probes.json`. When INT(A1) keeps its value after A1 changes from 1.1 to 1.9, HF evaluates only INT and both engines agree on the final value. When A1 becomes 2.1, both evaluate 4,999 formulas. Instrumentation uses private evaluator hooks for counts and is excluded from the timing benchmark.

| API idea | Current xlcalc coverage | Proposed use |
|---|---|---|
| [Immediate precedents and dependents](https://hyperformula.handsontable.com/docs/guide/dependency-graph.html#getting-the-immediate-precedents-of-a-cell-or-a-range) | `explain` supplies a precedent tree and formatted text; no reverse query | Small typed queries for grid dependency arrows. Define semantic cells/ranges rather than exposing compressed graph helpers: HF returns a range node as A1's dependent in the SUM(A1:B1) probe. |
| [Non-mutating operation checks](https://hyperformula.handsontable.com/docs/guide/basic-operations.html#isitpossibleto-methods) | Text and paste have previews; structure operations return refusals only when called | A structured check for drag/drop and disabled menu actions, with a refusal reason rather than a bare boolean. |
| [Mixed operation batching](https://hyperformula.handsontable.com/docs/guide/batch-operations.html) | Cell, text and layout batches are individually atomic | A serializable list of mixed operations could give one revision, one undo step and one recomputation. A callback API would not cross a worker boundary. |
| [Optional statistics](https://hyperformula.handsontable.com/docs/api/classes/hyperformula.html#getstats) | Timing is in benchmarks; cache accounting is internal | Opt-in phase timings and counts for application diagnostics, with no default timing overhead. HF exposes parser, graph, evaluation and criterion-cache statistics. |

Dependency queries are the first API candidate; unchanged-result propagation is the first algorithm candidate. Existing atomic batches and prefix/range caching already cover parts of HF's approach, so additions should address the gaps above rather than duplicate those features.

## Low-memory round

Measured source: `cb7078f`, including master `671e5e8` (0.21.1). Baseline: `794fade`, the preceding parity report. Public APIs and validation behavior are unchanged. Ordinary workload results are in `results-lowmem/`, retained-memory samples in `results/low-memory-comparison.json`.

Formula slots are reused after removal and trailing holes are trimmed. Replacing one formula 1,000 times previously retained 1,001 slots with one live formula. Spill evaluation timestamps are released after the round settles. The range cache stores values with their metadata in one map; it implements the ordinary Map interface, including order and undefined values, but is no longer a native Map subclass. Its byte estimates and cap are unchanged.

Small-root acyclic value edits can skip unchanged scalar paths. The existing reachability set is reused, with at most 64 root ids copied temporarily. At the first changed value or array result the remaining ordered pass runs normally. Existing arrays, structural edits and iterative calculation retain the full pass. The unchanged INT probe now evaluates one formula in either mode, matching HF; a changed INT still evaluates all 4,999. This conservative cutoff avoids a second edge walk or persistent graph.

Eval releases temporary generated text after linking. Executable functions remain; the private debugSource helper regenerates text from canonical formulas on request, without executing new Function. The helper now also gives code for closure sessions. Three independent probes released 13.01 MiB of text and associated string fragments on fanout, and 1.764 MiB on the large business book.

### Memory

Five sequential alternating pairs in fresh children, with a 512 MiB heap cap and 30-second process deadline. Shared input is excluded from retained heap; history is disabled to isolate engine storage. No same-book warmup precedes the heap baseline. The live session is pinned, with two event-loop turns and two GCs before reading memory. Every non-cache case matches an independent full computation afterward.

| Case | Mode | Heap before, MiB | After | Process peak RSS before, MiB | After |
|---|---|---:|---:|---:|---:|
| unchanged | closures | 14.55 | 14.10 | 201.75 | 157.31 |
| unchanged | eval | 29.40 | 19.62 | 232.56 | 190.28 |
| churn | closures | 2.13 | 1.83 | 138.53 | 100.19 |
| churn | eval | 2.18 | 1.86 | 140.19 | 100.47 |
| cache-only | closures | 12.62 | 10.81 | 102.28 | 94.98 |
| fanout | closures | 65.03 | 63.26 | 232.56 | 231.17 |
| fanout | eval | 105.76 | 90.99 | 389.98 | 389.25 |
| prefix-statistics | closures | 13.85 | 13.20 | 118.33 | 117.31 |
| prefix-statistics | eval | 19.68 | 16.21 | 149.20 | 148.48 |
| business-34000 | closures | 52.23 | 52.10 | 362.25 | 361.53 |
| business-34000 | eval | 58.05 | 56.22 | 380.61 | 379.97 |
| lookup | closures | 10.99 | 10.97 | 143.06 | 141.56 |
| lookup | eval | 12.15 | 11.64 | 146.33 | 146.39 |

RSS is the whole-process watermark before the full-compute verifier runs, from [process.resourceUsage](https://nodejs.org/api/process.html#processresourceusage), not peak JavaScript heap. Large-book RSS changes little because compilation still needs temporary text and native compiler storage. These are Node measurements, not browser measurements.

An initial measurement warmed the same book before the heap baseline. Residues produced zero or negative deltas, so those heap figures were rejected. The recorded comparison uses independent cold children; edit timing below includes runtime warmup and is separate from the ordinary editing medians.

| Special probe | Mode | Total edits before, ms | After | Edits |
|---|---|---:|---:|---:|
| churn | closures | 329.20 | 45.60 | 5000 |
| churn | eval | 334.62 | 50.91 | 5000 |
| unchanged | closures | 91.64 | 52.95 | 50 |
| unchanged | eval | 111.94 | 52.65 | 50 |

Churn replaces one formula 5,000 times. Unchanged alternates an input 50 times in a 10,000-cell chain whose INT result stays unchanged. Cache-only holds 50,000 small entries. Reproduce with:

```sh
node bench/hyperformula/memory.mjs --before /path/to/compiled/baseline --after . --out /tmp/xlcalc-memory
node bench/hyperformula/memory.mjs --before /path/to/compiled/baseline --case cache-only --out /tmp/xlcalc-cache
node bench/hyperformula/run.mjs --out /tmp/xlcalc-lowmem-speed
```

### Speed and checks

The original nine-workload protocol still includes normal history and both execution modes. Twenty-nine variants match; HF times out opening the large business book. All workload hashes and input hashes match the previous report, across 80,622,748 repeated cell comparisons. A 160 ms count probe overlapped the initial first child, so that business-3400/closures record was replaced by one isolated repeat after the full run. The discarded record and reason remain in raw.json.

| Edit | closures, ms | eval, ms | HF, ms |
|---|---:|---:|---:|
| prefix-sum | 4.66 | 4.30 | 4.86 |
| prefix-statistics | 10.16 | 9.82 | 10.96 |
| chain | 5.25 | 5.35 | 5.11 |
| fanout | 34.50 | 40.37 | 43.60 |
| bulk | 4.68 | 4.70 | 3.09 |
| lookup | 3.49 | 3.42 | 2.38 |

The batch and indexed lookup gap remain. More selective pruning after another root changes, persistent key-column indexes and shared relative formula templates are further candidates. None is added here; these patches reduce retained storage without expanding the cache budget.

Full source and dist suites each pass 20,261 tests with 42 skipped. Typecheck, build and jscpd (zero clones) pass. The new conformance checks cover all three session kinds and both modes. Exact synthetic hashes remain unchanged.

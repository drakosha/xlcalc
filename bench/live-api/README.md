# Component acceptance for live workbook APIs

Run from the library checkout after `npm run build`:

```sh
node bench/live-api/component.mjs --component /path/to/xlcalc-sheet
```

The runner copies the component into a temporary directory, keeps its source checkout and installed library unchanged, and installs a physical copy of this checkout's compiled `xlcalc` package. It runs the component's typecheck, builds both entries and runs its tests. Hidden dependency-cache directories are local to the copy.

`component-adapter.test.ts.template` adds a synthetic agent scenario. The existing local source and the actual Vue grid share one session. A mixed cell/name/format proposal is previewed, applied with its revision and clock, and delivered to the grid through its existing reload event. The grid changes from 7 to 9 and shows the new number format. Find and summary run in the session. Inline and pool calculators are tested in closures and eval; a Node WorkerLike adapter lets the pool run under the DOM test environment.

Library conformance is in `test/session-live-api.test.ts`, with all three session kinds and both execution modes. It covers atomic refusal, one history step, dependent impacts, explicit live clocks, projection, overlap deduplication, cursors, large spills, reverse dependencies and cancellation. The pool cancellation case uses zero-length slices to make the message observable before the computation completes, rather than assuming a fixed edit duration.

The engine benchmarks remain in `bench/hyperformula`. `results-live-api/` holds the nine-workload HyperFormula comparison and five alternating pairs of fresh-process memory samples against 0.21.2. `results-live-api-before/` holds the same xlcalc speed protocol on the frozen 0.21.2 dist. Use a compiled baseline without switching the worktree:

```sh
XLCALC_BENCH_ROOT=/path/to/compiled/baseline node bench/hyperformula/run.mjs --engine xl-closures,xl-eval --out /tmp/xlcalc-before
node bench/hyperformula/memory.mjs --before /path/to/compiled/baseline --after . --out /tmp/xlcalc-memory
```

Run measurements sequentially without concurrent tests or CPU profiles. Retained heap excludes shared input and holds the session after two event-loop turns and two GCs; maximum RSS is the process watermark before full-compute verification. It includes native compiler storage and is not browser heap.


## Completed acceptance

Fresh master `9113b25` (0.22.0) was merged into the work branch. The merge preserves the external-link implementation and format categories. Entered linked formulas use an aside-built batch state, so a preview cannot leak link refresh markers or cached errors into the original session; six fail-before checks cover the three session kinds and two modes.

Source and dist each pass 20,563 tests with 42 skipped. Both full suites use one test worker because the inherited cold-budget test failed under parallel load; its 1,500 ms threshold is unchanged. The strict-argument method catalogue now includes the six added methods. The component copy passes typecheck, both builds and all 309 tests, including four synthetic live-agent/grid cases. The runner installs dist as an ordinary directory, so the component's bundle-boundary path assertions are checked as in a normal package installation. jscpd reports zero clones across 284 source files.

The work branch changes no existing validation policy, publishes no version, and has no own package-version or changelog change relative to the merged master. The app and the the app were not migrated. Component source and its original dependency were not edited.

Package checks (publint and the ESM export/type profile) pass. After the master merge, both execution modes retain the synthetic hashes `101578a9c6b600b1` (3400 rows) and `e64d5f4d0f0ac632` (34000 rows), and the input hashes `1fb3575501a5acbf` and `ca4999a5e730cfc4`. Performance artifacts identify their measured commits before the 0.22.0 merge; full correctness and component acceptance identify the merged implementation. No browser-memory result is claimed.

# Sheetmark

**A fidelity benchmark for spreadsheet calculation engines. It measures how
faithfully an engine reproduces Microsoft Excel's computed results, cell by
cell, on a fixed public corpus — and whether the engine fails loudly or
silently on the cells it can't reproduce.**

Sheetmark is the published spreadsheet-recalculation fidelity benchmark from the
Recalc project. It runs a public corpus of real workbooks through a calculation
engine and scores every formula cell against Excel's own computed value. It
reports two things a raw "% accurate" never tells you: exactly *where* an engine
diverges, and — the axis no one else publishes — whether it flags the cells it
can't do, or quietly returns a wrong number.

---

## The property that matters

> **Exactly Excel — or loudly flagged. Never a silent wrong answer.**

Every cell an engine reports is either **exactly** what Excel computes (at a
documented tolerance), or it is **explicitly flagged as not attempted**. Guessing
a plausible-looking value is not allowed. A cell an engine declines — returning a
distinguishable sentinel such as `#UNSUPPORTED!` — is scored as *not attempted*,
never counted as a match and never counted as a fidelity failure.

That is a stronger guarantee than "accurate," and it is worth more to anyone who
has to *defend* a number. A 99.99% engine that silently lies on the other 0.01%
is worse than a 98% engine that tells you precisely which cells it couldn't do.
A wrong number you can see is a footnote; a wrong number you can't see is a
liability. Reliability isn't "usually right" — it's knowing exactly when not to
trust the answer.

Fidelity here is a **measured number, not an adjective**: measured agreement with
Excel's own computed results on a public corpus, at a documented tolerance. The
percentages below are *evidence* for the trust property. They are never the
headline.

---

## Headline results — Recalc, Sheetmark v3

Measured on the FUSE public corpus (3,640 `.xlsx`/`.xlsm` workbooks, 5,667,851
formula cells with an oracle value), engine build `bbb6984`, re-scored
2026-07-19. Full detail and the dated result set: **[RESULTS.md](RESULTS.md)**.

A single percentage hides the story, so we publish the whole funnel first — every
oracle cell in exactly one bucket — then the two figures derived from it. Strict
and lenient always travel together, and the 15-significant-figure headline always
carries its conservative bit-exact floor.

| Metric | 15-sig headline | Bit-exact floor |
|---|---|---|
| **Engine fidelity** (lenient — of the cells it attempted) | **98.637%** | 97.497% |
| **Coverage-inclusive** (strict — over every oracle cell) | **78.888%** | 77.977% |
| Genuine disagreements (mismatch cells) | 61,801 | 113,458 |

- **Lenient** is the engine-quality signal: of the 4,533,056 cells Recalc
  attempted (80.0% of the corpus's oracle cells), how many matched Excel.
- **Strict** is quality × coverage: it counts every declined cell as a miss, so
  it is bounded by how much of the corpus falls inside the engine's declared
  scope — not by whether the math was right.

Neither number implies 100%, and 100% is not the target — for reasons that are
measured, not rhetorical. A clean 100% is *stricter than Excel meets itself*:
Excel does not reproduce its own results bit-for-bit across versions and
platforms, and its own arithmetic is not mathematics — a dense probe of
`=ERFC(x)` at 1,130 points found Excel's error function drifting up to ~62 ULP
off the mathematically-correct value, so Recalc had to *delete* a precision
refinement to match it. Recalc returns what Excel returns, not what mathematics
returns. And part of the residual is the oracle itself being stale — a pre-2007
conversion cache, or precision-as-displayed rounding — where the engine's answer
is arguably the more correct one. The full argument, decomposed to the same
discipline as the declined bucket:
**[why 100% is not on the table — measured](METHODOLOGY.md#why-100-is-not-on-the-table)**.

### The funnel — every oracle cell in exactly one bucket

| Bucket | Cells | Share of corpus |
|---|---:|---:|
| **Computed and matched** (at the 15-sig tolerance) | 4,471,255 | 78.89% |
| **Declined loudly** (`#UNSUPPORTED!`, per cell, never guessed) | 1,134,795 | 20.02% |
| **Genuine disagreement** (the only fidelity failure) | 61,801 | 1.09% |

The declined bucket is dominated by **corpus composition, not engine
correctness**. Its measured by-cause split:

| Cause of decline | Share of declined | Cells |
|---|---:|---:|
| External-workbook links (references to other workbooks the corpus doesn't ship) | 86.90% | 986,069 |
| Unimplemented functions (led by `TTEST`, `MMULT`, `COUNTBLANK`, `ATAN2`) | 0.58% | 6,616 |
| Other (structured references, parse errors, reference-form constructs, runtime refusals) | 12.09% | 137,185 |
| Volatile / UDF / blocked-I/O / array residual | 0.43% | 4,925 |

The single largest class — 86.90% — is cells that reference *other workbooks the
benchmark does not ship*. Recalc returns `#UNSUPPORTED!` there **correctly**: no
network and no filesystem from a formula is a hard engine rule, and
external-workbook resolution is an explicit non-goal. The engine is declining
for lack of inputs, cell by cell — not getting a calculation wrong. This is why
strict must never be read as an engine-quality failing without its context.
[LEADERBOARD.md](LEADERBOARD.md) measures this class per engine.

The **unimplemented-function class** — 15.05% of declined in v2 — now stands at
**0.58% (6,616 cells)**, led by `TTEST`, `MMULT`, `COUNTBLANK`, and `ATAN2`:
formulas calling functions the engine does not yet implement, each declined
loudly and never guessed. This is the class that shrinks as function coverage
grows — the direct measure of the engine's remaining scope.

---

## Corpus, oracle, tolerance — in one line each

- **Corpus.** FUSE (Barik, Lubick, Smith, Slankas, Murphy-Hill, MSR 2015) — a
  public, web-crawled spreadsheet corpus. Sheetmark scores the formula-bearing
  `.xlsx`/`.xlsm` cut: 3,640 workbooks, identified and frozen by the publisher's
  per-file SHA-1 manifest.
- **Oracle.** For every result set to date, each workbook's own last-computed Excel value (the cached
  `<v>` Excel wrote into every formula cell) — disclosed plainly, producer-
  heterogeneous, no oracle dependency. A pinned Microsoft 365 build,
  **16.0.20131**, backs the semantic probe experiments that fix the engine's
  behavior.
- **Tolerance.** Text, logical, and error values match **exactly**. Numbers match
  when equal to **15 significant figures** — Excel's own cache-storage precision —
  and every 15-sig figure is published alongside its bit-exact floor.

Full definitions, the classification rules, the integrity discipline, and the
known caveats: **[METHODOLOGY.md](METHODOLOGY.md)**.

---

## Leaderboard — first measured release

Sheetmark is built as a *comparative* benchmark: same corpus, same oracle, same
tolerance, and the same loud-vs-silent-failure axis for every engine.

**First measured release — see [LEADERBOARD.md](LEADERBOARD.md):** LibreOffice
25.2.3.2, pycel 1.0b30, and formulas 1.2.10, measured 2026-07-19 on the full
corpus, both tolerances, alongside Recalc's own funnel — with the
external-workbook-reference decomposition, a mechanical glossary, and the full
caveat list.

**No competitor number appears here until Sheetmark has actually measured that
engine.** No engine's figure — including Recalc's own — is ever quoted from
memory or a vendor's own claim. A comparative result requires a measured
competitor run first, and not one minute before.

---

## Documents

- **[METHODOLOGY.md](METHODOLOGY.md)** — the full method: what is measured and how
  it's classified, the two metrics and their denominators, the corpus and oracle,
  the tolerance regime, scope, reproducibility, and the known caveats.
- **[RESULTS.md](RESULTS.md)** — the dated measured result sets (v3 latest; v2
  and v1 retained below it for the record). Each re-score mints a new dated
  result set.
- **[LEADERBOARD.md](LEADERBOARD.md)** — the measured engine comparison:
  LibreOffice, pycel, formulas, and Recalc on the same corpus, oracle, and
  classification rules, with the external-reference decomposition, glossary,
  and caveats.

---

## FAQ

**How accurate is it?**
On the cells it computes, Recalc matched Excel's stored results on **98.637% of
the 4,533,056 cells it attempted** — 80.0% of a 5,667,851-cell public corpus — at
a documented 15-significant-figure tolerance; **97.497% bit-exact** with no
tolerance at all. Genuine disagreements: **61,801 cells, 1.09% of the whole
corpus** — an upper bound on engine error, because it also counts Excel
disagreeing with its own stored values (stale caches, precision-as-displayed
rounding). Every cell it can't compute faithfully —
20.02% of this corpus, 86.90% of it referencing other workbooks we weren't given —
is declined loudly rather than guessed, so you always know exactly which cells to
trust.

**Why is the coverage-inclusive (strict) number lower than the engine (lenient)
number?**
Because strict counts every declined cell as a miss — and 86.90% of those declines
reference other workbooks the corpus doesn't ship (external links). Recalc
correctly refuses to fabricate those (no network, no filesystem from a formula).
That's a refusal over missing inputs, not a calculation error. Lenient is the
engine-quality signal; strict is quality × coverage, bounded by what the corpus
asks of the engine's declared scope.

**Will an engine ever return a wrong number without flagging it?**
That is the one thing the design forbids. Every cell is either exactly Excel's
result to the documented tolerance, or a distinguishable `#UNSUPPORTED!`. Guessing
a plausible value is not allowed anywhere in the engine.

**What is the number measured against?**
Excel's own computed results, at a documented tolerance — a bar that is *stricter
than Excel meets itself*, since Excel doesn't reproduce its own results
bit-for-bit across versions and platforms, and its own floating-point arithmetic
diverges from mathematical truth (measured: up to ~62 ULP on `ERFC`). A number
measured that strictly is a strong claim, not a weak one — and it is why a clean
100% is not on the table:
**[the measured argument](METHODOLOGY.md#why-100-is-not-on-the-table)**.

**Was Excel reverse-engineered, or its code copied?**
No copied code. Recalc is a clean-room implementation: behavior is observed by
running a pinned Excel build as an oracle and matching its outputs, never by
reading a competing spreadsheet's source.

**Do I need Excel or Windows to run the engine?**
No. The engine is headless and self-contained — no Excel install, no Windows, no
GUI automation.

---

## Trademark and independence

Microsoft and Excel are trademarks of the Microsoft group of companies. Recalc
and Sheetmark are independent projects, not affiliated with, sponsored by, or
endorsed by Microsoft. References to Excel are nominative — they identify the
product whose behavior Sheetmark measures compatibility against.

Sheetmark's benchmark documents are published under
[CC BY 4.0](LICENSE) — reuse with attribution. The FUSE corpus is credited to
its authors (Barik, Lubick, Smith, Slankas, Murphy-Hill, MSR 2015): its
canonical Zenodo record ([DOI 10.5281/zenodo.581678](https://doi.org/10.5281/zenodo.581678))
is published under CC BY 4.0, and the binaries distribution Sheetmark actually
pins (barik.net) is described by its publisher as "unencumbered by any license
agreements." Details in [METHODOLOGY.md](METHODOLOGY.md).

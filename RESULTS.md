# Sheetmark Results — v2

**Engine:** Recalc, build `f8fd3b7`
**Corpus:** FUSE `.xlsx`/`.xlsm` cut — 3,640 workbooks, 5,667,851 formula cells
with an oracle value
**Oracle:** cached-value path (each workbook's own last-computed Excel value)
**Tolerance:** 15 significant figures (headline), with the bit-exact floor
published alongside
**Result set dated:** 2026-07-16 (corrected decline attribution + a small set of
engine fixes; zero load failures)

This is the v2 measured result set. **Each full re-score mints a new dated result
set**, stamped with its engine build, corpus freeze, oracle path, and tolerance
mode; the numbers below are frozen at this build and are not edited in place. The
v1 set (build `66fa5f9`, 2026-07-15) is preserved unchanged further down, for the
record. For the definitions of every term, the classification rules, and the
honest caveats, see [METHODOLOGY.md](METHODOLOGY.md).

All figures publish **verbatim and unrounded**, always strict + lenient together,
and the 15-significant-figure headline always with its bit-exact floor.

---

## The funnel — every oracle cell in exactly one bucket

| Bucket | Cells | Share of corpus |
|---|---:|---:|
| **Computed and matched** (at the 15-sig tolerance) | 4,308,561 | 76.02% |
| **Declined loudly** (`#UNSUPPORTED!`, per cell, never guessed) | 1,310,658 | 23.12% |
| **Genuine disagreement** (the only fidelity failure) | 48,632 | 0.86% |
| **Total oracle cells** | 5,667,851 | 100% |

Attempted cells (the lenient denominator) number **4,357,193 — 76.9%** of the
5,667,851 oracle cells. Every other cell is declined loudly, per cell, never
guessed.

---

## Fidelity — both metrics, both tolerances

| Metric | 15-sig headline | Bit-exact floor |
|---|---|---|
| **Engine fidelity** (lenient — on attempted cells) | **98.884%** (= 4,308,561 / 4,357,193) | 97.857% |
| **Coverage-inclusive** (strict — over all oracle cells) | **76.018%** (= 4,308,561 / 5,667,851) | 75.228% |
| Mismatch cells | 48,632 | 93,373 |

- **Lenient** is the engine-quality signal — of the cells the engine attempted,
  how many matched Excel.
- **Strict** is quality × coverage — it counts every declined cell as a miss, so
  it is bounded by how much of the corpus falls inside the engine's declared
  scope, not by whether the math was right.

---

## The declined bucket, by cause

The declined bucket (1,310,658 cells, 23.12% of the corpus) is dominated by corpus
composition, not engine correctness. Its measured full-corpus decomposition:

| Cause of decline | Share of declined | Cells |
|---|---:|---:|
| External-workbook links (references to other workbooks the corpus doesn't ship) | 75.23% | 986,069 |
| Unimplemented functions (led by `NORMDIST`, `HYPERLINK`, `RANK`, `SIN`) | 15.05% | 197,268 |
| Other (structured references, parse errors, reference-form constructs, runtime refusals) | 9.34% | 122,377 |
| Shared-formula residual (retired — see "What changed vs v1") | 0.00% | 0 |
| Volatile / UDF / blocked-I/O | 0.38% | 4,944 |

*(Rows sum to the declined total exactly: 986,069 + 197,268 + 122,377 + 0 + 4,944
= 1,310,658.)*

The single largest class — 75.23% — is cells referencing *other workbooks the
benchmark does not ship*. The engine returns `#UNSUPPORTED!` there correctly: no
network and no filesystem from a formula is a hard rule, and external-workbook
resolution is a declared non-goal. That dominant class is not the engine getting a
calculation wrong; the engine is declining for lack of inputs, and saying so
cell by cell.

The second-largest class is **unimplemented functions** — 15.05%, led by
`NORMDIST`, `HYPERLINK`, `RANK`, and `SIN`. These are formulas calling functions
the engine does not yet implement; each is declined loudly, per cell, and none is
guessed. This is the class that grows as the engine's function coverage grows, and
it is the direct measure of remaining engine scope.

---

## Genuine disagreement

**48,632 cells — 0.86% of the corpus** — at the 15-significant-figure tolerance
(**93,373 cells** bit-exact). This is the only bucket that counts as a fidelity
failure. A meaningful share of even this residual is Excel disagreeing with its own
stored values — stale `.xls`-conversion caches and precision-as-displayed rounding
— rather than the engine's math (see the caveats in [METHODOLOGY.md](METHODOLOGY.md)),
so the published residual is an upper bound on genuine engine disagreement.

---

## What changed vs v1

Two independent effects, kept separate.

1. **A re-attribution audit corrected the declined-bucket split.** The v1
   disclosure labelled 16.3% of declined cells (213,468) as a "shared-formula
   residual" — groups whose master couldn't yet be expanded. A re-attribution
   audit found that label was wrong: those cells' group masters expand fine (the
   expander now handles 100% of the corpus's shared follow-on cells), and the
   cells decline for other, downstream reasons — **dominantly unimplemented
   functions**, plus external-workbook links and runtime refusals. The classifier
   was corrected and the split re-measured. This changes labels, not the declined
   *set*, so the funnel is untouched by it. The re-measured split retires the
   "shared-formula residual" class to **0** and re-homes its cells to their true
   causes: unimplemented functions rises from 37,168 to 197,268, external links
   from 965,848 to 986,069, and the remainder to the volatile and other classes.

2. **A small set of engine fixes moved ~2,600 cells from declined to exact.** Net
   of the engine-behavior changes since v1, the declined set shrank by **2,577
   cells**, and **all 2,577 became exactly correct** — a clean declined → exact
   trade in *both* the 15-sig and bit-exact scoring modes. **Mismatch counts are
   unchanged** in both modes (48,632 / 93,373): no previously-wrong cell was
   created or silently fixed. The strict headline lifts **+0.046 pp** (15-sig) and
   **+0.045 pp** (bit-exact); lenient holds.

The number published is always the one *after* such a fix, and the classification
is always the corrected one — never the flattering-by-omission or the mislabelled
version.

---

# Sheetmark Results — v1

**Engine:** Recalc, build `66fa5f9`
**Corpus:** FUSE `.xlsx`/`.xlsm` cut — 3,640 workbooks, 5,667,851 formula cells
with an oracle value
**Oracle:** cached-value path (each workbook's own last-computed Excel value)
**Tolerance:** 15 significant figures (headline), with the bit-exact floor
published alongside
**Result set dated:** 2026-07-15 (post shared-formula expansion; zero load
failures)

This is the v1 measured result set. **Each full re-score mints a new dated result
set**, stamped with its engine build, corpus freeze, oracle path, and tolerance
mode; the numbers below are frozen at this build and are not edited in place. For
the definitions of every term, the classification rules, and the known caveats,
see [METHODOLOGY.md](METHODOLOGY.md).

All figures publish **verbatim and unrounded**, always strict + lenient together,
and the 15-significant-figure headline always with its bit-exact floor.

---

## The funnel — every oracle cell in exactly one bucket

| Bucket | Cells | Share of corpus |
|---|---:|---:|
| **Computed and matched** (at the 15-sig tolerance) | 4,305,984 | 75.97% |
| **Declined loudly** (`#UNSUPPORTED!`, per cell, never guessed) | 1,313,235 | 23.17% |
| **Genuine disagreement** (the only fidelity failure) | 48,632 | 0.86% |
| **Total oracle cells** | 5,667,851 | 100% |

Attempted cells (the lenient denominator) number **4,354,616 — 76.8%** of the
5,667,851 oracle cells. Every other cell is declined loudly, per cell, never
guessed.

---

## Fidelity — both metrics, both tolerances

| Metric | 15-sig headline | Bit-exact floor |
|---|---|---|
| **Engine fidelity** (lenient — on attempted cells) | **98.883%** (= 4,305,984 / 4,354,616) | 97.856% |
| **Coverage-inclusive** (strict — over all oracle cells) | **75.972%** (= 4,305,984 / 5,667,851) | 75.183% |
| Mismatch cells | 48,632 | 93,373 |

- **Lenient** is the engine-quality signal — of the cells the engine attempted,
  how many matched Excel.
- **Strict** is quality × coverage — it counts every declined cell as a miss, so
  it is bounded by how much of the corpus falls inside the engine's declared
  scope, not by whether the math was right.

---

## The declined bucket, by cause

The declined bucket (1,313,235 cells, 23.17% of the corpus) is dominated by corpus
composition, not engine correctness. Its measured full-corpus decomposition:

| Cause of decline | Share of declined | Cells |
|---|---:|---:|
| External-workbook links (references to other workbooks the corpus doesn't ship) | 73.5% | 965,848 |
| Shared-formula residual (groups whose master can't yet be parsed/expanded) | 16.3% | 213,468 |
| Unimplemented functions (led by `HYPERLINK`, `ERFC`, `RANK`) | 2.8% | 37,168 |
| Other (structured references, parse errors, reference-form constructs) | 7.3% | 96,263 |
| Volatile / UDF / blocked-I/O | <0.1% | 488 |

*(Rows sum to the declined total exactly: 965,848 + 213,468 + 37,168 + 96,263 +
488 = 1,313,235.)*

The single largest class — 73.5% — is cells referencing *other workbooks the
benchmark does not ship*. The engine returns `#UNSUPPORTED!` there correctly: no
network and no filesystem from a formula is a hard rule, and external-workbook
resolution is a declared non-goal. That dominant class is not the engine getting a
calculation wrong; the engine is declining for lack of inputs, and saying so
cell by cell.

---

## Genuine disagreement

**48,632 cells — 0.86% of the corpus** — at the 15-significant-figure tolerance
(**93,373 cells** bit-exact). This is the only bucket that counts as a fidelity
failure. A meaningful share of even this residual is Excel disagreeing with its own
stored values — stale `.xls`-conversion caches and precision-as-displayed rounding
— rather than the engine's math (see the caveats in [METHODOLOGY.md](METHODOLOGY.md)),
so the published residual is an upper bound on genuine engine disagreement.

---

## What moved the number since the prior result set

The first full decomposition of the declined bucket found that **65% of all
declined cells were unexpanded shared-formula follow-on cells** — an addressable
engine gap, not a scope boundary. Implementing shared-formula expansion (ECMA-376
§18.17.2) for this result set moved **~3.10M cells from declined to attempted**;
**3,072,264 (99.14%) of them match Excel's own cached values**, lifting strict from
**21.8% to 76.0%** while lenient held (and rose slightly). Expanding those cells
also *surfaced* **26,639 previously-hidden mismatches** — they were declined before,
so uncounted — now honestly counted in the 48,632 headline mismatch total. The
shared-formula residual fell from **2,873,625 to 213,468 cells, a 92.6% reduction**.

The number published is always the one *after* such a fix, never the
flattering-by-omission one before it.

---

## Trademark and independence

Microsoft and Excel are trademarks of the Microsoft group of companies. Recalc and
Sheetmark are independent projects, not affiliated with, sponsored by, or endorsed
by Microsoft. References to Excel are nominative — they identify the product whose
behavior Sheetmark measures compatibility against.

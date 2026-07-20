# Sheetmark Leaderboard — first measured release

> **As measured: LibreOffice 25.2.3.2 (520 Build:2), pycel 1.0b30, formulas
> 1.2.10, on 2026-07-19; results describe those versions only.** Supporting
> toolchain, recorded alongside the numbers: Python 3.13.5, numpy 2.5.1,
> openpyxl 3.1.5, schedula 1.6.15. Recalc build `bbb6984`, measured the same
> day on the same corpus (its funnel is the [v3 result set](RESULTS.md)). The
> external-reference decomposition below was measured 2026-07-20 with the same
> engine versions.

Every engine on this page is measured on the **same corpus, the same oracle,
the same classification rules, and the same funnel arithmetic**. The harness
that produced these numbers is published (Apache-2.0) — see
[Reproducing these numbers](#reproducing-these-numbers). Every term used on
this page is defined mechanically in the [Glossary](#glossary); every
percentage carries its denominator.

**Shared denominators (identical for all engines, all modes):** FUSE
`.xlsx`/`.xlsm` cut — 3,640 workbooks, **5,739,091 formula cells**, of which
**71,240** are `no_oracle` (unjudgeable) → **strict denominator 5,667,851**.
Oracle load failures: **0**. Every funnel row sums `exact + mismatch +
declined + no_oracle = 5,739,091` exactly; a crashed or timed-out workbook's
cells are counted per the conservation rule, never dropped.

---

## The funnels — both tolerances

**15-significant-figure tolerance (headline):**

| Engine | exact | mismatch | declined | no_oracle | dead workbooks | strict | lenient |
|---|---:|---:|---:|---:|---:|---:|---:|
| LibreOffice 25.2.3.2 | 5,494,671 | 152,826 | 20,354 | 71,240 | 5 | 96.94% | 97.29% |
| pycel 1.0b30 | 3,958,825 | 3,297 | 1,705,729 | 71,240 | 34 | 69.85% | 99.92% |
| formulas 1.2.10 | 3,130,924 | 229,508 | 2,307,419 | 71,240 | 190 | 55.24% | 93.17% |
| Recalc `bbb6984` | 4,471,255 | 61,801 | 1,134,795 | 71,240 | 0 | 78.89% | 98.64% |

**Bit-exact (conservative floor):**

| Engine | exact | mismatch | declined | no_oracle | dead workbooks | strict | lenient |
|---|---:|---:|---:|---:|---:|---:|---:|
| LibreOffice 25.2.3.2 | 5,404,796 | 242,701 | 20,354 | 71,240 | 5 | 95.36% | 95.70% |
| pycel 1.0b30 | 3,881,865 | 3,670 | 1,782,316 | 71,240 | 37 | 68.49% | 99.91% |
| formulas 1.2.10 | 3,111,130 | 320,890 | 2,235,831 | 71,240 | 186 | 54.89% | 90.65% |
| Recalc `bbb6984` | 4,419,598 | 113,458 | 1,134,795 | 71,240 | 0 | 77.98% | 97.50% |

Exact strict/lenient values, unrounded, from the merged run outputs:

| Engine | strict 15-sig | lenient 15-sig | strict bit-exact | lenient bit-exact |
|---|---:|---:|---:|---:|
| LibreOffice | 96.94452095% | 97.29391623% | 95.35882295% | 95.70250325% |
| pycel | 69.84701962% | 99.91678701% | 68.48918576% | 99.90554711% |
| formulas | 55.24005483% | 93.17028287% | 54.89082194% | 90.65011276% |
| Recalc | 78.888% | 98.637% | 77.977% | 97.497% |

`strict = exact / 5,667,851` (every declined cell counts as a miss);
`lenient = exact / (5,667,851 − declined)` (declined cells leave the
denominator). Recalc's strict/lenient are as printed by its native verifier on
the identical corpus, oracle, and totals; a published parity suite shows the
harness reproduces the native verifier's funnel cell-for-cell on shared
fixtures.

Reading the columns mechanically:

- The strict-vs-lenient split for each engine is entirely its `declined`
  column: LibreOffice declines 20,354 cells (0.36% of the strict
  denominator), pycel 1,705,729–1,782,316 (30.1–31.4%), formulas
  2,235,831–2,307,419 (39.4–40.7%), Recalc 1,134,795 (20.02%).
- Silent wrong answers (mismatch — see [Glossary](#glossary)) at 15-sig:
  LibreOffice 152,826, pycel 3,297, formulas 229,508, Recalc 61,801. At
  bit-exact: 242,701 / 3,670 / 320,890 / 113,458.
- Dead workbooks (engine-level crash or timeout on a workbook the oracle
  loaded; all its oracle-bearing cells fold into `declined` per the
  conservation rule): formulas 186–190, pycel 34–37, LibreOffice 5, Recalc 0.
- pycel's bit-exact and 15-sig mismatch counts are nearly identical (3,670 vs
  3,297), so its mismatches are genuine divergences rather than
  last-digit float noise. LibreOffice's mismatch count drops 242,701 →
  152,826 from bit-exact to 15-sig: 89,875 cells differ from the stored value
  only below the 15th significant figure.

### LibreOffice recalculation verification

The LibreOffice numbers are meaningful only if LibreOffice actually
recalculated rather than reading back the workbook's stored values. Three
recorded checks:

1. **Code path.** The dump macro calls the UNO
   `XCalculatable::calculateAll()` — an unconditional full-document
   recalculation — immediately after loading and before any value is read.
2. **Provenance.** The harness files that executed on the measurement host
   are md5-identical to the published harness sources.
3. **Measured signature.** A pure read-back of stored values would score
   ~100% bit-exact, because the oracle is the stored value. The measured
   funnel instead shows 242,701 bit-exact mismatches (89,875 of which clear
   at 15-sig — last-digit float divergence, the signature of an independent
   computation) and 20,354 declines dominated by LibreOffice-internal
   `Err:NNN` results, a state that only exists after recalculation.

The one systematic exception — cells referencing absent external workbooks,
where `calculateAll` retains the stored value — is measured separately in the
next section.

---

## External-workbook-reference decomposition

986,069 of the strict denominator's cells (**17.4%**) reference *other
workbooks the corpus does not ship* — 928,246 directly and 57,823 by
cascade, across 209 workbooks. On these cells no engine possesses the
inputs: the linked workbooks are absent. An engine can decline the cell, or
it can retain the workbook's stored (cached) value — and the stored value is
also the oracle, so a retained cached value scores `exact` without any
computation having occurred. The engines measured here differ sharply on
exactly this set, which is why it is decomposed rather than left inside the
aggregate.

The cell set is defined by Recalc's per-cell decline attribution (it declines
all such cells under its no-network/no-filesystem rule) and reconciles
exactly: 928,246 + 57,823 = 986,069 of its 1,134,795 declined cells. Each
engine's per-cell results on the 209 workbooks were then intersected with the
set, scored through the same oracle and classifier as the funnels above, with
both tolerances scored from a single engine run (`missing_ext_keys = 0` for
every engine).

**On the 986,069 external-reference cells (15-sig; the bit-exact column
differs only for LibreOffice, shown in parentheses):**

| Engine | exact (= stored value retained) | mismatch | declined | share of the 986,069 scored exact |
|---|---:|---:|---:|---:|
| Recalc `bbb6984` | 0 | 0 | 986,069 | 0.00% |
| LibreOffice 25.2.3.2 | 970,597 (969,644) | 15,472 (16,425) | 0 | 98.43% (98.33%) |
| pycel 1.0b30 | 70,101 | 0 | 915,968 | 7.11% |
| formulas 1.2.10 | 7,273 | 54,566 | 924,230 | 0.74% |

Measured statements, with denominators:

- **LibreOffice scores exact on 970,597 of the 986,069 external-reference
  cells (98.43%) with zero declines.** Those 970,597 cells are **17.7% of
  its entire 5,494,671-cell 15-sig exact mass**. Its 15,472 15-sig
  mismatches on the set are cells where it recomputed to a different value,
  so its behavior on the set is stored-value-dominant, not uniform.
- **Recalc declines 100%** of the set (986,069 / 986,069), by declared rule.
- **pycel declines 92.9%** (915,968 / 986,069) and scores 70,101 exact
  (7.11%) — cells where its loader surfaced the stored value.
- **formulas declines 93.7%** (924,230 / 986,069) and retains the stored
  value on only 7,273 (0.74%). It differs from pycel in one respect: on a
  further 54,566 cells (5.5%) it *recomputed* the absent link to a value that
  disagrees with the stored one — mismatches, not retained values — so it
  neither retains nor refuses on that slice. Its bit-exact and 15-sig external
  tallies are identical (an absent link is not a float-tolerance question);
  52 of the 209 workbooks hit an engine-level failure, whose cells fold into
  the declined total per the conservation rule.
- **Raw strict gap, LibreOffice − Recalc: 18.06 pp at 15-sig / 17.38 pp
  bit-exact.**
- **On the computable universe** — the 4,681,782 strict-denominator cells
  that remain when the 986,069 external-reference cells are removed from
  numerator and denominator for *both* engines — the gap is **1.13 pp at
  15-sig** (LibreOffice 96.63% = 4,524,074 / 4,681,782 vs Recalc 95.50% =
  4,471,255 / 4,681,782) and **0.33 pp bit-exact** (94.73% = 4,435,152 /
  4,681,782 vs 94.40% = 4,419,598 / 4,681,782).

**Measurement-fairness note.** Retaining a cached value is not a wrong
answer: the cached value is the cell's correct last-computed result, and no
mismatch is being alleged on these cells. This section exists because the
two behaviors — retain vs decline — are *scored* differently by a raw strict
percentage even though neither engine computed anything: on cells whose
inputs are absent, `exact` is reachable without computation, so a raw strict
comparison structurally favors an engine that retains stored values over one
that declines. The row exists so both numbers can be read for what they
measure. The computable-universe figures above apply the identical removal
to both engines; only directly measured per-engine rows are quoted for
pycel and formulas (their whole-corpus funnels carry run-to-run timeout
variance — caveat 8 — so no subtraction against those funnels is published).

---

## Glossary

Every loaded term on this page, defined mechanically:

- **Oracle.** The workbook's Excel-stored value for a formula cell: the
  cached `<v>` the last saving producer wrote into the file. Identical for
  every engine; snapshotted before the engine runs.
- **exact.** The engine's computed value equals the oracle at the stated
  tolerance (text case-sensitive; logical and error values by kind; numbers
  bit-exact or at 15 significant figures, per mode).
- **mismatch (silent wrong answer).** The engine computed and reported a
  value, that value ≠ the workbook's Excel-stored value at the stated
  tolerance, and the engine raised no flag on the cell. The engine's genuine
  Excel-style errors (`#DIV/0!` etc.) are compared as values, not treated as
  refusals.
- **declined.** An explicit per-cell refusal: the engine produced no
  comparable value for the cell — a sentinel such as `#UNSUPPORTED!`, an
  unimplemented-function or per-cell exception, a non-Excel internal error
  marker, or a cell lost to an engine-level crash/timeout (see dead
  workbook). Declined cells count as misses under strict and leave the
  denominator under lenient.
- **dead workbook.** An engine-level crash or timeout on a workbook the
  oracle loaded successfully. Counted per the **conservation rule**: every
  oracle-bearing formula cell of that workbook is scored `declined`, every
  oracle-less cell `no_oracle`; no cell is dropped and the per-corpus total
  is conserved.
- **no_oracle.** The cell's cached value is blank/absent; the cell is
  unjudgeable and excluded from both denominators.
- **strict.** `exact / (total − no_oracle)` — every declined cell is a miss.
- **lenient.** `exact / (total − no_oracle − declined)` — refusals leave the
  denominator.
- **15-sig / bit-exact.** Two numeric comparison modes: equal when both
  values are rounded to 15 significant figures (Excel's cache storage
  precision), or IEEE-754 bit equality (±0 merged). Both are always
  published together.
- **external-reference cell.** A formula cell whose value depends, directly
  or by cascade, on a workbook not shipped in the corpus.

---

## Caveats

Read before relying on any number above. Each caveat states its direction
where it has one.

1. **LibreOffice booleans** are recovered via the cell's number-format
   LOGICAL bit (the UNO API otherwise exposes a boolean formula result as
   numeric 1/0). A boolean result carrying a non-logical format is scored as
   numeric and may mismatch a boolean oracle. Direction: under-credits
   LibreOffice.
2. **LibreOffice internal errors** (`Err:NNN` results with no Excel-error
   analog) are scored `declined`, not mapped onto an Excel error. The seven
   canonical Excel error strings are non-localized in LibreOffice and
   matched directly.
3. **pycel / formulas array results:** a formula returning a non-1×1 array
   is `declined` (this release scores scalar formula cells only). A 1×1
   array is unwrapped to its scalar.
4. **formulas key matching** is by upper-cased sheet name + A1 address; two
   sheets differing only in case would collide, and such a cell is
   `declined`, never mis-scored.
5. **`declined` is conservative by construction.** Whenever a runner cannot
   confidently convert an engine result into a comparable scalar Excel
   value, it declines rather than fabricate an exact or a mismatch. This can
   under-credit a competitor engine's lenient coverage; it never invents
   fidelity in either direction.
6. **The oracle is Excel's cached value** — what the last saving producer
   wrote at save time. A stale cache is stale for every engine equally, so
   cross-engine comparisons are unaffected; absolute fidelity against a live
   pinned-build recalculation is a separate question and a separate oracle
   path (see [METHODOLOGY.md](METHODOLOGY.md) §4).
7. **Versions are pinned at the top of this page** and the results describe
   those versions only. A different LibreOffice, pycel, or formulas release
   is a different measurement.
8. **Cross-mode timeout variance.** The bit-exact and 15-sig full-corpus
   funnels are independent executions, not one output re-scored. Per-workbook
   timeouts land slightly differently between runs on a loaded host: pycel's
   dead-workbook count is 37 (bit-exact) vs 34 (15-sig) and formulas' 186 vs
   190, moving their declined totals accordingly. LibreOffice's two funnels
   are execution-stable (declined 20,354, dead workbooks 5, in both), and
   Recalc's native runs are deterministic with no timeouts. Within one mode,
   every engine faced the same corpus, oracle, and conservation rules; the
   external-reference decomposition scored both tolerances from a single run
   per engine and is free of this variance.

---

## Reproducing these numbers

The measurement harness is published under Apache-2.0 in the Recalc
repository at `tools/competitors/`: oracle extraction, the cell classifier,
the funnel arithmetic, one subprocess runner per engine with per-workbook
timeouts, the external-reference intersection tool, and a test suite that
pins the classifier against the native verifier's own test vectors and
reproduces its funnel cell-for-cell on shared fixtures. The corpus is the
FUSE `.xlsx`/`.xlsm` cut defined in [METHODOLOGY.md](METHODOLOGY.md) §4,
identified by the publisher's per-file SHA-1 manifest. Engine versions as
stamped at the top of this page.

---

## Trademark and independence

Microsoft and Excel are trademarks of the Microsoft group of companies.
LibreOffice is a trademark of The Document Foundation; pycel and formulas
are their respective authors' projects. Recalc and Sheetmark are independent
projects, not affiliated with, sponsored by, or endorsed by Microsoft or any
project measured here. References to each product are nominative — they
identify the software whose measured behavior is reported.

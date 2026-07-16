# Sheetmark — Fidelity Benchmark Methodology

Sheetmark measures **how faithfully a spreadsheet calculation engine reproduces
Microsoft Excel's computed results**, cell by cell, on a fixed public corpus,
against a documented tolerance. It exists to make one claim defensible under a
skeptical audit:

> Every cell an engine reports is either **exactly** what Excel computes (at a
> documented tolerance), or it is **loudly flagged** as not attempted — never a
> silent wrong answer.

This document is the canonical method. It defines what is measured, how each cell
is classified, the two metrics and their denominators, the corpus and the oracle,
the tolerance regime, the declared scope, reproducibility, and the honest caveats.
The dated measured result set lives in [RESULTS.md](RESULTS.md).

---

## 0. Purpose and the two principles it rests on

Two product principles govern the whole methodology and are not negotiable inside
it:

1. **Fidelity is a measured number, not an adjective.** "Excel-faithful" means
   *measured agreement with Excel's computed results on a public conformance
   corpus* — nothing else. Text, logical, and error values must match exactly;
   floating-point results match at a documented tolerance (§3).
2. **Never silently wrong.** An engine that declines a cell it cannot compute —
   returning a distinguishable sentinel such as `#UNSUPPORTED!` — is behaving
   correctly and is scored as *not attempted*, not as a fidelity failure. A
   benchmark that fails loudly and covers less beats one that guesses and claims
   more.

The credibility of every number here depends on **never touching the measurement
to flatter it** (§5). That discipline *is* the product.

---

## 1. What Sheetmark measures

For each formula cell in the corpus, Sheetmark compares the engine's recomputed
value against the **oracle value** — Excel's own computed value for that cell —
and classifies the pair. The comparison is defined by type:

| Oracle vs. computed | Rule | Result |
|---|---|---|
| **Text** | Byte-for-byte, **case-sensitive** string equality (not Excel's case-folding `=` operator: Excel cached the literal string it wrote, so `"Total"` vs `"TOTAL"` is a real capitalization bug) | Exact / Mismatch |
| **Logical (bool)** | Exact equality | Exact / Mismatch |
| **Error** | Exact error-kind equality (`#DIV/0!` vs `#VALUE!` is a mismatch) | Exact / Mismatch |
| **Number** | Bit-exact (IEEE `==`, which also distinguishes `+0.0`/`-0.0`), **or** equal at the documented 15-significant-figure tolerance (§3) | Exact / UlpDiff / Mismatch |
| **Blank** | Both blank | Exact |

Only formula cells are scored — literal cells are inputs, not predictions. Every
scored cell lands in exactly one of five statuses:

- **Exact** — agrees, per the table above.
- **UlpDiff** — both numbers, within a configured ULP tolerance (default `0`, so
  absent a human-gated ULP entry this bucket is empty; see §3).
- **Mismatch** — a genuine disagreement: the only status that counts as a fidelity
  *failure* and the only one that makes a workbook run "wrong."
- **EngineUnsupported** — the engine returned one of its own honest sentinels
  (`#UNSUPPORTED!` / `#BLOCKED!` / `#RESOURCE!`): a declared refusal, not a value
  to score. This is the "never silently wrong" principle made mechanical.
- **NoOracle** — the oracle has no value for this cell at all; unscoreable, and
  reporting anything else would be a fabricated pass.

**Classification order matters.** NoOracle wins first (nothing to judge against),
then EngineUnsupported (an honest refusal is never scored against the oracle),
then the by-type comparison. This ordering is what prevents the honest-gap bucket
from ever masquerading as either a pass or a failure.

---

## 2. The two metrics — defined by their denominators

Sheetmark reports **two** fidelity percentages, never one. Both share the same
numerator (cells that agree) and differ only in what they divide by:

```
numerator = Exact + UlpDiff

Lenient % = numerator / (total − EngineUnsupported − NoOracle) × 100
Strict  % = numerator / (total                       − NoOracle) × 100
```

- **Lenient** excludes both honest-gap buckets (declared refusals *and* no-oracle
  cells). It answers: **"of the cells we could judge and the engine actually
  attempted, how many matched Excel?"** → the **engine-quality signal**.
- **Strict** puts every `#UNSUPPORTED!` back into the denominator as a miss. It
  answers: **"of every cell we had an oracle for, how many did the engine get
  right — including the ones it flatly refused?"** → **quality × coverage**.

`NoOracle` cells are excluded from *both* denominators: there is nothing to be
right or wrong about without an expected value. Neither percentage is defined when
its denominator is zero.

### The measured funnel (Recalc, v1)

Measured on the FUSE `.xlsx`/`.xlsm` cut — 3,640 workbooks, 5,667,851 formula
cells with an oracle value — engine build `66fa5f9`, re-scored 2026-07-15 (post
shared-formula expansion; zero load failures). Every cell lands in exactly one
bucket:

- **Computed and matched: 4,305,984 cells — 75.97% of the corpus.** Matched
  Excel's stored result at the documented 15-significant-figure tolerance.
- **Declined loudly: 1,313,235 cells — 23.17% of the corpus.** Flagged
  `#UNSUPPORTED!`, cell by cell, never guessed.
- **Genuine disagreement: 48,632 cells — 0.86% of the corpus.**

| Metric | 15-sig headline | Bit-exact floor |
|---|---|---|
| **Engine fidelity** (lenient — on attempted cells) | **98.883%** (= 4,305,984 / 4,354,616) | 97.856% |
| **Coverage-inclusive** (strict — over all oracle cells) | **75.972%** (= 4,305,984 / 5,667,851) | 75.183% |
| Mismatch cells | 48,632 | 93,373 |

Attempted cells (the lenient denominator, 4,354,616) are **76.8%** of the
5,667,851 oracle cells; every other cell is declined loudly, per cell, never
guessed. Figures publish **verbatim, unrounded**, always strict + lenient
together, the 15-sig figure always with its bit-exact floor.

### The declined bucket — by cause

After shared-formula expansion, the remaining declined bucket (1,313,235 cells,
23.17% of the corpus) is dominated by **corpus composition, not engine
correctness**. Its measured full-corpus decomposition:

- **External-workbook links — 73.5% (965,848 cells).** Cells that reference *other
  workbooks the benchmark does not ship*. Recalc returns `#UNSUPPORTED!`
  **correctly**: "no network / no filesystem from formulas" is a hard engine rule,
  and external-workbook resolution is an explicit non-goal (§5). This is the single
  dominant cause of the residual strict gap.
- **Shared-formula residual — 16.3% (213,468 cells).** Shared groups whose master
  formula cannot yet be parsed, or whose master is absent, so the follow-ons can't
  be expanded (down from 2,873,625 before shared-formula expansion landed — a 92.6%
  reduction; the remainder is a smaller follow-up lever, safely declined, never
  guessed).
- **Unimplemented functions — 2.8% (37,168 cells)**, led by `HYPERLINK`, `ERFC`,
  `RANK`; and **7.3% (96,263) other** (structured references, parse errors, and
  reference-form constructs). Volatile / UDF / blocked-I/O: <0.1%.

So the residual `strict` gap (≈24%) is primarily *"the benchmark does not ship the
linked workbooks"* — **not** *"the engine got the math wrong."* Strict must never
be presented as an engine-quality failing without this context. **Lenient is the
engine-quality signal; strict is quality × coverage**, bounded by what the corpus
asks of the engine's declared scope.

### Why the decomposition, and not a lone percentage

The first full decomposition of the declined bucket is *what surfaced the biggest
gap*: it found that **65% of all declined cells were unexpanded shared-formula
follow-on cells** — an addressable engine gap, not a scope boundary. Implementing
shared-formula expansion (ECMA-376 §18.17.2) moved ~3.10M cells from declined to
attempted; **3,072,264 (99.14%) of them match Excel's own cached values**, lifting
strict from **21.8% to 76.0%** while lenient held (and rose slightly). Expanding
those cells also *surfaced* 26,639 previously-hidden mismatches — they were
declined before, so uncounted — now honestly counted in the 48,632 headline
mismatch total.

This is the point of decomposing rather than reporting a single number: the
decomposition is what finds the gaps, and the number published is always the one
*after* the honest fix — never the flattering-by-omission one before it.

Results are also reported at **per-function** and **per-workbook** granularity, so
a reader can see where *their* functions and *their* workbook shapes land rather
than trusting an aggregate.

---

## 3. Tolerance

Text, logical, and error values require **exact** match — there is no tolerance
for these types.

For **numbers**, the governing rule is the human-gated **"Global numeric
comparison (15 significant figures)"** entry in the tolerance table. Two numbers
count as Exact iff they are **equal when both are rounded to 15 significant
decimal figures** (round both operands to 15 sig figs, then compare).

**Justification — storage precision, not "operators fuzz."** Excel writes its
cached values at its fundamental **15-significant-digit** storage limit (corpus
forensics: the cache holds `49.3`, never `49.300000000000004`). An engine that
keeps full IEEE-754 doubles will differ from that cache below the 15th significant
figure — and that difference is **Excel's own output-rounding artifact**, not a
computational disagreement. Bit-exact comparison scored that sub-storage-resolution
noise as a mismatch; comparing at 15 sig figs reconstructs the only question the
cache can actually answer. This is confirmed by pinned semantic-probe experiments
run against the pinned Excel build (§4) and by direct cache forensics.

**This is not weakening-to-pass.** 15-sig equivalence merges only values agreeing
to ≲ 1e-15 relative — below Excel's storage resolution — so no *detectable*
disagreement is hidden. A genuine bug diverges by ≫ 15 sig figs and **stays a
Mismatch**. Evidence from the re-score: the structural mismatch clusters (SUM / IF
/ VLOOKUP / ISBLANK-array-eval) survived largely intact — the ISBLANK-array-eval
cluster (1,876 cells) was **unchanged** (a genuine array-eval cluster, not float
noise), and the tolerance only cleared float-accumulation noise (an AVERAGE cluster
500 → 32, plus a small ~106-cell IF float-noise subcluster, 2,206 → 2,100 genuine
IF mismatches), leaving every substantive disagreement in place. Tie-boundary
rounding can only produce *false Mismatches* (over-strict), never false passes.

**Dual disclosure is mandatory.** The verifier's zero-config default stays
**bit-exact** as a regression/drift diagnostic. The 15-sig comparison is the
documented tolerance and the basis for the published headline, **but the bit-exact
floor MUST be published alongside it** as the conservative floor (the 97.856%
lenient / 75.183% strict bit-exact column in the §2 table). Publishing a lone
15-sig number is prohibited.

**Per-function absolute tolerances** exist for iterative root-finders (IRR and
XIRR: absolute 1e-7 on the returned rate) because Excel's own result there is only
accurate to a documented ceiling and is guess-dependent — Excel returns different
values for the same root depending on the starting guess, so agreement tighter than
Excel's own reproducibility is unattainable. These absolute tolerances are **not
yet mechanically enforced** by the harness (which currently exposes a single global
ULP knob plus the 15-sig flag); until per-function absolute support lands, IRR/XIRR
corpus cells score as honest `Mismatch` under strict — disclosed, not silently
tolerated.

**Tolerance changes are human-gated.** Every tolerance is a human decision with a
rationale and oracle evidence. Tightening a tolerance (making it stricter) is
always allowed; weakening one is a logged human decision. Loosening a tolerance to
pass a test is a hard-rule violation.

**Caveat (honesty).** The 15-sig *storage-precision* rationale is **Excel-specific**.
A corpus workbook last saved by a non-Excel producer (for example LibreOffice or
openpyxl) may carry a non-15-sig cached value; 15 sig figs remains a defensible
float tolerance there, but the storage-precision argument does not literally apply.
This connects to the oracle-provenance discussion in §4.

---

## 4. The corpus and the oracle

### The corpus

The measurement corpus is **FUSE** — Barik, Ford, Murphy-Hill, Zimmermann, *"FUSE:
a reproducible, extendable, internet-scale corpus of spreadsheets"* (MSR 2015): a
public, web-crawled spreadsheet corpus. Sheetmark pins the verified source artifact
and the publisher's per-file SHA-1 manifest (`fuse-all.sha1.sorted-dec2014.txt`).
The SHA-1 manifest is the corpus's **canonical identity and integrity check**,
provided source-side.

FUSE is mixed-format. Only the **`.xlsx`/`.xlsm` subset is scoreable via the
cached-value path today** (the legacy `.xls` majority requires a licensed Excel
build once `.xls` container support lands). The measured numbers in this document
are over **3,640 FUSE workbooks = 5,739,091 formula cells** (strict denominator
5,667,851 cells with an oracle), engine build `66fa5f9` (2026-07-15, post
shared-formula expansion).

**How the 3,640 cut is defined.** The scoreable corpus is built by a deterministic
content filter over the byte-length-verified source tarball (7,325,296,301 bytes,
matching the pinned source): of the **249,376** FUSE binaries, a workbook is kept
**iff** it is an OOXML container (a `xl/workbook.xml` part is present) that carries
**at least one formula** (an `<f>` element in a worksheet part). That yields exactly
**3,640** workbooks — every formula-bearing `.xlsx`/`.xlsm` in FUSE, nothing pruned
by how it scores. The remaining 245,736 binaries are legacy `.xls`,
non-spreadsheet, or formula-free (inputs only), and are correctly outside the
cached-value cut. **The selection is by format and content, never by score** — the
legitimate-scoping side of the §5 line.

A second corpus — the **Enron spreadsheet corpus** (Hermans & Murphy-Hill, figshare
record 1221767, *"Enron Spreadsheets and Emails"*, CC BY 4.0 — the spreadsheet
archive, **not** the well-known Enron *email* dataset) — is used for profiling only.
It is almost entirely legacy `.xls` and therefore requires a licensed Excel build to
oracle; it is **never a Sheetmark headline candidate**, because it flatters the
lenient metric (pre-2007 `.xls`, a narrow function vocabulary that never exercises
the harder modern functions FUSE does). Sheetmark publishes on FUSE.

**Freezing / versioning.** The corpus is identified two complementary ways — the
publisher per-file SHA-1 manifest (source-side canonical identity) and, for the
scoreable cut, the source-artifact byte-length pin plus the deterministic
formula-bearing-OOXML extraction that yields the frozen 3,640. The exact corpus cut
backing the published Sheetmark is the FUSE 3,640 `.xlsx`/`.xlsm` cached-value cut,
frozen **on principle, before the payoff was measured** (§5).

### The oracle — two paths, and the integrity discipline

There are two oracle mechanisms, and the distinction is load-bearing for
credibility:

1. **Cached-value oracle (the current headline path).** Every real `.xlsx`/`.xlsm`
   already carries Excel's own last-computed value in each formula cell's `<v>`
   element. The harness snapshots that cached value **before** handing the workbook
   to the engine (the engine's recalculation overwrites the value store), then diffs
   the engine's recomputed value against it. A formula cell whose cached value is
   blank is treated as **NoOracle**, never a fabricated pass, because a
   genuinely-computed Excel formula always writes a value. This path needs **no
   oracle dependency**. Its honest limitation: the cached value was written by
   *whatever producer last saved each workbook*, which for a web-crawled corpus is
   heterogeneous (Excel of various vintages, and sometimes non-Excel producers — see
   the §3 caveat). It is Excel's own answer where Excel saved the file, which is the
   common case, but it is **not** a fresh dump from a single pinned build.

2. **Sidecar oracle (the pinned-build ground truth).** The authoritative oracle is a
   sidecar file dumped from a **pinned, licensed Excel build**. This is the oracle
   for every probe/experiment workbook (the semantic experiments that pin behavior —
   for example the 15-sig tolerance) and is the intended path for an authoritative
   published corpus re-dump. It fills the same lookup seam the cached-value source
   already implements, and awaits a human-approved dependency for its file format.

**The pinned Excel build.** The semantic probe experiments run against **Microsoft
365, build 16.0.20131** (full dynamic-array / `LET` / `LAMBDA` support), and every
sidecar records the exact build string it was dumped under, plus calc settings, date
system, and dump timestamp. To be explicit about what backs what: the **v1
cached-value headline is not backed by a single pinned build** — its oracle is each
workbook's own producer-heterogeneous cached value (above). Build **16.0.20131** is
the pin for the **semantic probe experiments** only. Naming a single authoritative
build string for a *future* farm re-dump (the sidecar path) is deferred with that
path and does **not** gate the v1 headline.

**Read-only, hash-verified sidecars — the integrity discipline that makes the
number trustworthy:**

- **Read-only.** Sidecars are regenerated *only* against genuine Excel on the
  pinned build. No process and no person may hand-edit a sidecar's rows or metadata.
  **Editing a sidecar (or any expected value) to make a test pass is the project's
  cardinal sin.**
- **Hash-verified.** Every sidecar embeds the SHA-256 of the exact workbook bytes it
  was dumped from, and the filename carries the first 12 hex chars. **CI recomputes
  the workbook hash and fails the build on any mismatch** — a stale or hand-edited
  sidecar is *structurally* detectable, not a matter of trust.
- **One sidecar per workbook per dump; provenance required.** Re-dumping mints a new
  file (new hash); metadata must answer "which build, which calc settings, when"
  with no out-of-band information.
- **Never redistributed.** The oracle corpus and sidecars are proprietary; only the
  format spec and the harness code are open. The benchmark harness is Apache-2.0.

This hash-verification loop is the concrete mechanism behind the trust claim: a
published Sheetmark number cannot have been produced by quietly adjusting what it
was measured against.

---

## 5. Scope — declared, frozen, and reported separately

Sheetmark measures **self-contained workbook recalculation**: given a workbook's
inputs and formulas, does the engine reproduce Excel's computed values. The
following are **declared out of scope** and are engine non-goals:

- **External-workbook resolution** — formulas that reference other workbooks not
  shipped with the corpus. (The dominant driver of the strict gap, §2.)
- **VBA / macro execution** — macro code is never executed (parsed for dependency
  extraction only). The presence of a VBA project is reported as a workbook flag,
  not run.
- **Add-ins / UDFs** — no third-party or user-defined function code is executed.
- **Blocked I/O functions** — `WEBSERVICE` / `RTD` / `STOCKHISTORY` return
  `#BLOCKED!` by design (no network/filesystem from formulas).

**How out-of-scope cells are reported.** A cell the engine correctly declines is
classified `EngineUnsupported` and reported on a **separate "not-attempted" line** —
excluded from the lenient denominator, and **counted in strict** so coverage is
never hidden. An out-of-scope cell is **never** a fidelity *failure*; a `Mismatch`
is reserved for a genuine disagreement on a cell the engine *did* attempt. This is
precisely the funnel the reporting exposes: *computed exactly / correctly declined /
genuine residual.*

**The hard line — scope is frozen on principle, never tuned to flatter the number.**
This is the difference between legitimate scoping and gaming, and the benchmark's
credibility depends on staying on the right side of it:

- **Gaming (forbidden):** dropping corpus workbooks *because they score badly* to
  inflate a headline. This is adjacent to the cardinal sin and it dissolves the exact
  credibility the number is meant to buy.
- **Scoping (legitimate):** defining the conformance corpus and scope to match the
  product's *declared* scope (self-contained workbooks; no external-workbook
  resolution, no VBA/add-in execution), **decided on principle and frozen before
  seeing the payoff**, and reporting out-of-scope cells as a separate not-attempted
  line rather than as failures. This is standard conformance practice and honest
  *only* under that discipline.

If a strict headline that is not misleadingly pessimistic is ever wanted, the only
legitimate routes are (a) a frozen in-scope corpus cut defined and frozen before
measuring its payoff, or (b) raising real coverage — **never** pruning hard cases.

---

## 6. Reporting and reproducibility

- **Two numbers, always together.** Every published figure carries lenient + strict,
  and the 15-sig headline carries its bit-exact floor (§3). No lone percentage.
- **Granularity.** Results are reported at **per-function** and **per-workbook**
  granularity, not only as an aggregate. Per-function fidelity lets a prospect see
  *their* functions; per-workbook fidelity lets them see *their* workbook shapes. The
  aggregate hides that most functions and workbooks are individually at 100% or
  cleanly flagged. Machine-readable per-cell status tags —
  `exact` / `ulp_diff` / `mismatch` / `engine_unsupported` / `no_oracle` — back every
  report.
- **Versioned result sets.** Each full re-score is a result set stamped with its
  engine build (the funnel numbers in §2 are from build `66fa5f9`, 2026-07-15). A
  published number always names its engine build, the corpus freeze, the **oracle
  path** (cached-value vs. a pinned-build re-dump — §4), and the tolerance mode.
- **Engine provenance.** Every report records the engine version and short build
  hash, so any number traces to the exact build that produced it.
- **Reproducibility.** The engine recalculates single-threaded and
  deterministically (and, under an optional parallel feature, is bit-identical to
  serial by construction). Given the frozen corpus (SHA-1 manifest), the oracle
  path, the tolerance mode, and the pinned engine build, a run is reproducible. For
  the current cached-value headline the oracle values are each workbook's own
  embedded cached value, fixed by the frozen corpus bytes; for the pinned-build path
  (probe experiments today, a future authoritative re-dump) the sidecar
  hash-verification loop (§4) additionally guarantees the expected values are the
  ones the pinned Excel build actually produced.

---

## 7. Honest caveats

None of these are excuses; they are disclosed because a benchmark that hides them is
not measuring what it claims to.

- **Excel does not reproduce itself bit-for-bit.** Agreement with Excel's own
  computed results at a *documented* tolerance is **stricter than Excel meets
  itself** — Excel produces subtly different floating-point results across versions
  and platforms. A number measured that strictly is a strong claim, not a weak one,
  and it is also why a clean 100% is not a realistic target: part of any residual is
  Excel's own non-determinism, not an engine bug.
- **Stale `.xls`-conversion caches.** For some workbooks the cached value is the
  product of a pre-2007 conversion and no longer matches what Excel would recompute
  today. Where that happens, the engine's recomputed value is arguably the *more*
  correct answer, yet it scores as a disagreement against the stale cache. So a
  meaningful share of the genuine-disagreement residual is the oracle being stale,
  not the engine being wrong.
- **Precision-as-displayed workbooks.** A workbook saved with "precision as
  displayed" stores a rounded-to-shown value; the engine keeps full precision. That
  is a disagreement in bookkeeping, not in math, and it too lands in the residual.
- **Producer heterogeneity of the cached-value oracle.** The v1 cached-value oracle
  is whatever producer last saved each workbook (§4). It is Excel's own answer in the
  common case, but not a fresh dump from a single pinned build — which is exactly why
  the sidecar path (a pinned-build re-dump) is the intended authoritative oracle, and
  why the tolerance's storage-precision rationale is disclosed as Excel-specific (§3).

Taken together: the published residual (0.86% of the corpus) is an *upper bound* on
genuine engine disagreement, and a real share of it is Excel's own stored-value
quirks rather than the engine's math.

---

## Trademark and independence

Microsoft and Excel are trademarks of the Microsoft group of companies. Recalc and
Sheetmark are independent projects, not affiliated with, sponsored by, or endorsed
by Microsoft. References to Excel are nominative — they identify the product whose
behavior Sheetmark measures compatibility against.

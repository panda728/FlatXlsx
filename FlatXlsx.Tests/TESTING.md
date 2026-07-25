# What this suite underwrites

A green run means: no outstanding claim against the guarantees listed here. It does not mean
"FlatXlsx is correct". Anything absent from this page is deliberately not covered, and code that
depends on it is depending on an accident.

## Scope — what is guaranteed, and to whom

| Suite | Guarantee | Who relies on it |
|---|---|---|
| `SerializerContractTests` | Every supported type reaches the reader as its own text, fills a fixed number of cells, and is matched one-for-one by header cells. | Anyone opening the workbook. |
| `WorkbookContractTests` | Column order and headings, number/date display formats, wrapping, auto-fit widths, filter range, row count. | The person reading the sheet. |
| `WriterHistoryTests` | Properties over *sequences* of writes: text survives, the shared-string table stays consistent as it grows, row count never changes the result, async matches sync. | Callers exporting more than one row — that is, all of them. |
| `UntrustedInputTests` | Markup, control characters, hostile sizes and lazy sources cannot corrupt the file or exhaust the process. | Teams exporting data they did not author. |
| `AmbientEnvironmentTests` | Numbers and dates are stored independently of the machine's locale. | Anyone running on a server whose culture nobody chose. |
| `MessageLocalizationTests` | Messages follow the UI culture, fall back to English when untranslated, keep their numbers, and never change the workbook. | The developer reading a failure, and the support desk reading their screenshot. |
| `OutputTargetTests` | File, stream, `IBufferWriter`, and pipe destinations behave: no seeking, no closing what the caller owns, no file when there is nothing to write. | The calling application. |
| `RegressionLedgerTests` | Defects that already reached a build stay fixed. | Everyone; this is the claim history. |

The contract suites are parameterised over `SerializerCase.All`. Adding a type there subjects it
to every clause at once — that is the intended way to extend coverage.

## Deliberately not guaranteed (AS IS)

Depending on any of these is unsupported; they may change without a test turning red.

- **The bytes of the file.** Cell encoding (shared string vs inline string), style indices, zip
  entry order and compression output are free to change. Assertions are written against values
  read back through `Support/Workbook`, never against markup.
- **The exact wording of a message.** Translations may be reworded or added. What is guaranteed
  is the language it arrives in and the facts it carries, so assertions elsewhere in the suite
  quote the numbers in a message rather than its prose.
- **How a spreadsheet application renders the result.** The suite checks what a conforming reader
  parses. Excel's own layout, fonts and column-width units are outside it; the release check for
  that is opening the file by hand.
- **Leading and trailing whitespace in a cell.** `xml:space="preserve"` is not emitted, so a
  reader may trim it.
- **`XlsxWriter`'s internals.** `ToString()`, `SharedStrings` and `ColumnMaxLength` are visible
  but are not part of what is underwritten; only whole-workbook behaviour is.
- **Throughput and allocations.** Measured by the benchmark project, not defended by this suite.
- **Thread safety of a single `XlsxWriter`.** Providers and their caches are safe to share; a
  writer is not.
- **`netstandard2.0` / `netstandard2.1` behaviour at runtime.** The suite runs on net8.0 and
  net10.0 only; the older targets are compiled but not exercised. See Conditions.

## Conditions — where the guarantee applies

Frameworks: `net8.0` and `net10.0`, on Windows and Linux (both run in CI).

The netstandard builds share all source but take different `#if` branches for text formatting and
dictionary access. Those branches are compiled on every build and reviewed by eye; they are not
executed by any test. Exercising them needs a `net481` test target, which is the obvious next
extension of the Conditions fence rather than a defect in it.

## Remedy — what happens when one goes red

Every test here blocks the build. That is affordable because the whole suite is deterministic and
finishes in well under a second; there is no category worth demoting to a warning.

Two rules keep it that way:

- **A red test is diagnosed, not adjusted.** Either the contract changed — in which case the
  expectation changes and the pull request says which guarantee moved and why — or the test was
  attached to an implementation detail, in which case the test is wrong and gets raised to the
  level of the promise. Editing an expected value to match current output is neither, and is not
  an accepted fix.
- **A flaky test is quarantined the day it is noticed**, and the guarantee is announced as
  suspended until it is repaired. A red that people learn to ignore costs more than the guarantee
  was ever worth.

Property tests print the seed that failed; reproduce with `CsCheck_Seed=<seed>` before changing
anything.

## Limit — the cost of keeping it

The suite must stay fast enough that its verdict arrives while the author is still looking at the
change: **under two seconds** for the whole run. Property tests are capped by iteration count for
this reason, and the reference reader in `Support/` is deliberately minimal — it exists to read
back what was written, not to become a second implementation worth testing on its own.

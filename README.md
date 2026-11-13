
# TeamApps Reporting

TeamApps reporting utilities:

* Design beautiful report templates with Microsoft Word, LibreOffice, OpenOffice or any RTF-Editor
* Programmatically populate report templates with data of your application
* Render the reports as Microsoft Word Document, OpenDocument, RTF or Image
* Convert between file formats using this library (DOCX, PDF and PNG only) or via REST-Service (Docker image coming soon - for now use: https://github.com/EugenMayer/docker-image-jodconverter)

### Setting up the dependencies

Current TeamApps Reporting version: [![Maven Central](https://maven-badges.herokuapp.com/maven-central/org.teamapps/teamapps-reporting/badge.svg)](https://maven-badges.herokuapp.com/maven-central/org.teamapps/teamapps-reporting)
```xml
<dependency>
  <groupId>org.teamapps</groupId>
  <artifactId>teamapps-reporting</artifactId>
  <version>1.7</version>
</dependency>
```


## Template row selection

A table may contain several template rows, such as detail rows, subtotals and totals.
For each input row, the generator selects a template row based on matching replacement
keys and the configured `TemplateRowSelectionMode`.

The known variables are the union of replacement keys supplied by **all input rows of
that table**. The generator does not discover arbitrary variables in the template and
does not impose a placeholder naming convention for this check. The `<...>` notation
in the examples is an application convention. A variable never supplied by any input
row cannot be recognized as an unresolved variable by this selection check.

A fully replaceable candidate has at least one matching key and contains none of the
known variables missing from the current input row. Extra input keys need not appear
in the template. This is a selection check, not a validation of the complete output.

| Mode | Selection | If no fully replaceable row exists |
| --- | --- | --- |
| `NO_VARIABLE_CHECK` | Original behavior: select the row with the most matching keys, ignoring unresolved variables. | Use that row. If no key matches any row, `strictMode` determines whether to abort or skip. |
| `FAIL` | Select the fully replaceable row with the most matching keys. | Abort with `ReportGenerationException`. |
| `USE_BEST_MATCH` | Prefer the fully replaceable row with the most matching keys. | Select the row with the fewest known unresolved variables, then the most matching keys. Abort if no candidate has at least one matching key. |
| `SKIP` | Select the fully replaceable row with the most matching keys. | Omit the input row. |

When selection scores are equal, the first matching template row wins.
Existing constructors and factory methods default to `NO_VARIABLE_CHECK`, preserving
selection behavior from before variable checks were introduced.

Configure the mode through the constructor or setter:

```java
import org.teamapps.reporting.builder.*;
import org.teamapps.reporting.convert.DocumentFormat;

// templateInput is an InputStream containing a DOCX template.
ReportBuilder report = new ReportBuilder(
        DocumentFormat.DOCX, templateInput, true, TemplateRowSelectionMode.FAIL);

// Alternatively, keep the existing factory API:
ReportBuilder reportFromFactory = ReportBuilder.create(
        DocumentFormat.DOCX, anotherTemplateInput, true)
        .setTemplateRowSelectionMode(TemplateRowSelectionMode.USE_BEST_MATCH);
```

`strictMode` remains a separate structural setting: a missing requested table aborts
in strict mode and is ignored otherwise. For row selection, `FAIL`, `USE_BEST_MATCH`
and `SKIP` apply independently of strict mode. Only `NO_VARIABLE_CHECK` retains the
original strict-mode handling when no row has any matching key.

### Per-row selection modes

The report-wide mode is a default. Use `addRow(mode)` to override it for one input
row, for example to require transaction rows while allowing a template to omit subtotals:

```java
ReportBuilder report = ReportBuilder.create(DocumentFormat.DOCX, templateInput)
        .setTemplateRowSelectionMode(TemplateRowSelectionMode.FAIL);
TableBuilder table = report.createTableBuilder("<transactions>");

table.addRow(TemplateRowSelectionMode.FAIL)
        .setColumnValue("<account>", "Donation")
        .setColumnValue("<amount>", "100.00");

table.addRow(TemplateRowSelectionMode.SKIP)
        .setColumnValue("<subtotal>", "Subtotal")
        .setColumnValue("<amount>", "100.00");

// No override: inherit the current report-wide mode at generation time.
table.addRow().setColumnValue("<total>", "100.00");
```

If no acceptable transaction template row exists, `FAIL` aborts generation even if
the global default is `SKIP`. A subtotal marked `SKIP` can be omitted even with a
global default of `FAIL`. `USE_BEST_MATCH` and `NO_VARIABLE_CHECK` can also be selected
per row. An explicit mode must not be null; use `addRow()` to inherit the default.

The effective row mode also determines diagnostics and the handling of a missing
row. `strictMode` still controls missing tables and the legacy no-match behavior
when the effective mode is `NO_VARIABLE_CHECK`. Diagnostic row indices continue
to refer to the original input order, including skipped rows.

Known variables still come from all input rows of the table, including optional
rows. A mode override does not define a new variable scope. Skipping a row means
no acceptable template row was selected; it does not remove unused template rows
automatically. Use `removeUnusedTemplateRow(...)` for unused template variants.

### Example: choosing between template rows

Suppose the current input row supplies `<A>`, `<B>` and `<C>`. Another input row
in the same table supplies `<D>`, so all four keys are known variables.

| Template row (in document order) | Matching keys | Known unresolved variables |
| --- | --- | --- |
| Row 1: `<A> <B> <C> <D>` | 3 | `<D>` |
| Row 2: `<A> <B>` | 2 | None |

`NO_VARIABLE_CHECK` chooses **Row 1** because it has the most matches. `<D>` remains
unresolved. `FAIL`, `USE_BEST_MATCH` and `SKIP` choose **Row 2** because it is fully
replaceable. The extra input key `<C>` is unused. None of these selections produces
a diagnostic: the legacy mode does not check unresolved variables, and the other
modes found a fully replaceable row.

Now suppose the template contains only **Row 1**:

| Mode | Result for the current input row | Diagnostic |
| --- | --- | --- |
| `NO_VARIABLE_CHECK` | Use Row 1; leave `<D>` unresolved. | None |
| `FAIL` | Abort generation. | `NO_MATCHING_TEMPLATE_ROW`, also carried by `ReportGenerationException` |
| `USE_BEST_MATCH` | Use Row 1 as a fallback; leave `<D>` unresolved. | `INCOMPLETE_TEMPLATE_ROW_USED`, with `<D>` in `unresolvedVariables()` |
| `SKIP` | Omit the input row. | `INPUT_ROW_SKIPPED` |

This also shows why `NO_VARIABLE_CHECK` differs from `USE_BEST_MATCH`: the latter
prefers a fully replaceable row even when the original search would choose another.
If no input row supplied `<D>`, the variable check would not recognize it, and all
four modes would select Row 1 in this example.

## Row selection diagnostics

Register an optional listener to collect structured diagnostics. The library does not
show UI messages or log diagnostics automatically. Existing `build()` methods still
return a `File`; no listener is required.

```java
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import org.teamapps.reporting.builder.*;
import org.teamapps.reporting.convert.DocumentFormat;

List<ReportDiagnostic> diagnostics = new ArrayList<>();
ReportBuilder report = new ReportBuilder(
        DocumentFormat.DOCX, templateInput, false,
        TemplateRowSelectionMode.USE_BEST_MATCH)
        .setDiagnosticListener(diagnostics::add);

TableBuilder items = report.createTableBuilder("<items>");
items.addRow()
        .setColumnValue("<name>", "Chocolate")
        .setColumnValue("<amount>", "13.20");
items.addRow()
        .setColumnValue("<subtotal>", "Subtotal")
        .setColumnValue("<amount>", "13.20");

try {
    File output = report.build();
    // Publish the output and collected diagnostics through your application's job result.
} catch (ReportGenerationException exception) {
    ReportDiagnostic failure = exception.getDiagnostic();
    // Mark the job as failed and expose the structured failure to the application.
    // The listener has already received this same diagnostic: do not add it twice.
}
// Other build exceptions remain possible and must be handled by the caller.
```

| Diagnostic code | Meaning |
| --- | --- |
| `INCOMPLETE_TEMPLATE_ROW_USED` | `USE_BEST_MATCH` selected a fallback containing known variables missing from the input row. |
| `INPUT_ROW_SKIPPED` | `SKIP` omitted an input row, or `NO_VARIABLE_CHECK` found no matching row with strict mode disabled. |
| `NO_MATCHING_TEMPLATE_ROW` | Selection failed and generation aborts with `ReportGenerationException`. |

Successful complete selections emit no diagnostic. `NO_VARIABLE_CHECK` performs no
additional unresolved-variable check, so it emits no incomplete-row diagnostics.
The API currently covers row-selection outcomes; it does not report missing tables,
conversion errors or other generation failures as structured diagnostics.

Each immutable `ReportDiagnostic` provides:

- `code()`: the machine-readable outcome.
- `tableKeys()`: the keys used to locate the table.
- `inputRowIndex()`: the zero-based index within that table's input rows, including
  header, subtotal or other rows supplied by the application.
- `replacementKeys()`: the keys supplied for the affected input row, without values.
- `unresolvedVariables()`: known variables present in the selected template row but
  absent from the input row. Empty when no row was selected; this does not mean the
  entire template is valid or contains no unresolved placeholders.

A diagnostic is emitted for each affected input row, including repeated rows using
a cached template selection. The listener is called synchronously on the generation
thread. In a background job, collect diagnostics locally and pass them to the UI
through the application's normal job completion mechanism. Accessing the collection
concurrently requires synchronization. Exceptions thrown by the listener propagate
and abort generation. A failed build can already have emitted earlier diagnostics;
these do not indicate that a complete report was produced.

## License

The TeamApps Reporting is released under version 2.0 of the [Apache License](https://www.apache.org/licenses/LICENSE-2.0).

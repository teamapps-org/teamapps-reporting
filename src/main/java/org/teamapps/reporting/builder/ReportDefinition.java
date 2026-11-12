package org.teamapps.reporting.builder;

import org.teamapps.reporting.convert.DocumentFormat;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** An immutable snapshot of the prepared report, independent of its template and output.
 * Captures supplied values and requests, not rendered results or template validation.
 * Table and removal-list order is preserved; replacement-map iteration order is not
 * an execution-order guarantee. Template streams, output files and listeners are excluded.
 */
public record ReportDefinition(DocumentFormat inputFormat, boolean strictMode,
                               Map<String, String> replacements,
                               List<TableDefinition> tables,
                               List<List<String>> tablesToRemove,
                               List<String> paragraphsToRemove) {
    public ReportDefinition {
        replacements = Collections.unmodifiableMap(new LinkedHashMap<>(replacements));
        tables = List.copyOf(tables);
        tablesToRemove = tablesToRemove.stream()
                .map(keys -> Collections.unmodifiableList(new ArrayList<>(keys)))
                .toList();
        paragraphsToRemove = Collections.unmodifiableList(new ArrayList<>(paragraphsToRemove));
    }
}

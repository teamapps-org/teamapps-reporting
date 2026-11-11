package org.teamapps.reporting.builder;

import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Immutable row-selection diagnostic, without replacement values.
 * inputRowIndex is zero-based within the table's input rows.
 * unresolvedVariables contains known keys present in the selected template row but absent
 * from the input row. It is empty when no template row was selected; it does not imply
 * that an arbitrary template has been validated for unknown placeholders.
 */
public record ReportDiagnostic(DiagnosticCode code, List<String> tableKeys, int inputRowIndex,
                               Set<String> replacementKeys, Set<String> unresolvedVariables) {
    public ReportDiagnostic {
        Objects.requireNonNull(code, "code");
        tableKeys = List.copyOf(tableKeys);
        replacementKeys = Set.copyOf(replacementKeys);
        unresolvedVariables = Set.copyOf(unresolvedVariables);
    }
}

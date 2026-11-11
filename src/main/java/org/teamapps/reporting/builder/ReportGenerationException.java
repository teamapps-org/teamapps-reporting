package org.teamapps.reporting.builder;

import java.util.Objects;

/** Generation aborted because no acceptable template row could be selected. */
public class ReportGenerationException extends Exception {
    private final ReportDiagnostic diagnostic;

    public ReportGenerationException(ReportDiagnostic diagnostic) {
        super("Error: missing template row for keys " + String.join(", ", diagnostic.replacementKeys())
                + " in table " + diagnostic.tableKeys() + " at input row " + diagnostic.inputRowIndex());
        this.diagnostic = Objects.requireNonNull(diagnostic, "diagnostic");
    }

    public ReportDiagnostic getDiagnostic() {
        return diagnostic;
    }
}

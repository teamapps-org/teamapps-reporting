package org.teamapps.reporting.builder;

/** Controls template row selection independently of structural strict-mode checks.
 * Variables are the replacement keys supplied across the input rows, regardless of syntax.
 */
public enum TemplateRowSelectionMode {
    /** Original selection: most matching keys, retaining the first row on ties. */
    NO_VARIABLE_CHECK,
    /** Require a fully replaceable row, otherwise fail generation. */
    FAIL,
    /** Prefer fully replaceable rows; otherwise minimize open variables, then maximize matches.
     * Fails if no row has any matching key.
     */
    USE_BEST_MATCH,
    /** Skip input rows without a fully replaceable template row. */
    SKIP
}

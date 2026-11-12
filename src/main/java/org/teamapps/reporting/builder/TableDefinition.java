package org.teamapps.reporting.builder;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** An immutable snapshot of a table's input rows and template-row removal requests.
 * Keys locate the template table; rows and removal requests retain their input order.
 * These are the supplied values, before template matching or replacement takes place.
 */
public record TableDefinition(List<String> keys, boolean copyTable,
                              List<Map<String, String>> rows,
                              List<List<String>> templateRowsToRemove) {
    public TableDefinition {
        keys = Collections.unmodifiableList(new ArrayList<>(keys));
        // Unlike Map.copyOf, this also preserves null replacement values for inspection.
        rows = rows.stream()
                .map(row -> Collections.unmodifiableMap(new LinkedHashMap<>(row)))
                .toList();
        templateRowsToRemove = templateRowsToRemove.stream()
                .map(rowKeys -> Collections.unmodifiableList(new ArrayList<>(rowKeys)))
                .toList();
    }
}

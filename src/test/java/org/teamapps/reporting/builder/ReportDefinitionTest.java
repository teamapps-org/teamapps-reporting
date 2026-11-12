package org.teamapps.reporting.builder;

import org.junit.Test;
import org.teamapps.reporting.convert.DocumentFormat;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.*;

public class ReportDefinitionTest {
    private ReportBuilder report() {
        return new ReportBuilder(DocumentFormat.DOCX, new InputStream() {
            @Override public int read() {
                throw new AssertionError("Snapshots must not read the template");
            }
        }, true);
    }

    @Test public void capturesCompletePreparedStateInInputOrder() {
        ReportBuilder builder = report();
        builder.addReplacement("title", "Annual report");
        builder.addReplacement("title", "Final annual report");
        TableBuilder first = builder.createTableBuilder(List.of("income", "account"));
        first.addRow().setColumnValue("name", "Donation").setColumnValue("amount", "100.00");
        first.addRow().setColumnValue("subtotal", "Total").setColumnValue("amount", "100.00");
        first.removeUnusedTemplateRow("name", "amount");
        first.removeUnusedTemplateRow("subtotal");
        builder.copyTable("rates").addRow().setColumnValue("currency", "GBP");
        builder.removeTable("unused", "section");
        builder.removeTable("other");
        builder.removeParagraph("draft");
        builder.removeParagraph("notes");

        ReportDefinition expected = new ReportDefinition(DocumentFormat.DOCX, true,
                Map.of("title", "Final annual report"), List.of(
                        new TableDefinition(List.of("income", "account"), false,
                                List.of(Map.of("name", "Donation", "amount", "100.00"),
                                        Map.of("subtotal", "Total", "amount", "100.00")),
                                List.of(List.of("name", "amount"), List.of("subtotal"))),
                        new TableDefinition(List.of("rates"), true,
                                List.of(Map.of("currency", "GBP")), List.of())),
                List.of(List.of("unused", "section"), List.of("other")), List.of("draft", "notes"));
        assertEquals(expected, builder.getDefinition());
        assertEquals(expected.tables().get(0), first.getDefinition());
    }

    @Test public void snapshotIsDetachedFromAllBuilderCollectionsAndCallerArrays() {
        ReportBuilder builder = report();
        String[] tableKeys = {"table"};
        String[] removalKeys = {"unused"};
        String[] rowRemovalKeys = {"row"};
        builder.addReplacement("title", "Before");
        TableBuilder table = builder.createTableBuilder(tableKeys);
        RowBuilder row = table.addRow().setColumnValue("amount", "10");
        table.removeUnusedTemplateRow(rowRemovalKeys);
        builder.removeTable(removalKeys);
        builder.removeParagraph("draft");
        ReportDefinition before = builder.getDefinition();

        tableKeys[0] = "changed table";
        removalKeys[0] = "changed removal";
        rowRemovalKeys[0] = "changed row removal";
        row.getRowMap().put("amount", "20");
        table.addRow().setColumnValue("amount", "30");
        table.removeUnusedTemplateRow("another row");
        builder.addReplacement("title", "After");
        builder.copyTable("another table");
        builder.removeTable("another removal");
        builder.removeParagraph("another paragraph");

        assertEquals(new ReportDefinition(DocumentFormat.DOCX, true,
                Map.of("title", "Before"), List.of(new TableDefinition(List.of("table"), false,
                List.of(Map.of("amount", "10")), List.of(List.of("row")))),
                List.of(List.of("unused")), List.of("draft")), before);
        ReportDefinition after = builder.getDefinition();
        assertEquals("After", after.replacements().get("title"));
        assertEquals(2, after.tables().size());
        assertEquals(2, after.tables().get(0).rows().size());
        assertEquals("20", after.tables().get(0).rows().get(0).get("amount"));
    }

    @Test public void everyExposedCollectionIsUnmodifiable() {
        ReportBuilder builder = report();
        builder.addReplacement("title", "Report");
        TableBuilder table = builder.createTableBuilder("table");
        table.addRow().setColumnValue("amount", "10");
        table.removeUnusedTemplateRow("row");
        builder.removeTable("unused");
        builder.removeParagraph("draft");
        ReportDefinition definition = builder.getDefinition();
        TableDefinition tableDefinition = definition.tables().get(0);

        assertThrows(UnsupportedOperationException.class, () -> definition.replacements().put("title", "Changed"));
        assertThrows(UnsupportedOperationException.class, () -> definition.tables().clear());
        assertThrows(UnsupportedOperationException.class, () -> definition.tablesToRemove().clear());
        assertThrows(UnsupportedOperationException.class, () -> definition.tablesToRemove().get(0).clear());
        assertThrows(UnsupportedOperationException.class, () -> definition.paragraphsToRemove().clear());
        assertThrows(UnsupportedOperationException.class, () -> tableDefinition.keys().clear());
        assertThrows(UnsupportedOperationException.class, () -> tableDefinition.rows().clear());
        assertThrows(UnsupportedOperationException.class, () -> tableDefinition.rows().get(0).put("amount", "20"));
        assertThrows(UnsupportedOperationException.class, () -> tableDefinition.templateRowsToRemove().clear());
        assertThrows(UnsupportedOperationException.class, () -> tableDefinition.templateRowsToRemove().get(0).clear());
    }

    @Test public void recordConstructorsAlsoCopyTheirInputsAndPreserveNullValues() {
        Map<String, String> replacements = new HashMap<>();
        replacements.put("optional", null);
        Map<String, String> row = new HashMap<>(replacements);
        List<String> keys = new ArrayList<>(List.of("table"));
        List<Map<String, String>> rows = new ArrayList<>(List.of(row));
        List<List<String>> removals = new ArrayList<>(List.of(keys));
        TableDefinition table = new TableDefinition(keys, false, rows, removals);
        List<TableDefinition> tables = new ArrayList<>(List.of(table));
        ReportDefinition definition = new ReportDefinition(DocumentFormat.DOCX, false, replacements, tables, removals, keys);

        row.put("optional", "Changed");
        replacements.clear();
        keys.clear();
        rows.clear();
        removals.clear();
        tables.clear();
        assertTrue(definition.replacements().containsKey("optional"));
        assertNull(definition.replacements().get("optional"));
        assertTrue(table.rows().get(0).containsKey("optional"));
        assertNull(table.rows().get(0).get("optional"));
        assertEquals(List.of("table"), table.keys());
        assertEquals(List.of(List.of("table")), table.templateRowsToRemove());
        assertEquals(List.of(List.of("table")), definition.tablesToRemove());
        assertEquals(List.of("table"), definition.paragraphsToRemove());
        assertEquals(List.of(table), definition.tables());
    }

    @Test public void emptyBuilderCanBeInspectedWithoutATemplate() {
        ReportBuilder builder = ReportBuilder.create(DocumentFormat.DOCX, (InputStream) null);
        assertEquals(new ReportDefinition(DocumentFormat.DOCX, false,
                Map.of(), List.of(), List.of(), List.of()), builder.getDefinition());
    }
}

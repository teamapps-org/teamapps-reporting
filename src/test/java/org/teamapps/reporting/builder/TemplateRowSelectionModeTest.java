package org.teamapps.reporting.builder;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.junit.Test;
import java.util.*;
import static org.junit.Assert.*;

public class TemplateRowSelectionModeTest {
    private final DocumentBuilder builder = new DocumentBuilder();

    private WordprocessingMLPackage template(String... rows) throws Exception {
        WordprocessingMLPackage document = WordprocessingMLPackage.createPackage();
        Tbl table = new Tbl();
        document.getMainDocumentPart().addObject(table);
        for (String value : rows) {
            Tr row = new Tr();
            Tc cell = new Tc();
            P paragraph = new P();
            R run = new R();
            Text text = new Text();
            text.setValue(value);
            run.getContent().add(text);
            paragraph.getContent().add(run);
            cell.getContent().add(paragraph);
            row.getContent().add(cell);
            table.getContent().add(row);
        }
        return document;
    }

    private List<String> fill(TemplateRowSelectionMode mode, boolean strict, String... rows) throws Exception {
        WordprocessingMLPackage document = template(rows);
        // The second input declares D and E as known variables and has its own template.
        builder.fillTable(List.of(Map.of("<A>", "a", "<B>", "b", "<C>", "c"),
                        Map.of("<D>", "d", "<E>", "e")), List.of(), document, strict,
                List.of(), false, mode);
        List<Tr> result = builder.getAllElements(document.getMainDocumentPart(), new Tr());
        return result.stream().map(row -> builder.getAllElements(row, new Text()).stream()
                .map(Text::getValue).reduce("", String::concat)).toList();
    }

    @Test public void legacyKeepsMostMatchesEvenWhenCompleteAlternativeExists() throws Exception {
        assertEquals(List.of("complete <A><B>", "legacy abc<D>", "known de"),
                fill(TemplateRowSelectionMode.NO_VARIABLE_CHECK, false,
                        "legacy <A><B><C><D>", "complete <A><B>", "known <D><E>"));
    }

    @Test public void checkedModesPreferCompleteAlternative() throws Exception {
        for (TemplateRowSelectionMode mode : List.of(TemplateRowSelectionMode.FAIL,
                TemplateRowSelectionMode.SKIP, TemplateRowSelectionMode.USE_BEST_MATCH)) {
            assertEquals(List.of("legacy <A><B><C><D>", "complete ab", "known de"),
                    fill(mode, false, "legacy <A><B><C><D>", "complete <A><B>", "known <D><E>"));
        }
    }

    @Test public void failIsIndependentOfStrictMode() throws Exception {
        for (boolean strict : List.of(false, true)) {
            Exception error = assertThrows(Exception.class, () -> fill(TemplateRowSelectionMode.FAIL,
                    strict, "incomplete <A><D>", "known <D><E>"));
            assertTrue(error.getMessage().contains("missing template row"));
        }
    }

    @Test public void skipIsIndependentOfStrictMode() throws Exception {
        for (boolean strict : List.of(false, true)) {
            assertEquals(List.of("incomplete <A><D>", "known de"),
                    fill(TemplateRowSelectionMode.SKIP, strict, "incomplete <A><D>", "known <D><E>"));
        }
    }

    @Test public void fallbackMinimizesOpenVariablesBeforeMaximizingMatches() throws Exception {
        assertEquals(List.of("more <A><B><C><D><E>", "fewer a<D>", "known de"),
                fill(TemplateRowSelectionMode.USE_BEST_MATCH, false,
                        "more <A><B><C><D><E>", "fewer <A><D>", "known <D><E>"));
    }

    @Test public void fallbackRequiresAtLeastOneMatch() throws Exception {
        assertThrows(Exception.class, () -> fill(TemplateRowSelectionMode.USE_BEST_MATCH,
                false, "known <D><E>"));
    }

    @Test public void originalOverloadsKeepLegacySelectionAndFirstTie() throws Exception {
        WordprocessingMLPackage document = template("first <A><D>", "second <A>");
        Tbl table = builder.getAllElements(document.getMainDocumentPart(), new Tbl()).get(0);
        Tr first = builder.getAllElements(table, new Tr()).get(0);
        assertSame(first, builder.findBestRowInTable(table, "<A>"));
        assertSame(first, builder.findBestRowInTable(table, List.of("<A>")));
        builder.fillTable(List.of(Map.of("<A>", "a")), document, false, List.of());
        assertEquals(List.of("second <A>", "first a<D>"),
                builder.getAllElements(table, new Text()).stream().map(Text::getValue).toList());
    }
    @Test public void constructorPassesPolicyThroughReportGeneration() throws Exception {
        java.io.ByteArrayOutputStream bytes = new java.io.ByteArrayOutputStream();
        template("<A><D>", "<D><E>").save(bytes);
        for (TemplateRowSelectionMode mode : TemplateRowSelectionMode.values()) {
            ReportBuilder report = new ReportBuilder(org.teamapps.reporting.convert.DocumentFormat.DOCX,
                    new java.io.ByteArrayInputStream(bytes.toByteArray()), false, mode);
            TableBuilder table = report.createTableBuilder(List.of());
            table.addRow().setColumnValue("<A>", "a");
            table.addRow().setColumnValue("<D>", "d").setColumnValue("<E>", "e");
            if (mode == TemplateRowSelectionMode.FAIL) {
                Exception error = assertThrows(Exception.class, report::build);
                assertTrue(error.getMessage().contains("missing template row"));
            } else {
                java.io.File output = report.build();
                try {
                    WordprocessingMLPackage result = WordprocessingMLPackage.load(output);
                    List<String> rows = builder.getAllElements(result.getMainDocumentPart(), new Tr()).stream()
                            .map(row -> builder.getAllElements(row, new Text()).stream()
                                    .map(Text::getValue).reduce("", String::concat)).toList();
                    assertEquals(mode == TemplateRowSelectionMode.SKIP
                            ? List.of("<A><D>", "de") : List.of("a<D>", "de"), rows);
                } finally {
                    java.nio.file.Files.deleteIfExists(output.toPath());
                }
            }
        }
    }

}

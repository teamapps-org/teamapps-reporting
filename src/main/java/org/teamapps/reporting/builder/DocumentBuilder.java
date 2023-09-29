/*-
 * ========================LICENSE_START=================================
 * TeamApps Reporting
 * ---
 * Copyright (C) 2020 - 2023 TeamApps.org
 * ---
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * =========================LICENSE_END==================================
 */
package org.teamapps.reporting.builder;


import jakarta.xml.bind.JAXBElement;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.wml.*;
import org.jvnet.jaxb2_commons.ppp.Child;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.*;
import java.util.stream.Collectors;

public class DocumentBuilder {

    public Map<String, String> createReplaceRowMap(String... values) {
        Map<String, String> map = new HashMap<>();
        for (int i = 0; i < values.length; i += 2) {
            map.put(values[i], values[i + 1]);
        }
        return map;
    }

    public WordprocessingMLPackage getTemplate(String path) throws Docx4JException, FileNotFoundException {
        return WordprocessingMLPackage.load(new FileInputStream(new File(path)));
    }

    public <T> T copyElement(T t) {
        return XmlUtils.deepCopy(t);
    }

    public <T> List<T> getAllElements(Object element, T toSearch) {
        List<T> result = new ArrayList<>();
        if (element instanceof JAXBElement) element = ((JAXBElement<?>) element).getValue();

        if (element.getClass().equals(toSearch.getClass()))
            result.add((T) element);
        else if (element instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) element).getContent();
            for (Object child : children) {
                result.addAll(getAllElements(child, toSearch));
            }
        }
        return result;
    }

    public void save(WordprocessingMLPackage template, String target) throws Docx4JException {
        File f = new File(target);
        template.save(f);
    }

    public void fillTable(List<Map<String, String>> textToAdd, WordprocessingMLPackage template, boolean strictMode, String... keys) throws Exception {
        fillTable(textToAdd, template, strictMode, Arrays.asList(keys));
    }

    public void fillTable(List<Map<String, String>> textToAdd, WordprocessingMLPackage template, boolean strictMode, List<String> keys) throws Exception {
        fillTable(textToAdd, Collections.emptyList(), template, strictMode, keys);
    }

    public void fillTable(List<Map<String, String>> textToAdd, List<List<String>> removeTemplateRows, WordprocessingMLPackage template, boolean strictMode, List<String> keys) throws Exception {
        fillTable(textToAdd, removeTemplateRows, template, strictMode, keys, false);
    }

    public void fillTable(List<Map<String, String>> textToAdd, List<List<String>> removeTemplateRows, WordprocessingMLPackage template, boolean strictMode, List<String> keys, boolean copyTable) throws Exception {
        Tbl matchingTable = findTable(template.getMainDocumentPart(), keys);
        if (strictMode && matchingTable == null) {
            throw new Exception("Error: missing template table for keys" + String.join(", ", keys));
        }

        Map<Set<String>, Tr> templateRowByColumnsSet = new HashMap<>();
        Set<Tr> removeSet = new HashSet<>();
        if (matchingTable != null) {

            if (copyTable) {
                Tbl tableCopy = copyElement(matchingTable);
                template.getMainDocumentPart().addObject(tableCopy);
                matchingTable = tableCopy;
            }

            for (Map<String, String> replaceMap : textToAdd) {
                Tr templateRow = templateRowByColumnsSet.get(replaceMap.keySet());
                if (templateRow == null) {
                    templateRow = findBestRowInTable(matchingTable, replaceMap.keySet());
                    templateRowByColumnsSet.put(replaceMap.keySet(), templateRow);
                }
                if (templateRow == null) {
                    if (strictMode) {
                        throw new Exception("Error: missing template row for keys" + String.join(", ", replaceMap.keySet()));
                    } else {
                        continue;
                    }
                }
                removeSet.add(templateRow);
                Tr row = copyElement(templateRow);
                for (Map.Entry<String, String> entry : replaceMap.entrySet()) {
                    replaceObjectDataWithinMarkers(entry.getKey(), entry.getValue(), row);
                }
                matchingTable.getContent().add(row);
            }

            for (List<String> removeTemplateRow : removeTemplateRows) {
                Tr templateRow = findBestRowInTable(matchingTable, removeTemplateRow);
                removeSet.add(templateRow);
            }

            for (Tr tr : removeSet) {
                matchingTable.getContent().remove(tr);
            }
        }
    }

    public Tbl findTable(Object element, List<String> keys) {
        List<Tbl> tables = getAllElements(element, new Tbl());
        Tbl matchingTable = null;
        for (Tbl table : tables) {
            boolean hit = true;
            for (String key : keys) {
                if (getParagraphWithText(table, key) == null) {
                    hit = false;
                    break;
                }
            }
            if (hit) {
                matchingTable = table;
                break;
            }
        }
        return matchingTable;
    }

    public void replaceParagraph(String key, String value, Object element) {
        P paragraph = getParagraphWithText(element, key);
        if (paragraph == null) {
            return;
        }
        List<Text> texts = getAllElements(paragraph, new Text());
        texts.get(0).setValue(value);
        if (texts.size() == 1) {
            return;
        }
        for (int i = 1; i < texts.size(); i++) {
            removeChild(paragraph, texts.get(i));
        }
    }

    public void replaceParagraphTextRun(String key, String value, Object element) {
        P paragraph = getParagraphWithText(element, key);
        if (paragraph == null) {
            return;
        }
        List<Text> texts = getAllElements(paragraph, new Text());
        for (Text text : texts) {
            if (text.getValue().contains(key)) {
                String replacedTextValue = text.getValue().replace(key, value);
                text.setValue(replacedTextValue);
            }
        }
    }

    public void replaceTextRun(String key, String value, Object element) {
        List<Text> texts = getAllElements(element, new Text());
        for (Text text : texts) {
            if (text.getValue().contains(key)) {
                String replacedTextValue = text.getValue().replace(key, value);
                text.setValue(replacedTextValue);
            }
        }
    }

    private void replaceObjectDataWithinMarkers(String key, String value, Object element) {
        P paragraph = getParagraphWithText(element, key);
        replaceParagraphMarkerText(key, value, paragraph);
    }

    private void replaceTextRunWithinMarkers(String key, String value, Text text) {
        P paragraph = getParagraphOfText(text);
        replaceParagraphMarkerText(key, value, paragraph);
    }

    private void replaceParagraphMarkerText(String key, String value, P paragraph) {
        if (paragraph != null) {
            List<Text> texts = getMarkerTextRuns(paragraph, key, "<", ">");
            if (texts == null || texts.isEmpty()) {
                return;
            }
            if (texts.size() == 1) {
                Text t = texts.get(0);
                t.setValue(t.getValue().replace(key, value));
            } else if (texts.size() == 2) {
                Text start = texts.get(0);
                Text end = texts.get(1);
                boolean leftSpace = !start.getValue().trim().startsWith("<") && !end.getValue().trim().startsWith("<");
                int pos = start.getValue().lastIndexOf('<');
                start.setValue(start.getValue().substring(0, pos));
                pos = end.getValue().indexOf('>');
                setValueOrAddParagraphs(paragraph, end, value, end.getValue().substring(pos + 1));
                addSpace(end, leftSpace, false);
            } else {
                Text start = texts.get(0);
                Text mid = texts.get(1);
                Text end = texts.get(texts.size() - 1);

                boolean leftSpace = !start.getValue().trim().startsWith("<") && !mid.getValue().trim().startsWith("<");
                boolean rightSpace = !mid.getValue().trim().endsWith(">") && !end.getValue().trim().endsWith(">");

                int pos = start.getValue().lastIndexOf('<');
                start.setValue(start.getValue().substring(0, pos));
                pos = end.getValue().indexOf('>');
                end.setValue(end.getValue().substring(pos + 1));

                setValueOrAddParagraphs(paragraph, mid, value);
                addSpace(mid, leftSpace, rightSpace);
                if (texts.size() > 3) {
                    for (int i = 2; i < texts.size() - 1; i++) {
                        texts.get(i).setValue("");
                    }
                }
            }
        }
    }

    private List<Text> getMarkerTextRuns(P paragraph, String key, String startMarker, String endMarker) {
        List<Text> texts = getAllElements(paragraph, new Text());
        List<Text> resultRuns = null;
        for (Text text : texts) {
            boolean withStartMarker = false;
            if (text.getValue().contains(startMarker)) {
                resultRuns = new ArrayList<>();
                resultRuns.add(text);
                withStartMarker = true;
            }
            if (text.getValue().contains(endMarker) && resultRuns != null) {
                if (!withStartMarker) {
                    resultRuns.add(text);
                }
                if (resultRuns.stream().map(Text::getValue).collect(Collectors.joining()).contains(key)) {
                    return resultRuns;
                } else {
                    resultRuns = null;
                }
            } else if (!withStartMarker && resultRuns != null) {
                resultRuns.add(text);
            }
        }
        return null;
    }

    private void addSpace(Text text, boolean left, boolean right) {
        Object parent = text.getParent();
        if (parent instanceof R) {
            R run = (R) parent;
            if (run.getContent().size() == 1) {
                if (left) {
                    Text space = new Text();
                    space.setSpace("preserve");
                    space.setValue(" ");
                    run.getContent().add(0, space);
                }
                if (right) {
                    Text space = new Text();
                    space.setSpace("preserve");
                    space.setValue(" ");
                    run.getContent().add(space);
                }
            }
        }
    }

    public void replaceTextRunWithFootersAndHeaders(String key, String value, Object element, WordprocessingMLPackage template) {
        List<Text> texts = getAllElements(element, new Text());
        texts.addAll(getHeaderFooterTexts(template));
        for (Text text : texts) {
            if (key.startsWith("<") && key.endsWith(">")) {
                replaceTextRunWithinMarkers(key, value, text);
            } else {
                if (text.getValue().contains(key)) {
                    String replacedTextValue = text.getValue().replace(key, value);
                    text.setValue(replacedTextValue);
                }
            }
        }
    }

    private List<Text> getHeaderFooterTexts(WordprocessingMLPackage template) {
        List<SectionWrapper> sectionWrappers = template.getDocumentModel().getSections();
        List<Text> texts = new ArrayList<>();
        for (SectionWrapper sectionWrapper : sectionWrappers) {
            HeaderFooterPolicy headerFooterPolicy = sectionWrapper.getHeaderFooterPolicy();
            if (headerFooterPolicy.getDefaultHeader() != null) {
                texts.addAll(getAllElements(headerFooterPolicy.getDefaultHeader(), new Text()));
            }
            if (headerFooterPolicy.getDefaultFooter() != null) {
                texts.addAll(getAllElements(headerFooterPolicy.getDefaultFooter(), new Text()));
            }
        }
        return texts;
    }

    public void removeChild(Object parent, Object child) {
        if (parent instanceof ContentAccessor) {
            ContentAccessor contentAccessor = (ContentAccessor) parent;
            List<Object> children = contentAccessor.getContent();
            Set<Object> contentSet = new HashSet<>(children);
            removeChild(contentAccessor, contentSet, child);
        }
    }

    public void removeChild(ContentAccessor contentAccessor, Set<Object> contentSet, Object child) {
        if (contentSet.contains(child)) {
            contentAccessor.getContent().remove(child);
        } else if (child instanceof Child) {
            Child contentChild = (Child) child;
            Object parent = contentChild.getParent();
            if (parent != null) {
                removeChild(contentAccessor, contentSet, parent);
            }
        }
    }

    public Tr findBestRowInTable(Tbl table, Collection<String> keys) {
        return findBestRowInTable(table, keys.toArray(new String[0]));
    }

    public Tr findBestRowInTable(Tbl table, String... keys) {
        List<Tr> rows = getAllElements(table, new Tr());
        int bestHitScore = 0;
        Tr bestRow = null;
        for (Tr row : rows) {
            int score = 0;
            for (String key : keys) {
                if (getParagraphWithText(row, key) != null) {
                    score++;
                }
            }
            if (score > bestHitScore) {
                bestRow = row;
                bestHitScore = score;
            }
        }
        return bestRow;
    }

    public Tr findRowInTable(Tbl table, Collection<String> keys) {
        return findRowInTable(table, keys.toArray(new String[0]));
    }

    public Tr findRowInTable(Tbl table, String... keys) {
        List<Tr> rows = getAllElements(table, new Tr());
        for (Tr row : rows) {
            boolean hit = true;
            for (String key : keys) {
                if (getParagraphWithText(row, key) == null) {
                    hit = false;
                    break;
                }
            }
            if (hit) {
                return row;
            }
        }
        return null;
    }

    public P getParagraphWithText(Object element, String key) {
        List<P> paragraphs = getAllElements(element, new P());
        for (P paragraph : paragraphs) {
            List<Text> texts = getAllElements(paragraph, new Text());
            StringBuilder sb = new StringBuilder();
            texts.forEach(text -> sb.append(text.getValue()));
            if (sb.length() > 0 && sb.toString().contains(key)) {
                return paragraph;
            }
        }
        return null;
    }

    public P getParagraphOfText(Text text) {
        Object parent = text.getParent();
        if (parent != null && parent instanceof R) {
            R run = (R) parent;
            Object paragraph = run.getParent();
            if (paragraph != null && paragraph instanceof P) {
                return (P) paragraph;
            }
        }
        return null;
    }

    public void setValueOrAddParagraphs(P paragraph, Text text, String value) {
        setValueOrAddParagraphs(paragraph, text, value, null);
    }

    public void setValueOrAddParagraphs(P paragraph, Text text, String value, String firstRunLefOver) {
        if (value.contains("\n")) {
            String[] parts = value.split("\n");
            for (int i = 0; i < parts.length; i++) {
                String part = parts[i];
                if (i == 0) {
                    if (firstRunLefOver != null) {
                        text.setValue(part + firstRunLefOver);
                    } else {
                        text.setValue(part);
                    }
                } else if (!part.isBlank()) {
                    addParagraphWithText(paragraph, text, part);
                }
            }
        } else {
            if (firstRunLefOver != null) {
                text.setValue(value + firstRunLefOver);
            } else {
                text.setValue(value);
            }
        }
    }

    public void addParagraphWithText(P paragraph, Text text, String value) {
        R run = (R) text.getParent();
        addParagraphWithText(paragraph, run, value);
    }

    public void addParagraphWithText(P paragraph, R run, String value) {
        R newRun = copyElement(run);
        newRun.getContent().clear();
        Text text = new Text();
        text.setValue(value);
        newRun.getContent().add(text);
        P p = new P();
        p.getContent().add(newRun);
        if (paragraph.getParent() instanceof ContentAccessor) {
            ContentAccessor contentAccessor = (ContentAccessor) paragraph.getParent();
            addContentAfterExisting(contentAccessor, paragraph, p);
        }
    }

    public <E> void addContentAfterExisting(ContentAccessor contentAccessor, E existing, E addElement) {
        for (int i = 0; i < contentAccessor.getContent().size(); i++) {
            if (contentAccessor.getContent().get(i).equals(existing)) {
                contentAccessor.getContent().add(i + 1, addElement);
                break;
            }
        }
    }

    public void addImage(WordprocessingMLPackage template, HeaderPart part, P paragraph, byte[] bytes) throws Exception {
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(template, part, bytes);
        Inline inline = imagePart.createImageInline(null, null, 0, 1, false, 800);
        ObjectFactory factory = new ObjectFactory();
        Drawing drawing = factory.createDrawing();
        drawing.getAnchorOrInline().add(inline);
        R run = factory.createR();
        paragraph.getContent().add(run);
        run.getContent().add(drawing);
    }
}


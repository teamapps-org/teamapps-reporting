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
package org.teamapps.reporting.parse;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.teamapps.reporting.builder.DocumentBuilder;
import org.teamapps.reporting.builder.DocumentTemplateLoader;
import org.teamapps.reporting.convert.DocumentConverter;
import org.teamapps.reporting.convert.DocumentFormat;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class DocumentParser {

	private WordprocessingMLPackage template;
	private DocumentBuilder documentBuilder;

	public static DocumentParser create(DocumentFormat inputFormat, File templateFile) throws Exception {
		return create(inputFormat, new BufferedInputStream(new FileInputStream(templateFile)));
	}

	public static DocumentParser create(DocumentFormat inputFormat, File templateFile, DocumentConverter documentConverter) throws Exception {
		return create(inputFormat, new BufferedInputStream(new FileInputStream(templateFile)), documentConverter);
	}


	public static DocumentParser create(DocumentFormat inputFormat, InputStream inputStream) throws Exception {
		return new DocumentParser(inputFormat, inputStream);
	}

	public static DocumentParser create(DocumentFormat inputFormat, InputStream inputStream, DocumentConverter documentConverter) throws Exception {
		return new DocumentParser(inputFormat, inputStream, documentConverter);
	}

	private final DocumentFormat inputFormat;
	private final InputStream inputStream;
	private final DocumentConverter documentConverter;

	protected DocumentParser(DocumentFormat inputFormat, InputStream inputStream) throws Exception {
		this(inputFormat, inputStream, null);
	}

	protected DocumentParser(DocumentFormat inputFormat, InputStream inputStream, DocumentConverter documentConverter) throws Exception {
		this.inputFormat = inputFormat;
		this.inputStream = inputStream;
		this.documentConverter = documentConverter;
		this.documentBuilder = new DocumentBuilder();
		readTemplate();
	}

	private void readTemplate() throws Exception {
		InputStream documentInputStream = inputStream;
		if (inputFormat != DocumentFormat.DOCX) {
			File tempFile = File.createTempFile("temp", "." + DocumentFormat.DOCX.getFormat());
			documentConverter.convertDocument(inputStream, inputFormat, tempFile, DocumentFormat.DOCX);
			documentInputStream = new BufferedInputStream(new FileInputStream(tempFile));
		}
		template = DocumentTemplateLoader.getTemplate(documentInputStream);
	}

	public List<List<String>> readTableData(String... keys) {
		Tbl table = documentBuilder.findTable(template.getMainDocumentPart(), Arrays.asList(keys));
		return readTableData(table);
	}

	public List<List<String>> readTableData(int tableIndex) {
		List<Tbl> tables = documentBuilder.getAllElements(template.getMainDocumentPart(), new Tbl());
		if (tables.size() <= tableIndex) {
			return null;
		}
		Tbl table = tables.get(tableIndex);
		return readTableData(table);
	}

	private List<List<String>> readTableData(Tbl table) {
		List<List<String>> tableData = new ArrayList<>();
		for (Object element : table.getContent()) {
			if (element instanceof Tr) {
				Tr row = (Tr) element;
				List<String> rowData = new ArrayList<>();
				tableData.add(rowData);
				for (Object cell : row.getContent()) {
					List<Text> texts = documentBuilder.getAllElements(cell, new Text());
					String cellValue = texts.stream()
							.map(Text::getValue)
							.collect(Collectors.joining());
					rowData.add(cellValue);
				}
			}
		}
		return tableData;
	}
}

package org.teamapps.reporting.parse;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.teamapps.reporting.builder.DocumentTemplateLoader;
import org.teamapps.reporting.convert.DocumentConverter;
import org.teamapps.reporting.convert.DocumentFormat;
import org.teamapps.reporting.convert.UnsupportedFormatException;

import javax.xml.bind.JAXBElement;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class DocumentParser {

	public static DocumentParser create(DocumentFormat inputFormat, File templateFile) throws FileNotFoundException {
		return create(inputFormat, new BufferedInputStream(new FileInputStream(templateFile)));
	}

	public static DocumentParser create(DocumentFormat inputFormat, InputStream inputStream) {
		return new DocumentParser(inputFormat, inputStream);
	}

	private final DocumentFormat inputFormat;
	private final InputStream inputStream;

	protected DocumentParser(DocumentFormat inputFormat, InputStream inputStream) {
		this.inputFormat = inputFormat;
		this.inputStream = inputStream;
	}

	public List<List<String>> parseTableData(int tableIndex) throws Exception {
		if (inputFormat != DocumentFormat.DOCX) {
			UnsupportedFormatException.throwException(inputFormat);
		}
		return parseTableData(tableIndex, inputStream);
	}

	public List<List<String>> parseTableData(int tableIndex, DocumentConverter documentConverter) throws Exception {
		InputStream documentInputStream = inputStream;
		if (inputFormat != DocumentFormat.DOCX) {
			File tempFile = File.createTempFile("temp", "." + DocumentFormat.DOCX.getFormat());
			documentConverter.convertDocument(inputStream, inputFormat, tempFile, DocumentFormat.DOCX);
			documentInputStream = new BufferedInputStream(new FileInputStream(tempFile));
		}
		return parseTableData(tableIndex, documentInputStream);
	}

	private List<List<String>> parseTableData(int tableIndex, InputStream documentInputStream) throws Docx4JException {
		WordprocessingMLPackage template = DocumentTemplateLoader.getTemplate(documentInputStream);
		List<Tbl> tables = getAllElements(template.getMainDocumentPart(), new Tbl());
		if (tables.size() <= tableIndex) {
			return null;
		}
		Tbl table = tables.get(tableIndex);
		List<List<String>> tableData = new ArrayList<>();
		for (Object element : table.getContent()) {
			if (element instanceof Tr) {
				Tr row = (Tr) element;
				List<String> rowData = new ArrayList<>();
				tableData.add(rowData);
				for (Object cell : row.getContent()) {
					List<Text> texts = getAllElements(cell, new Text());
					String cellValue = texts.stream()
							.map(Text::getValue)
							.collect(Collectors.joining());
					rowData.add(cellValue);
				}
			}
		}
		return tableData;
	}

	private static <T> List<T> getAllElements(Object element, T toSearch) {
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
}

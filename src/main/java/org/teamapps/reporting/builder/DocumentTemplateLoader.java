package org.teamapps.reporting.builder;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.*;

public class DocumentTemplateLoader {

	public static WordprocessingMLPackage getTemplate(String path) throws Docx4JException, FileNotFoundException {
		return getTemplate(new File(path));
	}

	public static WordprocessingMLPackage getTemplate(File file) throws FileNotFoundException, Docx4JException {
		return getTemplate(new BufferedInputStream(new FileInputStream(file)));
	}

	public static WordprocessingMLPackage getTemplate(InputStream inputStream) throws Docx4JException {
		return WordprocessingMLPackage.load(inputStream);
	}
}

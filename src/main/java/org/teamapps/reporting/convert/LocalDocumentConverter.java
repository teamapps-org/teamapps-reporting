package org.teamapps.reporting.convert;

import org.apache.xmlgraphics.util.MimeConstants;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.teamapps.reporting.builder.DocumentTemplateLoader;

import java.io.*;

public class LocalDocumentConverter implements DocumentConverter {

	@Override
	public boolean convertDocument(File input, DocumentFormat inputFormat, File output, DocumentFormat outputFormat) throws Exception {
		if (inputFormat != DocumentFormat.DOCX) {
			UnsupportedFormatException.throwException(inputFormat);
		}
		WordprocessingMLPackage wordprocessingMLPackage = DocumentTemplateLoader.getTemplate(input);
		return convertDocument(wordprocessingMLPackage, output, outputFormat);
	}

	@Override
	public boolean convertDocument(WordprocessingMLPackage wordprocessingMLPackage, File output, DocumentFormat outputFormat) throws Exception {
		if (outputFormat != DocumentFormat.PDF) {
			UnsupportedFormatException.throwException(outputFormat);
		}
		Mapper fontMapper = new IdentityPlusMapper();
		fontMapper.getFontMappings().putAll(PhysicalFonts.getPhysicalFonts());
		wordprocessingMLPackage.setFontMapper(fontMapper);

		BufferedOutputStream outputStream = new BufferedOutputStream(new FileOutputStream(output));
		FOSettings settings = Docx4J.createFOSettings();
		settings.setWmlPackage(wordprocessingMLPackage);
		settings.setApacheFopMime(MimeConstants.MIME_PDF);
		Docx4J.toFO(settings, outputStream, Docx4J.FLAG_EXPORT_PREFER_XSL);
		outputStream.close();
		return true;
	}

	@Override
	public void close() throws IOException {

	}
}

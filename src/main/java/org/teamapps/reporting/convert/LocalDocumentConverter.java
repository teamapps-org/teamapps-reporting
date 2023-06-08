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

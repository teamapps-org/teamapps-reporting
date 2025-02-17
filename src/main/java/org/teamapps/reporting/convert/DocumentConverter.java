/*-
 * ========================LICENSE_START=================================
 * TeamApps Reporting
 * ---
 * Copyright (C) 2020 - 2025 TeamApps.org
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

import org.apache.commons.io.FileUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

public interface DocumentConverter {

	static DocumentConverter createLocalConverter() {
		return new LocalDocumentConverter();
	}

	static DocumentConverter createRemoteConverter(String host) {
		return new RemoteDocumentConverter(host);
	}

	static DocumentConverter createRemoteConverter(String host, String user, String password) {
		return new RemoteDocumentConverter(host, user, password);
	}

	static DocumentConverter createRemoteConverter(String host, String user, String password, String proxyHost, int proxyPort) {
		return new RemoteDocumentConverter(host, user, password, proxyHost, proxyPort);
	}

	default boolean convertDocument(InputStream inputStream, DocumentFormat inputFormat, File output, DocumentFormat outputFormat) throws Exception {
		File tempFile = File.createTempFile("temp", "." + inputFormat.getFormat());
		FileUtils.copyInputStreamToFile(inputStream, tempFile);
		boolean result = convertDocument(tempFile, inputFormat, output, outputFormat);
		tempFile.delete();
		return result;
	}

	boolean convertDocument(File input, DocumentFormat inputFormat, File output, DocumentFormat outputFormat) throws Exception;

	boolean convertDocument(WordprocessingMLPackage wordprocessingMLPackage, File output, DocumentFormat outputFormat) throws Exception;

	void close() throws IOException;
}

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

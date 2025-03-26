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
import org.apache.hc.client5.http.classic.methods.HttpPost;
import org.apache.hc.client5.http.entity.mime.HttpMultipartMode;
import org.apache.hc.client5.http.entity.mime.MultipartEntityBuilder;
import org.apache.hc.client5.http.impl.classic.CloseableHttpClient;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.client5.http.protocol.HttpClientContext;
import org.apache.hc.core5.http.ContentType;
import org.apache.hc.core5.http.HttpEntity;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.*;
import java.util.Base64;

public class RemoteDocumentConverter implements DocumentConverter {
	private final String host;
	private String user;
	private String password;
	private boolean noHttps;

	private CloseableHttpClient client;
	private HttpClientContext context;

	public RemoteDocumentConverter(String host) {
		this(host, null, null);
	}

	public RemoteDocumentConverter(String host, String user, String password) {
		this.host = cleanHost(host);
		this.user = user;
		this.password = password;
		init();
	}

	public RemoteDocumentConverter(String host, String user, String password, String proxyHost, int proxyPort) {
		this.host = host;
		this.user = user;
		this.password = password;
		init();
	}

	private void init() {
		client = HttpClients.custom()
				.addRequestInterceptorLast((request, entity, context) ->
						request.setHeader("Authorization", "Basic " + Base64.getEncoder().encodeToString((user + ":" + password).getBytes())))
				.build();
	}

	public boolean isNoHttps() {
		return noHttps;
	}

	public void setNoHttps(boolean noHttps) {
		this.noHttps = noHttps;
	}

	private String cleanHost(String s) {
		if (s.startsWith("http://")) {
			return s.substring(7);
		} else if (s.startsWith("https://")) {
			return s.substring(8);
		} else {
			return s;
		}
	}

	public boolean convertDocument(File input, DocumentFormat inputFormat, File output, DocumentFormat outputFormat) throws Exception {
		BufferedInputStream inputStream = new BufferedInputStream(new FileInputStream(input));
		return processDocumentConversion(inputStream, inputFormat, output, outputFormat);
	}

	@Override
	public boolean convertDocument(WordprocessingMLPackage wordprocessingMLPackage, File output, DocumentFormat outputFormat) throws Exception {
		File tempFile = File.createTempFile("temp", ".docx");
		wordprocessingMLPackage.save(tempFile);
		BufferedInputStream inputStream = new BufferedInputStream(new FileInputStream(tempFile));
		boolean result = processDocumentConversion(inputStream, DocumentFormat.DOCX, output, outputFormat);
		tempFile.delete();
		return result;
	}

	public boolean processDocumentConversion(InputStream inputStream, DocumentFormat inputFormat, File output, DocumentFormat outputFormat) throws Exception {
		if (outputFormat == DocumentFormat.PNG) {
			UnsupportedFormatException.throwException(outputFormat);
		}
		String uri = (isNoHttps() ? "http://" : "https://") + host + "/conversion?format=" + outputFormat.getFormat();
		HttpPost post = new HttpPost(uri);
		MultipartEntityBuilder builder = MultipartEntityBuilder.create();
		builder.setMode(HttpMultipartMode.EXTENDED);
		builder.addBinaryBody("file", inputStream, ContentType.DEFAULT_BINARY, "input." + inputFormat.getFormat());
		HttpEntity entity = builder.build();
		post.setEntity(entity);
		return client.execute(post, context, response -> {
			int statusCode = response.getCode();
			if (statusCode == 200) {
				InputStream content = response.getEntity().getContent();
				FileUtils.copyInputStreamToFile(content, output);
			}
			return statusCode == 200;
		});
	}

	public void close() throws IOException {
		client.close();
	}

}

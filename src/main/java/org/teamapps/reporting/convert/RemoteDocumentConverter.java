/*-
 * ========================LICENSE_START=================================
 * TeamApps Reporting
 * ---
 * Copyright (C) 2020 - 2021 TeamApps.org
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
import org.apache.http.HttpEntity;
import org.apache.http.HttpHost;
import org.apache.http.HttpResponse;
import org.apache.http.auth.AuthScope;
import org.apache.http.auth.UsernamePasswordCredentials;
import org.apache.http.client.AuthCache;
import org.apache.http.client.CredentialsProvider;
import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.protocol.HttpClientContext;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.impl.auth.BasicScheme;
import org.apache.http.impl.client.BasicAuthCache;
import org.apache.http.impl.client.BasicCredentialsProvider;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.*;

public class RemoteDocumentConverter implements DocumentConverter {
	private final String host;
	private String user;
	private String password;

	private String proxyHost;
	private int proxyPort;

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
		this.proxyHost = proxyHost;
		this.proxyPort = proxyPort;
		init();
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

	private void init() {
		context = HttpClientContext.create();
		if (user != null) {
			HttpHost targetHost = new HttpHost(host, 443, "https");
			CredentialsProvider credentialsProvider = new BasicCredentialsProvider();
			credentialsProvider.setCredentials(AuthScope.ANY, new UsernamePasswordCredentials(user, password));
			AuthCache authCache = new BasicAuthCache();
			authCache.put(targetHost, new BasicScheme());
			context.setCredentialsProvider(credentialsProvider);
			context.setAuthCache(authCache);
		}

		if (proxyHost != null) {
			RequestConfig requestConfig = RequestConfig.custom()
					.setProxy(new HttpHost(proxyHost, proxyPort))
					.build();
			context.setRequestConfig(requestConfig);
		}
		client = HttpClients.custom().build();
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
		HttpPost post = new HttpPost("https://" + host + "/conversion?format=" + outputFormat.getFormat());
		MultipartEntityBuilder builder = MultipartEntityBuilder.create();
		builder.setMode(HttpMultipartMode.BROWSER_COMPATIBLE);
		builder.addBinaryBody("file", inputStream, ContentType.DEFAULT_BINARY, "input." + inputFormat.getFormat());
		HttpEntity entity = builder.build();
		post.setEntity(entity);
		HttpResponse response = client.execute(post, context);
		int statusCode = response.getStatusLine().getStatusCode();
		if (statusCode == 200) {
			InputStream content = response.getEntity().getContent();
			FileUtils.copyInputStreamToFile(content, output);
		}
		return statusCode == 200;
	}

	public void close() throws IOException {
		client.close();
	}

}

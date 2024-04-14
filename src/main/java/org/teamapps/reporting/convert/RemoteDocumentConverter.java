/*-
 * ========================LICENSE_START=================================
 * TeamApps Reporting
 * ---
 * Copyright (C) 2020 - 2024 TeamApps.org
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
import org.apache.hc.client5.http.auth.AuthCache;
import org.apache.hc.client5.http.auth.AuthScope;
import org.apache.hc.client5.http.auth.UsernamePasswordCredentials;
import org.apache.hc.client5.http.classic.methods.HttpGet;
import org.apache.hc.client5.http.classic.methods.HttpPost;
import org.apache.hc.client5.http.config.RequestConfig;
import org.apache.hc.client5.http.entity.mime.HttpMultipartMode;
import org.apache.hc.client5.http.entity.mime.MultipartEntityBuilder;
import org.apache.hc.client5.http.impl.auth.BasicAuthCache;
import org.apache.hc.client5.http.impl.auth.BasicCredentialsProvider;
import org.apache.hc.client5.http.impl.auth.BasicScheme;
import org.apache.hc.client5.http.impl.classic.CloseableHttpClient;
import org.apache.hc.client5.http.impl.classic.CloseableHttpResponse;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.client5.http.protocol.HttpClientContext;
import org.apache.hc.core5.http.ContentType;
import org.apache.hc.core5.http.HttpEntity;
import org.apache.hc.core5.http.HttpHost;
import org.apache.hc.core5.http.ParseException;
import org.apache.hc.core5.http.io.entity.EntityUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.*;
import java.net.URI;
import java.net.URISyntaxException;

public class RemoteDocumentConverter implements DocumentConverter {
	private final String host;
	private String user;
	private String password;
	private boolean noHttps;

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

	private void init() {
		context = HttpClientContext.create();
		if (user != null) {
			HttpHost targetHost = new HttpHost("https", host, 443);
			BasicCredentialsProvider credentialsProvider = new BasicCredentialsProvider();
			credentialsProvider.setCredentials(new AuthScope(host, 443), new UsernamePasswordCredentials(user, password.toCharArray()));
			AuthCache authCache = new BasicAuthCache();
			authCache.put(targetHost, new BasicScheme());
			context.setCredentialsProvider(credentialsProvider);
			context.setAuthCache(authCache);
		}

		if (proxyHost != null && !proxyHost.isBlank()) {
			RequestConfig requestConfig = RequestConfig.custom()
					.setProxy(new HttpHost(proxyHost, proxyPort))
					.build();
			context.setRequestConfig(requestConfig);
		}
		client = HttpClients.custom().build();
	}

	public static String getWithBasicAuth(final String url, final String user, final String pass) throws URISyntaxException, IOException, ParseException {
		String result = null;
		URI uri = new URI(url);
		final BasicCredentialsProvider credentialsProvider = new BasicCredentialsProvider();
		AuthScope authScope = new AuthScope(uri.getHost(), uri.getPort());
		credentialsProvider.setCredentials(authScope, new UsernamePasswordCredentials(user, pass.toCharArray()));
		try (final CloseableHttpClient httpclient = HttpClients.custom().setDefaultCredentialsProvider(credentialsProvider).build()) {
			final HttpGet httpget = new HttpGet(url);
			try (final CloseableHttpResponse response = httpclient.execute(httpget)) {
				result = EntityUtils.toString(response.getEntity());
			}
		} catch (IOException | ParseException e) {
			e.printStackTrace();
		}
		return result;
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
		CloseableHttpResponse response = client.execute(post, context);
		int statusCode = response.getCode();
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

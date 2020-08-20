package org.teamapps.reporting.convert;

public enum DocumentFormat {
	PDF("pdf"),
	ODT("odt"),
	DOCX("docx"),
	RTF("rtf"),
	TXT("txt"),
	PNG("png"),

	;
	private final String format;

	DocumentFormat(String format) {
		this.format = format;
	}

	public String getFormat() {
		return format;
	}
}

package org.teamapps.reporting.convert;

public class UnsupportedFormatException extends Exception{

	public static void throwException(DocumentFormat format) throws UnsupportedFormatException {
		throw new UnsupportedFormatException("Unsupported format:" + format.getFormat());
	}

	public UnsupportedFormatException(String message) {
		super(message);
	}
}

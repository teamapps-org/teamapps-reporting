package org.teamapps.reporting.builder;

import java.util.HashMap;
import java.util.Map;

public class RowBuilder {

	private Map<String, String> rowMap = new HashMap<>();

	public RowBuilder() {
	}

	public RowBuilder setColumnValue(String columnKey, String value) {
		rowMap.put(columnKey, value);
		return this;
	}

	public Map<String, String> getRowMap() {
		return rowMap;
	}
}

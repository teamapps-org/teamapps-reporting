package org.teamapps.reporting.builder;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class TableBuilder {

	private List<String> keys;
	private List<RowBuilder> rowBuilders = new ArrayList<>();

	protected TableBuilder(List<String> keys) {
		this.keys = keys;
	}

	public RowBuilder addRow() {
		RowBuilder rowBuilder = new RowBuilder();
		rowBuilders.add(rowBuilder);
		return rowBuilder;
	}

	public List<String> getKeys() {
		return keys;
	}

	public List<Map<String, String>> createReplacementMap() {
		List<Map<String, String>> rowsReplacementMap = new ArrayList<>();
		for (RowBuilder rowBuilder : rowBuilders) {
			rowsReplacementMap.add(rowBuilder.getRowMap());
		}
		return rowsReplacementMap;
	}


}

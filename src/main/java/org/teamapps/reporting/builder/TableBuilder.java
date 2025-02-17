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
package org.teamapps.reporting.builder;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

public class TableBuilder {

	private List<String> keys;
	private List<RowBuilder> rowBuilders = new ArrayList<>();
	private List<List<String>> removeUnusedTemplateRows = new ArrayList<>();
	private boolean copyTable;

	protected TableBuilder(List<String> keys, boolean copyTable) {
		this.keys = keys;
		this.copyTable = copyTable;
	}

	public RowBuilder addRow() {
		RowBuilder rowBuilder = new RowBuilder();
		rowBuilders.add(rowBuilder);
		return rowBuilder;
	}

	public List<String> getKeys() {
		return keys;
	}

	public boolean isCopyTable() {
		return copyTable;
	}

	public List<Map<String, String>> createReplacementMap() {
		List<Map<String, String>> rowsReplacementMap = new ArrayList<>();
		for (RowBuilder rowBuilder : rowBuilders) {
			rowsReplacementMap.add(rowBuilder.getRowMap());
		}
		return rowsReplacementMap;
	}

	public TableBuilder removeUnusedTemplateRow(String ... keys) {
		removeUnusedTemplateRows.add(Arrays.asList(keys));
		return this;
	}

	public List<List<String>> getRemoveUnusedTemplateRows() {
		return removeUnusedTemplateRows;
	}
}

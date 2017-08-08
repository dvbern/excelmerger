/*
 * Copyright 2017 DV Bern AG
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * limitations under the License.
 */

package ch.dvbern.oss.lib.excelmerger;

import java.util.Map;
import java.util.Optional;
import java.util.regex.Matcher;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import static com.google.common.base.Preconditions.checkNotNull;

public class Context {
	public static final int BASE_10 = 10;
	@Nonnull
	private final Workbook workbook;
	@Nonnull
	private final Sheet sheet;
	@Nonnull
	private final Map<String, MergeField<?>> mergeFields;

	private int currentRow = 0;

	@Nonnull
	private CellCopyPolicy cellCopyPolicy = new CellCopyPolicy.Builder()
		.build();

	Context(@Nonnull Workbook workbook, @Nonnull Sheet sheet, @Nonnull Map<String, MergeField<?>> mergeFields) {
		this(workbook, sheet, mergeFields, sheet.getFirstRowNum());
	}

	Context(
		@Nonnull Workbook workbook,
		@Nonnull Sheet sheet,
		@Nonnull Map<String, MergeField<?>> mergeFields,
		int startRow) {

		this.workbook = checkNotNull(workbook);
		this.sheet = checkNotNull(sheet);
		this.mergeFields = checkNotNull(mergeFields);
		this.currentRow = startRow;
	}

	@Nonnull
	public Workbook getWorkbook() {
		return workbook;
	}

	@Nonnull
	public Sheet getSheet() {
		return sheet;
	}

	public int currentRowNum() {
		return currentRow;
	}

	@Nonnull
	public Row currentRow() {
		Row row = sheet.getRow(currentRowNum());
		if (row == null) {
			row = sheet.createRow(currentRowNum());
		}
		return row;
	}

	public void advanceRow() {
		currentRow++;
	}

	@Nonnull
	Optional<GroupPlaceholder> detectGroup() {
		Row row = currentRow();

		// von hinten nach vorne durcharbeiten
		for (int i = Math.max(row.getLastCellNum(), 0); i >= Math.max(row.getFirstCellNum(), 0); i--) {
			Cell cell = row.getCell(i);

			Optional<GroupPlaceholder> groupPlaceholder = parsePlaceholder(cell)
				.filter(placeholder -> placeholder instanceof GroupPlaceholder)
				.map(placeholder -> (GroupPlaceholder) placeholder);

			if (groupPlaceholder.isPresent()) {
				return groupPlaceholder;
			}
		}

		return Optional.empty();
	}

	@Nonnull
	Optional<Placeholder> parsePlaceholder(@Nullable Cell cell) {
		if (cell == null || cell.getCellTypeEnum() != CellType.STRING) {
			return Optional.empty();
		}

		Matcher matcher = ExcelMerger.MERGEFIELD_REX.matcher(cell.getStringCellValue());
		if (!matcher.matches()) {
			return Optional.empty();
		}

		String pattern = matcher.group(ExcelMerger.REX_GROUP_PATTERN);
		String key = matcher.group(ExcelMerger.REX_GROUP_KEY);

		Integer groupRows = matcher.group(ExcelMerger.REF_GROUP_ROWS) != null ?
			Integer.valueOf(matcher.group(ExcelMerger.REF_GROUP_ROWS), BASE_10) :
			null;

		MergeField<?> field = mergeFields.get(key);

		if (field == null) {
			return Optional.empty();
		}

		if (field.getType() == MergeField.Type.REPEAT_ROW) {
			GroupPlaceholder groupPlaceholder = new GroupPlaceholder(cell, pattern, key, field, groupRows);

			return Optional.of(groupPlaceholder);
		}

		Placeholder placeholder = new Placeholder(cell, pattern, key, field);

		return Optional.of(placeholder);
	}

	@Nonnull
	public CellCopyPolicy getCellCopyPolicy() {
		return cellCopyPolicy;
	}

	public void setCellCopyPolicy(@Nonnull CellCopyPolicy cellCopyPolicy) {
		this.cellCopyPolicy = cellCopyPolicy;
	}
}

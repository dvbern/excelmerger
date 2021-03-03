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

import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.RepeatRowMergeField;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class RowFiller {

	private static final int RANDOM_ACCESS_WINDOW_SIZE = 10;

	@Nonnull
	private final SXSSFSheet sheet;
	@Nonnull
	private final Context ctx;
	@Nonnull
	private final Row sourceRow;
	private int numberOfTargetRows = 0;
	private int numberOfMergedRows = 0;
	@Nullable
	private ExcelMergerDTO firstRowData = null;

	public RowFiller(
		@Nonnull SXSSFSheet sheet,
		@Nonnull Context ctx,
		@Nonnull Row sourceRow,
		int numberOfTargetRows) {

		this.sheet = sheet;
		this.ctx = ctx;
		this.sourceRow = sourceRow;
		this.numberOfTargetRows = numberOfTargetRows;
	}

	/**
	 * POI has a very high memory usage when all rows must be kept in-memory (normal ExcelMerger behaviour).
	 * For XSSF Sheets a streaming low-memory approach can be used. Unfortunately it comes with some limitations:
	 * <ul>
	 * <li>Rows cannot be copied</li>
	 * <li>Rows cannot be shifted</li>
	 * <li>Formulas are only supported when they refer to colls on the same row
	 * (no evaluation over the whole sheet)</li>
	 * </ul>
	 *
	 * <p>All data rows that are written with the RowFiller will overwrite any existing rows.
	 * Hence, the RowFiller should only be used when below it's source row no other merge fields are defined.</p>
	 *
	 * <p>All limitations of SXSSF apply. It is your responsibility to cleanup the temporary files by calling
	 * {@code workbook.dispose();} after writing the workbook.</p>
	 *
	 * @param sheet the sheet you want to write to
	 * @param fields should contain a {@link RepeatRowMergeField}, which is used to determine the source row,
	 * and all merge fields that are needed to fill that source row with data.
	 * @param numberOfDataRows the number of data rows that will be filled. E.g. the number of
	 * {@link RowFiller#fillRow(ExcelMergerDTO)} executions
	 * @return a RowFiller, which can be used to write a single {@link ExcelMergerDTO} data row.
	 * @see <a href="https://poi.apache.org/spreadsheet/how-to.html#sxssf">SXSSF HowTo</a>
	 */
	@SuppressWarnings("PMD.CloseResource")
	@Nonnull
	public static RowFiller initRowFiller(
		@Nonnull XSSFSheet sheet,
		@Nonnull List<MergeField<?>> fields,
		int numberOfDataRows) {

		Map<String, MergeField<?>> fieldMap = fields.stream()
			.collect(Collectors.toMap(MergeField::getKey, field -> field));

		Context ctx = new Context(sheet.getWorkbook(), sheet, fieldMap);

		GroupPlaceholder groupPlaceholder = IntStream.rangeClosed(0, sheet.getLastRowNum())
			.mapToObj(sheet::getRow)
			.map(ctx::detectGroup)
			.filter(Optional::isPresent)
			.map(Optional::get)
			.findFirst()
			.orElseThrow(() -> new ExcelMergeRuntimeException("No RepeatRowMergeField marker found"));

		if (groupPlaceholder.getRows() != 1) {
			throw new IllegalStateException("Currently, only 1 source row is supported, but "
				+ groupPlaceholder
				+ " given");
		}

		Row sourceRow = groupPlaceholder.getCell().getRow();
		groupPlaceholder.clearPlaceholder();

		SXSSFWorkbook wb = new SXSSFWorkbook(sheet.getWorkbook());
		wb.setCompressTempFiles(true);
		// wenn true, dann evaluiert Excel die Formeln. LibreOffice kann das leider nicht
		wb.setForceFormulaRecalculation(true);

		SXSSFSheet sh = wb.getSheetAt(sheet.getWorkbook().getSheetIndex(sheet));
		// keep 10 rows in memory, exceeding rows will be flushed to disk
		sh.setRandomAccessWindowSize(RANDOM_ACCESS_WINDOW_SIZE);

		return new RowFiller(sh, ctx, sourceRow, numberOfDataRows);
	}

	@Nonnull
	public SXSSFSheet getSheet() {
		return sheet;
	}

	public void fillRow(@Nonnull ExcelMergerDTO rowData) {
		if (firstRowData == null) {
			// we cannot write to the sourceRow yet, because we need it as a template for the remaining rows
			// -> store it for later writing
			firstRowData = rowData;
		} else {
			SXSSFRow targetRow = sheet.createRow(sourceRow.getRowNum() + numberOfMergedRows + 1);
			// copy styles & formulas
			ExcelMerger.copyCells(sourceRow, targetRow);

			mergeRow(rowData, targetRow);
		}

		if (isLastRow()) {
			// since we are creating new rows, we have to write the data we held back to our template
			ExcelMerger.mergeRow(ctx, firstRowData, sourceRow);
		}
	}

	private int getNumberOfRemainingRows() {
		return numberOfTargetRows - numberOfMergedRows;
	}

	private boolean isLastRow() {
		return getNumberOfRemainingRows() == 1;
	}

	private void mergeRow(@Nonnull ExcelMergerDTO rowData, @Nonnull Row targetRow) {
		ExcelMerger.mergeRow(ctx, rowData, targetRow);
		numberOfMergedRows++;
	}
}

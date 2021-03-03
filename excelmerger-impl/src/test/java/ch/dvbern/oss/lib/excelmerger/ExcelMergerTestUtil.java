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

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.function.Function;

import javax.annotation.Nonnull;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import static java.util.Objects.requireNonNull;

public final class ExcelMergerTestUtil {

	private ExcelMergerTestUtil() {
		// Util
	}

	static final String BASE = "ch/dvbern/oss/lib/excelmerger/";
	static final String BELEGUNGSPLAN = BASE + "belegungsplan.xlsx";
	static final String WARTELISTE = BASE + "warteliste.xlsx";

	@Nonnull
	static final Function<String, Workbook> GET_WORKBOOK = (name) -> {
		InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream(name);

		return ExcelMerger.createWorkbookFromTemplate(requireNonNull(is));
	};

	@Nonnull
	static String getVal(@Nonnull Sheet sheet, int rowName, @Nonnull String colName) {
		int rowNum = rowName - 1; // poi is zero-based
		int colNum = CellReference.convertColStringToIndex(colName);
		Cell cell = sheet.getRow(rowNum).getCell(colNum);
		return cell == null ? "" : cell.getStringCellValue();
	}

	static double getNumVal(@Nonnull Sheet sheet, int rowName, @Nonnull String colName) {
		int rowNum = rowName - 1; // poi is zero-based
		int colNum = CellReference.convertColStringToIndex(colName);
		Cell cell = sheet.getRow(rowNum).getCell(colNum);
		return cell.getNumericCellValue();
	}

	static int named(@Nonnull String columnName) {
		return CellReference.convertColStringToIndex(columnName);
	}

	@Nonnull
	public static Cell createCell(@Nonnull Workbook wb, @Nonnull String pattern) {
		Sheet sheet = wb.createSheet("new sheet");

		// Create a row and put some cells in it. Rows are 0 based.
		Row row = sheet.createRow(0);
		// Create a cell
		Cell cell = row.createCell(0);
		cell.setCellValue(pattern);

		return cell;
	}

	public static void setDataFormat(@Nonnull Workbook wb, @Nonnull Cell cell, @Nonnull String format) {
		CellStyle cellStyle = wb.createCellStyle();
		CreationHelper createHelper = wb.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(format));
		cell.setCellStyle(cellStyle);
	}

	@Nonnull
	public static String writeWorkbookToFile(@Nonnull Workbook wb, @Nonnull String sheetName) throws IOException {
		String name = "target/" + sheetName;
		try (OutputStream out = new FileOutputStream(name)) {
			wb.write(out);
		}

		return name;
	}
}

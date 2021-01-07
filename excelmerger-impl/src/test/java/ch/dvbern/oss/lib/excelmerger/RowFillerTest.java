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

import java.io.FileInputStream;
import java.util.Arrays;
import java.util.Collections;
import java.util.stream.IntStream;

import javax.annotation.Nonnull;

import ch.dvbern.oss.lib.excelmerger.converters.StandardConverters;
import ch.dvbern.oss.lib.excelmerger.mergefields.RepeatRowMergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.SimpleMergeField;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.writeWorkbookToFile;
import static org.junit.jupiter.api.Assertions.assertEquals;

public class RowFillerTest {

	private static final SimpleMergeField<Integer> VALUE_1 =
		new SimpleMergeField<>("value1", StandardConverters.INTEGER_CONVERTER);
	private static final SimpleMergeField<Integer> VALUE_2 =
		new SimpleMergeField<>("value2", StandardConverters.INTEGER_CONVERTER);
	private static final RepeatRowMergeField REPEAT_ROW = new RepeatRowMergeField("row");

	@Test
	public void testSingleRow() {
		XSSFSheet xssfSheet = init();
		int initialNumberOfRows = xssfSheet.getPhysicalNumberOfRows();

		int numberOfDataRows = 1;

		RowFiller rowFiller = executeTestRun(xssfSheet, numberOfDataRows);

		assertEquals(initialNumberOfRows, xssfSheet.getPhysicalNumberOfRows());
		assertEquals(0, rowFiller.getSheet().getPhysicalNumberOfRows());

		// dispose of temporary files backing this workbook on disk
		rowFiller.getSheet().getWorkbook().dispose();
	}

	@Test
	public void testManyRows() throws Exception {
		XSSFSheet xssfSheet = init();
		int initialNumberOfRows = xssfSheet.getPhysicalNumberOfRows();

		int numberOfDataRows = SpreadsheetVersion.EXCEL2007.getMaxRows() - initialNumberOfRows + 1;

		RowFiller rowFiller = executeTestRun(xssfSheet, numberOfDataRows);

		writeWorkbookToFile(rowFiller.getSheet().getWorkbook(), "sxssf-filled.xlsx");

		assertEquals(initialNumberOfRows, xssfSheet.getPhysicalNumberOfRows());
		assertEquals(numberOfDataRows - 1, rowFiller.getSheet().getPhysicalNumberOfRows());

		// dispose of temporary files backing this workbook on disk
		rowFiller.getSheet().getWorkbook().dispose();
	}

	@Test
	public void testCombinationOfExcelMergerAndRowFiller() throws Exception {
		XSSFSheet xssfSheet = init();
		int initialNumberOfRows = xssfSheet.getPhysicalNumberOfRows();

		int numberOfDataRows = 3;

		SimpleMergeField<String> name = new SimpleMergeField<>("name", StandardConverters.STRING_CONVERTER);
		ExcelMergerDTO base = new ExcelMergerDTO();
		String expectedName = "This is my name";
		base.addValue(name, expectedName);
		ExcelMerger.mergeData(xssfSheet, Collections.singletonList(name), base);

		RowFiller rowFiller = executeTestRun(xssfSheet, numberOfDataRows);

		String file = writeWorkbookToFile(rowFiller.getSheet().getWorkbook(), "sxssf-combined.xlsx");

		assertEquals(initialNumberOfRows, xssfSheet.getPhysicalNumberOfRows());
		assertEquals(2, rowFiller.getSheet().getPhysicalNumberOfRows());
		assertEquals(expectedName, xssfSheet.getRow(0).getCell(1).getStringCellValue());

		// dispose of temporary files backing this workbook on disk
		rowFiller.getSheet().getWorkbook().dispose();

		Workbook workbookFromTemplate = ExcelMerger.createWorkbookFromTemplate(new FileInputStream(file));
		Sheet filledSheet = workbookFromTemplate.getSheetAt(0);

		// Verify formula shifting
		IntStream.rangeClosed(3, 5)
			.forEach(i -> {
				Row row = filledSheet.getRow(i - 1);

				assertEquals(String.format("A%1$d+B%1$d", i), row.getCell(2).getCellFormula());
				assertEquals(String.format("3*B%1$d", i), row.getCell(5).getCellFormula());
			});
	}

	@Nonnull
	private XSSFSheet init() {
		String filename = "sxssf.xlsx";
		XSSFWorkbook wb_template =
			(XSSFWorkbook) ExcelMergerTestUtil.GET_WORKBOOK.apply(ExcelMergerTestUtil.BASE + filename);

		return wb_template.getSheetAt(0);
	}

	@Nonnull
	private RowFiller executeTestRun(@Nonnull XSSFSheet sheet, int numberOfDataRows) {
		RowFiller rowFiller = RowFiller.initRowFiller(
			sheet,
			Arrays.asList(VALUE_1, VALUE_2, REPEAT_ROW),
			numberOfDataRows
		);

		IntStream.range(0, numberOfDataRows).forEach(i -> {
			ExcelMergerDTO rowDTO = new ExcelMergerDTO();
			rowDTO.addValue(VALUE_1, 1);
			rowDTO.addValue(VALUE_2, 2);

			rowFiller.fillRow(rowDTO);
		});

		return rowFiller;
	}

}

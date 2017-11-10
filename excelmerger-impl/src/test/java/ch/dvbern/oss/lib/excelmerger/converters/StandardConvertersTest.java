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

package ch.dvbern.oss.lib.excelmerger.converters;

import java.math.BigDecimal;

import ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import static org.junit.Assert.assertEquals;

public class StandardConvertersTest {

	@Test
	public void testBigDecimalConversion() throws Exception {
		Workbook wb = new XSSFWorkbook();

		String pattern = "{percentTest}";
		Cell cell = ExcelMergerTestUtil.createCell(wb, pattern);

		BigDecimal bigDecimal = BigDecimal.valueOf(70)
			.setScale(0, BigDecimal.ROUND_HALF_UP);

		StandardConverters.PERCENT_CONVERTER.setCellValue(cell, pattern, bigDecimal);

		BigDecimal expectedValue = BigDecimal.valueOf(0.7);
		BigDecimal actualValue = BigDecimal.valueOf(cell.getNumericCellValue());
		assertEquals(0, expectedValue.compareTo(actualValue));

		ExcelMergerTestUtil.writeWorkbookToFile(wb, "percentageTest.xlsx");
	}

	@Test
	public void testBigDecimalConversionWithText() throws Exception {
		Workbook wb = new XSSFWorkbook();

		String pattern = "{percentTest}";
		Cell cell = ExcelMergerTestUtil.createCell(wb, "My Percentage " + pattern);

		BigDecimal bigDecimal = BigDecimal.valueOf(70)
			.setScale(0, BigDecimal.ROUND_HALF_UP);

		StandardConverters.PERCENT_CONVERTER.setCellValue(cell, pattern, bigDecimal);

		assertEquals("My Percentage 70%", cell.getStringCellValue());

		ExcelMergerTestUtil.writeWorkbookToFile(wb, "percentageWithTextTest.xlsx");
	}
}

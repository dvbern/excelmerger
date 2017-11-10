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

package ch.dvbern.oss.lib.excelmerger.mergefields;

import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;

import ch.dvbern.oss.lib.excelmerger.ExcelMerger;
import ch.dvbern.oss.lib.excelmerger.ExcelMergerDTO;
import ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import static ch.dvbern.oss.lib.excelmerger.converters.StandardConverters.DATE_CONVERTER;
import static org.junit.Assert.assertEquals;

public class MergeFieldTest {

	@Test
	public void testLocalDateMergeField() throws Exception {
		Workbook wb = new XSSFWorkbook();
		Cell cell = ExcelMergerTestUtil.createCell(wb, "{date}");

		ExcelMergerTestUtil.setDataFormat(wb, cell, "dd.mm.yyyy");
		SimpleMergeField<LocalDate> localDateMergeField = new SimpleMergeField<>("date", DATE_CONVERTER);

		ExcelMergerDTO excelMergerDTO = new ExcelMergerDTO();
		LocalDate localDate = LocalDate.of(2017, 9, 30);
		excelMergerDTO.addValue(localDateMergeField, localDate);

		ExcelMerger.mergeData(wb.getSheetAt(0), new MergeField[] { localDateMergeField }, excelMergerDTO);

		Date expected = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
		assertEquals(expected, cell.getDateCellValue());

		ExcelMergerTestUtil.writeWorkbookToFile(wb, "localDateTest.xlsx");
	}

}

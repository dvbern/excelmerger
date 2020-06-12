/*
 * Copyright 2020 DV Bern AG
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

import java.io.IOException;
import java.util.Arrays;

import ch.dvbern.oss.lib.excelmerger.converters.Converter;
import ch.dvbern.oss.lib.excelmerger.converters.StandardConverters;
import ch.dvbern.oss.lib.excelmerger.mergefields.RepeatRowMergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.SimpleMergeField;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.writeWorkbookToFile;

public class RepeatRowAuthHeightTest {

	@SuppressWarnings("JUnitTestMethodWithNoAssertions")
	@Test
	public void repeatRowAutoHeightTest() throws ExcelMergeException, IOException {
		String filename = "repeatRowAutoHeight.xlsx";
		Workbook wb = ExcelMergerTestUtil.GET_WORKBOOK.apply(ExcelMergerTestUtil.BASE + filename);

		RepeatRowMergeField repeatRowField = new RepeatRowMergeField("repeatRow");
		Converter<String> autoHeightConverter =
			StandardConverters.autoHeightConverter(StandardConverters.STRING_CONVERTER);
		SimpleMergeField<String> value = new SimpleMergeField<>("someValue", autoHeightConverter);

		ExcelMergerDTO dto = new ExcelMergerDTO();
		StringBuilder repeatedText = new StringBuilder("I will do my homework.");

		for (int i = 0; i < 10; i++) {
			ExcelMergerDTO row = dto.createGroup(repeatRowField);
			row.addValue(value, repeatedText.toString());

			repeatedText.append(' ').append(repeatedText.toString());
		}

		Sheet sheet = wb.getSheet("Sheet1");
		ExcelMerger.mergeData(sheet, Arrays.asList(repeatRowField, value), dto);

		writeWorkbookToFile(wb, "repeatRowAutoHeight-filled.xlsx");
	}
}

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

import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import ch.dvbern.oss.lib.excelmerger.converters.StandardConverters;
import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.RepeatRowMergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.SimpleMergeField;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class CopyRowTest {

	@Test
	public void testMergedRegionsAreCopied() throws Exception {
		String filename = "copyRowWithMergedRegion.xlsx";
		Workbook wb = ExcelMergerTestUtil.GET_WORKBOOK.apply(ExcelMergerTestUtil.BASE + filename);

		RepeatRowMergeField repeatRowField = new RepeatRowMergeField("repeatRow");
		SimpleMergeField<String> mergedColumns =
			new SimpleMergeField<>("mergedColumns", StandardConverters.STRING_CONVERTER);

		RepeatRowMergeField repeatRow2Field = new RepeatRowMergeField("repeatRow2");
		SimpleMergeField<String> mergedColumns2 =
			new SimpleMergeField<>("mergedColumns2", StandardConverters.STRING_CONVERTER);

		ExcelMergerDTO excelMergerDTO = new ExcelMergerDTO();

		int numberOfMergedColumns = 10;
		for (int i = 0; i < numberOfMergedColumns; i++) {
			ExcelMergerDTO repeatRow = excelMergerDTO.createGroup(repeatRowField);
			repeatRow.addValue(mergedColumns, "foo " + (i + 1));
		}

		int numberOfMergedColumns2 = 7;
		for (int i = 0; i < numberOfMergedColumns2; i++) {
			ExcelMergerDTO group = excelMergerDTO.createGroup(repeatRow2Field);
			group.addValue(mergedColumns2, "bar " + (i + 1));
		}

		Sheet sheet = wb.getSheetAt(0);

		List<MergeField<?>> fields =
			Arrays.asList(repeatRowField, mergedColumns, repeatRow2Field, mergedColumns2);

		ExcelMerger.mergeData(sheet, fields, excelMergerDTO);

		ExcelMergerTestUtil.writeWorkbookToFile(wb, filename);

		assertEquals(numberOfMergedColumns + numberOfMergedColumns2, sheet.getNumMergedRegions());

		for (int i = 0; i < numberOfMergedColumns; i++) {
			CellRangeAddress cellRangeAddress = new CellRangeAddress(i, i, 0, 2);
			assertTrue(sheet.getMergedRegions().contains(cellRangeAddress));
		}

		for (int i = 0; i < numberOfMergedColumns2; i++) {
			int row = numberOfMergedColumns + 1 + i;
			CellRangeAddress cellRangeAddress = new CellRangeAddress(row, row, 0, 1);
			assertTrue(sheet.getMergedRegions().contains(cellRangeAddress));
		}
	}

	@Test
	public void testMergedRegionsAreCopiedWithGroups() throws Exception {
		String filename = "copyGroupWithMergedRegion.xlsx";
		Workbook wb = ExcelMergerTestUtil.GET_WORKBOOK.apply(ExcelMergerTestUtil.BASE + filename);

		RepeatRowMergeField repeatRowField = new RepeatRowMergeField("repeatGroup");
		SimpleMergeField<String> mergedColumn =
			new SimpleMergeField<>("mergedColumn", StandardConverters.STRING_CONVERTER);

		SimpleMergeField<String> mergedRowAndColumn =
			new SimpleMergeField<>("mergedRowAndColumn", StandardConverters.STRING_CONVERTER);

		ExcelMergerDTO excelMergerDTO = new ExcelMergerDTO();

		int numberOfCopies = 10;
		for (int i = 0; i < numberOfCopies; i++) {
			ExcelMergerDTO repeatRow = excelMergerDTO.createGroup(repeatRowField);
			repeatRow.addValue(mergedColumn, "foo " + (i + 1));
			repeatRow.addValue(mergedRowAndColumn, "bar " + (i + 1));
		}

		Sheet sheet = wb.getSheetAt(0);
		ExcelMerger.mergeData(sheet, Arrays.asList(repeatRowField, mergedColumn, mergedRowAndColumn), excelMergerDTO);

		ExcelMergerTestUtil.writeWorkbookToFile(wb, filename);

		assertEquals(numberOfCopies * 2L, sheet.getNumMergedRegions());
		for (int i = 0; i < numberOfCopies; i++) {
			CellRangeAddress cellRangeAddress = new CellRangeAddress(i, i, 0, 2);
			sheet.getMergedRegions().contains(cellRangeAddress);
		}
	}

	@Test
	public void testNamedRegion() throws Exception {
		String filename = "copyRowWithNamedRegion.xlsx";
		Workbook wb = ExcelMergerTestUtil.GET_WORKBOOK.apply(ExcelMergerTestUtil.BASE + filename);

		RepeatRowMergeField repeatRowField = new RepeatRowMergeField("repeatRow");
		SimpleMergeField<Long> mergedColumns =
			new SimpleMergeField<>("value", StandardConverters.LONG_CONVERTER);

		RepeatRowMergeField repeatRow2Field = new RepeatRowMergeField("repeatRow2");
		SimpleMergeField<Long> mergedColumns2 =
			new SimpleMergeField<>("value2", StandardConverters.LONG_CONVERTER);

		ExcelMergerDTO excelMergerDTO = new ExcelMergerDTO();

		int numberOfRows1 = 10;
		for (int i = 0; i < numberOfRows1; i++) {
			ExcelMergerDTO repeatRow = excelMergerDTO.createGroup(repeatRowField);
			repeatRow.addValue(mergedColumns, Long.valueOf(i));
		}

		int numberOfRows2 = 7;
		for (int i = 0; i < numberOfRows2; i++) {
			ExcelMergerDTO group = excelMergerDTO.createGroup(repeatRow2Field);
			group.addValue(mergedColumns2, Long.valueOf(i));
		}

		Sheet sheet = wb.getSheetAt(0);

		List<MergeField<?>> fields =
			Arrays.asList(repeatRowField, mergedColumns, repeatRow2Field, mergedColumns2);

		ExcelMerger.mergeData(sheet, fields, excelMergerDTO);

		ExcelMergerTestUtil.writeWorkbookToFile(wb, filename);

		int totalRow = 3 + numberOfRows1 + numberOfRows2;
		int total1 = Double.valueOf(sheet.getRow(totalRow).getCell(1).getNumericCellValue()).intValue();
		int total2 = Double.valueOf(sheet.getRow(totalRow).getCell(3).getNumericCellValue()).intValue();
		assertEquals(45, total1);
		assertEquals(21, total2);
	}

	@Test
	public void testShiftDataValidations() throws Exception {
		String filename = "shiftDataValidations.xlsx";
		Workbook wb = ExcelMergerTestUtil.GET_WORKBOOK.apply(ExcelMergerTestUtil.BASE + filename);
		RepeatRowMergeField repeatRowField = new RepeatRowMergeField("repeatRow");

		ExcelMergerDTO excelMergerDTO = new ExcelMergerDTO();
		excelMergerDTO.createGroup(repeatRowField);
		excelMergerDTO.createGroup(repeatRowField);
		excelMergerDTO.createGroup(repeatRowField);

		Sheet sheet = wb.getSheetAt(0);

		assertEquals(1, sheet.getDataValidations().size());
		CellRangeAddress addressBefore = sheet.getDataValidations().get(0).getRegions().getCellRangeAddress(0);

		ExcelMerger.mergeData(sheet, Collections.singletonList(repeatRowField), excelMergerDTO);

		ExcelMergerTestUtil.writeWorkbookToFile(wb, filename);

		assertEquals(1, sheet.getDataValidations().size());
		CellRangeAddress addressAfter = sheet.getDataValidations().get(0).getRegions().getCellRangeAddress(0);

		// two rows are added on top of the data validations
		assertEquals(addressBefore.getFirstRow() + 2, addressAfter.getFirstRow());
	}
}

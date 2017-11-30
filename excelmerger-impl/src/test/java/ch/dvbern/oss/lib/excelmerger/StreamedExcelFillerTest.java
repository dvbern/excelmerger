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
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import javax.annotation.Nonnull;

import ch.dvbern.oss.lib.excelmerger.converters.StandardConverters;
import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.RepeatRowMergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.SimpleMergeField;
import com.google.common.base.Preconditions;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Ignore;
import org.junit.Test;

import static ch.dvbern.oss.lib.excelmerger.ExcelMerger.SAME_ROW_CELL_REF;
import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.writeWorkbookToFile;
import static com.google.common.base.Preconditions.checkNotNull;
import static org.junit.Assert.assertEquals;

public class StreamedExcelFillerTest {

	private static final SimpleMergeField<Integer> VALUE_1 =
		new SimpleMergeField<>("value1", StandardConverters.INTEGER_CONVERTER);
	private static final SimpleMergeField<Integer> VALUE_2 =
		new SimpleMergeField<>("value2", StandardConverters.INTEGER_CONVERTER);
	private static final RepeatRowMergeField REPEAT_ROW = new RepeatRowMergeField("row");

	@Ignore
	@Test
	public void test() throws Exception {
		String filename = "sxssf.xlsx";
		XSSFWorkbook wb_template =
			(XSSFWorkbook) ExcelMergerTestUtil.GET_WORKBOOK.apply(ExcelMergerTestUtil.BASE + filename);
		XSSFSheet xssfSheet = wb_template.getSheetAt(0);
		Row srcRow = xssfSheet.getRow(0);

		SXSSFWorkbook wb = new SXSSFWorkbook(wb_template);
		wb.setCompressTempFiles(true);

		SXSSFSheet sh = wb.getSheetAt(0);

		sh.setRandomAccessWindowSize(10);// keep 100 rows in memory, exceeding rows will be flushed to disk
		for (int rownum = 1; rownum < 100; rownum++) {
			Row row = sh.createRow(rownum);
			ExcelMerger.copyCells(srcRow, row);
		}

		// wenn true, dann evaluiert Excel die Formeln. LibreOffice kann das leider nicht
		wb.setForceFormulaRecalculation(true);
		writeWorkbookToFile(wb, "sxssf-filled.xlsx");
	}

	@Test
	public void testWithDTO() throws Exception {
		ExcelMergerDTO testDTO = createTestDTO();

		String filename = "sxssf.xlsx";
		XSSFWorkbook wb_template =
			(XSSFWorkbook) ExcelMergerTestUtil.GET_WORKBOOK.apply(ExcelMergerTestUtil.BASE + filename);
		XSSFSheet xssfSheet = wb_template.getSheetAt(0);

		SXSSFSheet sxssfSheet = fillRows(xssfSheet, Arrays.asList(VALUE_1, VALUE_2, REPEAT_ROW), testDTO);

		writeWorkbookToFile(sxssfSheet.getWorkbook(), "sxssf-filled.xlsx");

	}

	@Nonnull
	public SXSSFSheet fillRows(
		@Nonnull XSSFSheet sheet,
		@Nonnull List<MergeField<?>> fields,
		@Nonnull ExcelMergerDTO excelMergerDTO) throws ExcelMergeException {

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

		Preconditions.checkState(
			groupPlaceholder.getRows() == 1,
			"Currently, only 1 source row is supported, but {} given",
			groupPlaceholder);

		Row sourceRow = groupPlaceholder.getCell().getRow();
		groupPlaceholder.clearPlaceholder();

		SXSSFWorkbook wb = new SXSSFWorkbook(sheet.getWorkbook());
		wb.setCompressTempFiles(true);
		// wenn true, dann evaluiert Excel die Formeln. LibreOffice kann das leider nicht
		wb.setForceFormulaRecalculation(true);

		SXSSFSheet sh = wb.getSheetAt(sheet.getWorkbook().getSheetIndex(sheet));
		// keep 10 rows in memory, exceeding rows will be flushed to disk
		sh.setRandomAccessWindowSize(10);

		List<ExcelMergerDTO> nestedRows = excelMergerDTO.getGroup(groupPlaceholder.getField());
		if (nestedRows == null || nestedRows.isEmpty()) {
			// no rows to fill -> done
			return sh;
		}

		// since
		IntStream.range(1, nestedRows.size())
			.forEach(i -> {
				SXSSFRow targetRow = sh.createRow(sourceRow.getRowNum() + i);
				// copy styles & formulas
				ExcelMerger.copyCells(sourceRow, targetRow);
				// fill mergeFields
				ExcelMerger.mergeRow(ctx, nestedRows.get(i), targetRow);
			});

		// since we are creating new rows, we have to remove our template
		ExcelMerger.mergeRow(ctx, nestedRows.get(0), sourceRow);

		return sh;
	}

	@Nonnull
	private ExcelMergerDTO createTestDTO() {

		ExcelMergerDTO excelMergerDTO = new ExcelMergerDTO();
		int numberOfRows = 10_000_000;

		excelMergerDTO.createGroup(REPEAT_ROW, numberOfRows);
		checkNotNull(excelMergerDTO.getGroup(REPEAT_ROW))
			.forEach(rowDTO -> {
				rowDTO.addValue(VALUE_1, 1);
				rowDTO.addValue(VALUE_2, 2);
			});

		return excelMergerDTO;
	}

	@Test
	public void testFormulaReplacement() {
		String subst = "$1100";

		assertEquals("SUM(A100:B100)", SAME_ROW_CELL_REF.matcher("SUM(A1:B1)").replaceAll(subst));
		assertEquals("A100+A100*3", SAME_ROW_CELL_REF.matcher("A1+A1*3").replaceAll(subst));
		assertEquals("$A$1", SAME_ROW_CELL_REF.matcher("$A$1").replaceAll(subst));
		assertEquals("$A100", SAME_ROW_CELL_REF.matcher("$A1").replaceAll(subst));
		assertEquals("ABD100+A100", SAME_ROW_CELL_REF.matcher("ABD1+A1").replaceAll(subst));
		assertEquals("A1B1", SAME_ROW_CELL_REF.matcher("A1B1").replaceAll(subst));
		assertEquals("CONCAT(A100; \"A1\")", SAME_ROW_CELL_REF.matcher("CONCAT(A1; \"A1\")").replaceAll(subst));
	}
}

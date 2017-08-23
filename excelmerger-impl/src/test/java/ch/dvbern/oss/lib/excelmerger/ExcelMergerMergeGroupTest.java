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

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

import javax.annotation.Nonnull;

import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.MergeFieldProvider;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.BELEGUNGSPLAN;
import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.GET_WORKBOOK;
import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.WARTELISTE;
import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.getVal;
import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.writeWorkbookToFile;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

@SuppressWarnings("JUnitTestMethodWithNoAssertions")
public class ExcelMergerMergeGroupTest {

	@Test
	public void testMergeSingleGroup_shouldHandleZeroRows() throws Exception {
		final int numRowsTest = 0;
		Tester newMerger = new Tester(new SimpleGroup(numRowsTest, ExcelMerger::mergeSubGroup));
		newMerger.test("improved");
	}

	@Test
	public void testMergeSingleGroup_shouldHandleSingleRow() throws Exception {
		final int numRowsTest = 1;
		Tester newMerger = new Tester(new SimpleGroup(numRowsTest, ExcelMerger::mergeSubGroup));
		newMerger.test("improved");
	}

	@Test
	public void testMergeGroupWithMultipleRows_shouldHandleZeroGroups() throws Exception {
		Tester newMerger = new Tester(new GroupWithMultipleRows(0, ExcelMerger::mergeSubGroup));
		newMerger.test("improved");
	}

	@Test
	public void testMergeGroupWithMultipleRows_shouldHandleSingleGroup() throws Exception {
		Tester newMerger = new Tester(new GroupWithMultipleRows(1, ExcelMerger::mergeSubGroup));
		newMerger.test("improved");
	}

	@Test
	public void testMergeGroupWithMultipleRows_shouldHandleMultipleGroups() throws Exception {
		Tester newMerger = new Tester(new GroupWithMultipleRows(3, ExcelMerger::mergeSubGroup));
		newMerger.test("improved");
	}

	@Test
	public void testNestedGroups() throws Exception {
		Tester newMerger = new Tester(new NestedGroups(1, 2, ExcelMerger::mergeSubGroup));
		newMerger.test("improved");
	}

	private static class SimpleGroup extends ContextBuilder {

		public SimpleGroup(int numRowsTest, @Nonnull ExcelMerger.GroupMerger merger) {
			super(GET_WORKBOOK.apply(WARTELISTE), merger);
			setNumRowsBefore(5);
			setNumRowsAfter(2);
			setNumRowsTest(numRowsTest);

			for (int i = 0; i < numRowsTest; i++) {
				addKind(getExcelMergerDTO());
			}

			MergeFieldProvider.toMergeFields(MergeFieldWarteliste.values())
				.forEach(field -> getFieldMap().put(field.getKey(), field));
		}

		@Override
		protected void validate() {
			int numRowsTotal = getNumRowsBefore() + getNumRowsTest() + getNumRowsAfter();
			if (getNumRowsTest() == 0) {
				// Weil die bisherige Implementation eine leere Zeile erzeugt, wenn es keine Daten gibt
				numRowsTotal++;
			}

			assertEquals("Should not add too many rows", numRowsTotal, getSheet().getLastRowNum() + 1);
			for (int i = 1; i <= getNumRowsTest(); i++) {
				assertEquals("Gates", getVal(getSheet(), getNumRowsBefore() + i, "L"));
			}

			assertEquals("", getVal(getSheet(), numRowsTotal - 1, "A"));
			assertEquals("TOTAL", getVal(getSheet(), numRowsTotal, "A"));
		}

		private void addKind(@Nonnull ExcelMergerDTO excelData) {
			ExcelMergerDTO kind = excelData.createGroup(MergeFieldWarteliste.REPEAT_KIND.getMergeField());
			kind.addValue(MergeFieldWarteliste.NAME, "Gates");
			kind.addValue(MergeFieldWarteliste.VORNAME, "Bill");
			kind.addValue(MergeFieldWarteliste.GEBURTSTAG, LocalDate.now());
			kind.addValue(MergeFieldWarteliste.KITA_BESETZT, true);
			kind.addValue(MergeFieldWarteliste.KITA_BESETZT, false);
			kind.addValue(MergeFieldWarteliste.FIRMA, false);
			kind.addValue(MergeFieldWarteliste.FIRMA, true);
			kind.addValue(MergeFieldWarteliste.PENSUM_WUNSCH_MIN, BigDecimal.valueOf(0.5));
			kind.addValue(MergeFieldWarteliste.PENSUM_WUNSCH_MAX, BigDecimal.valueOf(1));
		}
	}

	private static class NestedGroups extends GroupWithMultipleRows {

		private final int numNestedGroups;

		public NestedGroups(int numGroups, int numNestedGroups, @Nonnull ExcelMerger.GroupMerger merger) {
			super(merger);
			setNumGroups(numGroups);
			this.numNestedGroups = numNestedGroups;
			setNumRowsTest(6 * Math.max(numGroups, 1)
				+ (Math.max(numNestedGroups, 1) - 1) * numGroups); // im Belegungsplan gibt es 6 Zeilen pro Gruppe

			addData();
		}

		@Override
		protected void addData() {
			for (int i = 0; i < getNumGroups(); i++) {
				ExcelMergerDTO group1 = getExcelMergerDTO().createGroup(MergeFieldBelegungsplan.REPEAT_GROUP);
				group1.addValue(MergeFieldBelegungsplan.GRUPPEN_NAME, "Helden");

				for (int j = 0; j < getNumNestedGroups(); j++) {
					ExcelMergerDTO group1kind1 = group1.createGroup(MergeFieldBelegungsplan.REPEAT_KIND);
					group1kind1.addValue(MergeFieldBelegungsplan.NAME, "Tester");
				}
			}
		}

		public int getNumNestedGroups() {
			return numNestedGroups;
		}
	}

	private static class GroupWithMultipleRows extends ContextBuilder {

		private int numGroups;

		public GroupWithMultipleRows(int numGroups, @Nonnull ExcelMerger.GroupMerger merger) {
			this(merger);
			this.numGroups = numGroups;
			setNumRowsTest(6 * Math.max(numGroups, 1)); // im Belegungsplan gibt es 6 Zeilen pro Gruppe

			addData();
		}

		protected GroupWithMultipleRows(@Nonnull ExcelMerger.GroupMerger merger) {
			super(GET_WORKBOOK.apply(BELEGUNGSPLAN), merger);
			setNumRowsBefore(5);
			setNumRowsAfter(5);

			MergeFieldProvider.toMergeFields(MergeFieldBelegungsplan.values())
				.forEach(field -> getFieldMap().put(field.getKey(), field));
		}

		protected void addData() {
			for (int i = 0; i < numGroups; i++) {
				ExcelMergerDTO group1 = getExcelMergerDTO().createGroup(MergeFieldBelegungsplan.REPEAT_GROUP);
				group1.addValue(MergeFieldBelegungsplan.GRUPPEN_NAME, "Helden");
			}
		}

		@Override
		protected void validate() {
			int numRowsTotal = getNumRowsBefore() + getNumRowsTest() + getNumRowsAfter();

			assertEquals("Should not add too many rows", numRowsTotal, getSheet().getLastRowNum() + 1);

			assertEquals("Belegung", getVal(getSheet(), numRowsTotal - 7, "A"));
			assertEquals("Plätze", getVal(getSheet(), numRowsTotal - 6, "A"));
			assertEquals("Max Plätze", getVal(getSheet(), numRowsTotal - 5, "A"));

			assertEquals("", getVal(getSheet(), numRowsTotal - 4, "A"));
			assertEquals("Belegung Kita", getVal(getSheet(), numRowsTotal - 3, "A"));
			assertEquals("Plätze Kita", getVal(getSheet(), numRowsTotal - 2, "A"));
			assertEquals("Max Plätze Kita", getVal(getSheet(), numRowsTotal - 1, "A"));
			assertEquals("Bewilligte Plätze Kita", getVal(getSheet(), numRowsTotal, "A"));
		}

		protected void setNumGroups(int numGroups) {
			this.numGroups = numGroups;
		}

		public int getNumGroups() {
			return numGroups;
		}
	}

	private abstract static class ContextBuilder {
		private int numRowsBefore = 5;
		private int numRowsAfter = 2;
		private int numRowsTest;
		@Nonnull
		private final Map<String, MergeField<?>> fieldMap = new HashMap<>();
		@Nonnull
		private final Workbook workbook;
		@Nonnull
		private final ExcelMergerDTO excelMergerDTO = new ExcelMergerDTO();
		@Nonnull
		private final ExcelMerger.GroupMerger merger;

		protected ContextBuilder(@Nonnull Workbook workbook, @Nonnull ExcelMerger.GroupMerger merger) {
			this.workbook = workbook;
			this.merger = merger;
		}

		protected abstract void validate();

		protected void setNumRowsBefore(int numRowsBefore) {
			this.numRowsBefore = numRowsBefore;
		}

		public int getNumRowsBefore() {
			return numRowsBefore;
		}

		protected void setNumRowsAfter(int numRowsAfter) {
			this.numRowsAfter = numRowsAfter;
		}

		public int getNumRowsAfter() {
			return numRowsAfter;
		}

		protected void setNumRowsTest(int numRowsTest) {
			this.numRowsTest = numRowsTest;
		}

		public int getNumRowsTest() {
			return numRowsTest;
		}

		@Nonnull
		public Map<String, MergeField<?>> getFieldMap() {
			return fieldMap;
		}

		@Nonnull
		public Workbook getWorkbook() {
			return workbook;
		}

		@Nonnull
		public Sheet getSheet() {
			return workbook.getSheetAt(0);
		}

		@Nonnull
		public ExcelMergerDTO getExcelMergerDTO() {
			return excelMergerDTO;
		}

		@Nonnull
		public ExcelMerger.GroupMerger getMerger() {
			return merger;
		}
	}

	private static class Tester extends Context {

		private final int numRowsTest;
		@Nonnull
		private final ExcelMerger.GroupMerger merger;
		@Nonnull
		private final ExcelMergerDTO excelMergerDTO;
		@Nonnull
		private final Procedure<?> validator;
		@Nonnull
		private final Class<? extends ContextBuilder> builderClass;

		Tester(@Nonnull ContextBuilder builder) {
			super(builder.getWorkbook(), builder.getSheet(), builder.getFieldMap(), builder.getNumRowsBefore());
			this.numRowsTest = builder.getNumRowsTest();
			this.excelMergerDTO = builder.getExcelMergerDTO();
			this.merger = builder.getMerger();
			this.validator = builder::validate;
			this.builderClass = builder.getClass();
		}

		/**
		 * @return returns the duration of the merge
		 */
		@SuppressWarnings("UnusedReturnValue")
		public long test(@Nonnull String name) throws Exception {
			Row row = currentRow();
			Optional<GroupPlaceholder> group = detectGroup();
			assertTrue("Group could not be found. Is numRowsBefore wrong?", group.isPresent());

			long start = System.currentTimeMillis();
			ExcelMerger.mergeGroup(this, group.get(), excelMergerDTO, row, merger);
			long duration = System.currentTimeMillis() - start;

			String fileName = String.format("perf_test_%s_%s_sheet_%s_rows_%s.xlsx",
				builderClass.getSimpleName(), name, getSheet().getSheetName(), numRowsTest);

			writeWorkbookToFile(getWorkbook(), fileName);

			this.validator.execute();

			return duration;
		}
	}
}

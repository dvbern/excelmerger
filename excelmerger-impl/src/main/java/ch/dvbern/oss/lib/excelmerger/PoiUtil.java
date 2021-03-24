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
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import javax.annotation.Nonnull;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import static org.apache.poi.ss.SpreadsheetVersion.EXCEL2007;

public final class PoiUtil {

	private static final Logger LOG = LoggerFactory.getLogger(PoiUtil.class);

	private static final Pattern CONTAINS_SPACES = Pattern.compile("\\s");

	private PoiUtil() {
		// utility function
	}

	/*
	 * Workaround für POI >= 3.15: Bei shiftRows gehen die MergedRegions (Cell-Verbindungen) verloren.
	 *
	 * @see <a href="https://bz.apache.org/bugzilla/show_bug.cgi?id=60384">https://bz.apache.org/bugzilla/show_bug
	 * .cgi?id=60384</a>
	 */
	public static void shiftRowsAndMergedRegions(@Nonnull Sheet sheet, int startRow, int endRow, int anzNewRows) {
		List<CellRangeAddress> mergedRegionsBeforeShift = sheet.getMergedRegions();
		List<CellRangeAddress> containedMergedRegions = mergedRegionsBeforeShift.stream()
			.filter(cra -> isContained(cra, startRow, endRow))
			.collect(Collectors.toList());

		// All merged regions within [startRow, startRow + anzNewRows] remain intact,
		// but all further merged regions from [startRow + anzNewRows + 1, endRow] disappear due to shiftRows.
		sheet.shiftRows(startRow, endRow, anzNewRows);

		List<CellRangeAddress> remainingMergedRegions = sheet.getMergedRegions();
		// since the shift was executed, we should check containment with the new endRow: endRow + anzNewRows
		int newEndRow = endRow + anzNewRows;
		List<CellRangeAddress> containedRemainingMergedRegions = sheet.getMergedRegions().stream()
			.filter(cra -> isContained(cra, startRow, newEndRow))
			.collect(Collectors.toList());

		if (mergedRegionsBeforeShift.size() != remainingMergedRegions.size()) {
			// restore lost merged regions
			containedMergedRegions.stream()
				.map(cra -> {
					CellRangeAddress copy = cra.copy();
					copy.setFirstRow(cra.getFirstRow() + anzNewRows);
					copy.setLastRow(cra.getLastRow() + anzNewRows);
					return copy;
				})
				.filter(cra -> !containedRemainingMergedRegions.contains(cra))
				.forEach(sheet::addMergedRegion);
		}

		if (sheet.getMergedRegions().size() != mergedRegionsBeforeShift.size()) {
			LOG.warn("Lost some merged regions in sheet {} when shifting {} rows from {} to {}",
				sheet.getSheetName(), anzNewRows, startRow, endRow);
		}
	}

	public static void shiftDataValidations(@Nonnull Sheet sheet, int startRow, int endRow, int anzNewRows) {
		// shift data validations, for example dropdown constraints
		List<? extends DataValidation> validations = sheet.getDataValidations();

		if (validations.isEmpty() || !(sheet instanceof XSSFSheet)) {
			return;
		}

		validations.stream()
			.filter(validation -> validation.getValidationConstraint().getExplicitListValues() != null)
			.flatMap(validation -> Stream.of(validation.getRegions().getCellRangeAddresses()))
			.filter(cra -> isContained(cra, startRow, endRow))
			.forEach(cra -> {
				cra.setFirstRow(cra.getFirstRow() + anzNewRows);
				cra.setLastRow(cra.getLastRow() + anzNewRows);
			});

		// remove existing data validations
		// FIXME do only remove data validations with explicit list values, since these are the only ones we re-create
		((XSSFSheet) sheet).getCTWorksheet().unsetDataValidations();

		// add the shifted data validations back to the sheet
		DataValidationHelper validationHelper = sheet.getDataValidationHelper();
		for (DataValidation validation : validations) {
			String[] listValues = validation.getValidationConstraint().getExplicitListValues();
			if (listValues == null) {
				continue;
			}
			DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(listValues);
			DataValidation newValidation = validationHelper.createValidation(constraint, validation.getRegions());
			sheet.addValidationData(newValidation);
		}
	}

	private static boolean isContained(@Nonnull CellRangeAddress cra, int startRow, int endRow) {
		return isContained(cra.getFirstRow(), startRow, endRow) && isContained(cra.getLastRow(), startRow, endRow);
	}

	private static boolean isContained(int rowNum, int startRow, int endRow) {
		return startRow <= rowNum && rowNum <= endRow;
	}

	public static void shiftNamedRanges(@Nonnull Sheet sheet, int startRow, int endRow, int anzNewRows) {
		sheet.getWorkbook().getAllNames().stream()
			.filter(name -> name.getRefersToFormula() != null)
			.filter(name -> sheet.getSheetName().equals(name.getSheetName()))
			.filter(name -> {
				AreaReference areaReference = new AreaReference(name.getRefersToFormula(), EXCEL2007);
				return intersects(areaReference, startRow, endRow);
			})
			.forEach(name -> shiftNamedRange(name, anzNewRows));
	}

	private static void shiftNamedRange(@Nonnull Name name, int anzNewRows) {
		AreaReference areaReference = new AreaReference(name.getRefersToFormula(), EXCEL2007);
		CellReference firstCell = areaReference.getFirstCell();
		CellReference lastCell = areaReference.getLastCell();

		String formula = String.format("%s!$%s$%d:$%s$%d", getFormulaSafeSheetName(name),
			CellReference.convertNumToColString(firstCell.getCol()), firstCell.getRow() + 1,
			CellReference.convertNumToColString(lastCell.getCol()), lastCell.getRow() + 1 + anzNewRows
		);

		LOG.debug("formula conversion: {} -> {}", name.getRefersToFormula(), formula);
		name.setRefersToFormula(formula);
	}

	@Nonnull
	private static String getFormulaSafeSheetName(@Nonnull Name name) {
		String sheetName = name.getSheetName();
		return CONTAINS_SPACES.matcher(sheetName).find() ? '\'' + sheetName + '\'' : sheetName;
	}

	private static boolean intersects(@Nonnull AreaReference areaReference, int startRow, int endRow) {
		int firstRow = areaReference.getFirstCell().getRow();
		int lastRow = areaReference.getLastCell().getRow();

		if (firstRow < startRow && lastRow < startRow) {
			return false;
		}

		return !(firstRow > endRow && lastRow > endRow);
	}
}

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

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import javax.annotation.Nonnull;

import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField.Type;
import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import static ch.dvbern.oss.lib.excelmerger.PoiUtil.shiftDataValidations;
import static ch.dvbern.oss.lib.excelmerger.PoiUtil.shiftNamedRanges;
import static ch.dvbern.oss.lib.excelmerger.PoiUtil.shiftRowsAndMergedRegions;

public final class ExcelMerger {

	private static final Logger LOG = LoggerFactory.getLogger(ExcelMerger.class);

	static final Pattern MERGEFIELD_REX = Pattern.compile(".*(\\{([a-zA-Z1-9_]+)(:(\\d+))?}).*");
	static final int REX_GROUP_PATTERN = 1;
	static final int REX_GROUP_KEY = 2;
	static final int REF_GROUP_ROWS = 4;
	// nur ein willkuerlicher Counter, damit's kein while(true) geben muss
	private static final int MAX_PLACEHOLDERS_PER_CELL = 10;
	static final Pattern SAME_ROW_CELL_REF = Pattern.compile(
		"(?:\\b)([a-zA-Z]+)[0-9]+\\b(?!\\()(?=(?:[^\"']*[\"'][^\"']*[\"'])*[^\"']*$)");

	private ExcelMerger() {
		// utliity class
	}

	/**
	 * Fuellt ein Excel-Sheet mit den uebergebenen Daten aus.
	 * Das Sheet wird in Repeat-Gruppen aufgeteilt, die auch verschachtelt sein koennen.
	 * Repeat-Gruppen-Bezeichner ('z.B. {myRepeat}') muessen ein Feld vom Typ {@link Type#REPEAT_ROW} sein.
	 * Normale Felder - (also 1 Wert pro Repeat-Gruppe) sind vom Typ {@link Type#SIMPLE}.
	 * Spalten-Repeater sind vom Typ {@link Type#REPEAT_COL}.
	 * Findet sich in den Daten nicht ausreichend Werte, werden die Spalten ausgeblendet.
	 * Nuetzlich z.B. in Ueberschriften.
	 * Werte-Repeater gehoeren zu Spalten-Repeater und sind die Daten zur Ueberschrift.
	 * Sie unterscheiden sich zu Spalten-Repeatern dadurch, dass sie keine Spalten ausblenden.
	 * => Spalten-Repeater legen die anzahl sichtbarer Spalten (und ggf. defen Ueberschrift) fest und
	 * Werte-Repeater sind die dazugehoerigen Daten die m.o.w. vollstaendig sind.
	 */
	public static void mergeData(
		@Nonnull Sheet sheet,
		@Nonnull MergeField<?>[] fields,
		@Nonnull ExcelMergerDTO excelMergerDTO) throws ExcelMergeException {
		Objects.requireNonNull(sheet);
		Objects.requireNonNull(fields);
		Objects.requireNonNull(excelMergerDTO);

		mergeData(sheet, Arrays.asList(fields), excelMergerDTO);
	}

	public static void mergeData(
		@Nonnull Sheet sheet,
		@Nonnull List<MergeField<?>> fields,
		@Nonnull ExcelMergerDTO excelMergerDTO) throws ExcelMergeException {
		Objects.requireNonNull(sheet);
		Objects.requireNonNull(fields);
		Objects.requireNonNull(excelMergerDTO);

		Map<String, MergeField<?>> fieldMap = fields.stream()
			.collect(Collectors.toMap(MergeField::getKey, field -> field));

		Workbook wb = sheet.getWorkbook();
		Context ctx = new Context(wb, sheet, fieldMap);

		mergeData(excelMergerDTO, ctx);
	}

	public static void mergeData(@Nonnull ExcelMergerDTO excelMergerDTO, @Nonnull Context ctx)
		throws ExcelMergeException {

		mergeGroup(ctx, Collections.singletonList(excelMergerDTO), ctx.getSheet().getLastRowNum() + 1);

		FormulaEvaluator eval = ctx.getWorkbook().getCreationHelper().createFormulaEvaluator();
		eval.clearAllCachedResultValues();
		eval.evaluateAll();

	}

	@Nonnull
	public static Workbook createWorkbookFromTemplate(@Nonnull InputStream is) throws ExcelTemplateParseException {
		Objects.requireNonNull(is);

		try {
			InputStream poiCompatibleIS = toSeekable(is);
			// POI braucht einen Seekable InputStream
			return WorkbookFactory.create(poiCompatibleIS);

		} catch (IOException | RuntimeException | InvalidFormatException e) {
			throw new ExcelTemplateParseException("Error parsing template", e);
		}
	}

	@FunctionalInterface
	private interface TetraConsumer<T, U, V, S> {
		void accept(T a, U b, V c, S s) throws ExcelMergeException;
	}

	@FunctionalInterface
	interface GroupMerger extends TetraConsumer<Context, GroupPlaceholder, List<ExcelMergerDTO>, Row> {

	}

	@Nonnull
	private static InputStream toSeekable(@Nonnull InputStream is) throws IOException {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		IOUtils.copy(is, baos);
		baos.flush();

		return new ByteArrayInputStream(baos.toByteArray());
	}

	static void mergeRow(@Nonnull Context ctx, @Nonnull ExcelMergerDTO data, @Nonnull Row targetRow) {
		Map<MergeField<?>, Integer> valueOffsets = new HashMap<>();
		int start = Math.max(targetRow.getFirstCellNum(), 0);
		int end = Math.max(targetRow.getLastCellNum(), 0);

		IntStream.rangeClosed(start, end)
			.mapToObj(targetRow::getCell)
			.filter(Objects::nonNull)
			.forEach(cell -> mergePlaceholders(ctx, data, valueOffsets, cell));
	}

	private static void mergePlaceholders(
		@Nonnull Context ctx,
		@Nonnull ExcelMergerDTO data,
		@Nonnull Map<MergeField<?>, Integer> valueOffsets,
		@Nonnull Cell cell) {

		for (int i = 0; i < MAX_PLACEHOLDERS_PER_CELL; i++) {
			Optional<Placeholder> placeholderOpt = ctx.parsePlaceholder(cell);
			if (!placeholderOpt.isPresent()) {
				break; // gibt keine Placeholder, da kann sofort abgebrochen werden
			}

			MergeField<?> field = placeholderOpt.get().getField();
			String pattern = placeholderOpt.get().getPattern();

			if (Type.PAGE_BREAK == field.getType()) {
				int rowNum = cell.getRow().getRowNum();
				field.getConverter().setCellValue(cell, pattern, null);
				ctx.getSheet().setRowBreak(rowNum);
				break;
			}

			if (!field.getType().doMergeValue()) {
				break;
			}

			Integer valueOffset = 0;
			if (field.getType().doConsumeValue()) {
				// erhöht den valueOffset (repeat Felder)
				valueOffsets.compute(field, (key, oldVal) -> oldVal == null ? 0 : oldVal + 1);
				valueOffset = valueOffsets.get(field);
			}
			if (data.hasValue(field, valueOffset)) {
				// Schreibt den Wert
				Object value = data.getValue(field, valueOffset);
				field.getConverter().setCellValue(cell, pattern, value);
			} else {
				field.getConverter().setCellValue(cell, pattern, null);
				// Spalte ausblenden
				if (field.getType().doHideColumnOnEmpty()) {
					ctx.getSheet().setColumnHidden(cell.getColumnIndex(), true);
				}
			}
		}
	}

	private static void mergeGroup(@Nonnull Context ctx, @Nonnull List<ExcelMergerDTO> groupRows, int rowSize)
		throws ExcelMergeException {

		for (ExcelMergerDTO dto : groupRows) {
			for (int rowNum = 0; rowNum < rowSize; rowNum++) {
				try {
					Row row = ctx.currentRow();

					Optional<GroupPlaceholder> group = ctx.detectGroup();
					if (group.isPresent()) {
						mergeGroup(ctx, group.get(), dto, row, ExcelMerger::mergeSubGroup);
					} else {
						mergeRow(ctx, dto, ctx.currentRow());
						ctx.advanceRow();
					}
				} catch (RuntimeException rte) {
					throw new ExcelMergeException("Caught error in sheet "
						+ ctx.getSheet().getSheetName()
						+ " on row/col: "
						+ ctx.currentRowNum(), rte);
				}
			}
		}
	}

	static void mergeGroup(
		@Nonnull Context ctx,
		@Nonnull GroupPlaceholder group,
		@Nonnull ExcelMergerDTO dto,
		@Nonnull Row currentRow,
		@Nonnull GroupMerger merger) throws ExcelMergeException {

		List<ExcelMergerDTO> subGroups = dto.getGroup(group.getField());
		group.clearPlaceholder();
		if (subGroups == null) {
			mergeRow(ctx, dto, ctx.currentRow());
			ctx.advanceRow();
		} else {
			merger.accept(ctx, group, subGroups, currentRow);
		}
	}

	static void mergeSubGroup(
		@Nonnull Context ctx,
		@Nonnull GroupPlaceholder group,
		@Nonnull List<ExcelMergerDTO> subGroups,
		@Nonnull Row currentRow) throws ExcelMergeException {

		duplicateRowsWithStylesMultipleRowShift(ctx, currentRow, group.getRows(), subGroups.size());
		mergeGroup(ctx, subGroups, group.getRows());
	}

	/**
	 * Dupliziert Rows:
	 * 1. Platz machen fuer die neuen Rows (i.E.: shift rows)
	 * 2. Zellen inkl. Styles kopieren
	 * 3. Ggf. Named-Ranges um die neuen Zeilen erweitern
	 */
	private static void duplicateRowsWithStylesMultipleRowShift(
		@Nonnull Context ctx,
		@Nonnull Row startRow,
		@Nonnull Integer anzSrcRows,
		@Nonnull Integer anzGroups) {

		final int startNeuerBereich = startRow.getRowNum() + anzSrcRows;
		final int anzRows = anzSrcRows * (anzGroups - 1);

		boolean isXSSFSheet = ctx.getSheet() instanceof XSSFSheet;

		// + 1 ist wichtig, sonst verschwindet beim Filtern die Total-Zeile
		int lastRow = ctx.getSheet().getLastRowNum() + 1;

		// Wenns nach dem zu duplizierenden Bereich noch Zeilen hat: nach unten wegschieben
		if (anzRows > 0 && startNeuerBereich <= lastRow) {
			shiftRowsAndMergedRegions(ctx.getSheet(), startNeuerBereich, lastRow, anzRows);
			// shiftRows does not shift DataValidations or NamedRanges. We have to shift them manually.
			shiftDataValidations(ctx.getSheet(), startNeuerBereich, lastRow + anzRows, anzRows);
			shiftNamedRanges(ctx.getSheet(), startRow.getRowNum(), lastRow, anzRows);
		}

		// Kopieren
		if (isXSSFSheet) {
			copyXssfRows(ctx, startRow, anzSrcRows, anzGroups, startNeuerBereich);
		} else {
			copyRows(ctx, startRow, anzSrcRows, anzGroups, startNeuerBereich);
		}

	}

	/**
	 * Issues
	 * <ul>
	 * <li>Slow!</li>
	 * <li>Beta</li>
	 * <li>XSSF dependent</li>
	 * </ul>
	 */
	private static void copyXssfRows(
		@Nonnull Context ctx,
		@Nonnull Row startRow,
		@Nonnull Integer anzSrcRows,
		@Nonnull Integer anzGroups,
		int startNeuerBereich) {

		XSSFSheet sheet = (XSSFSheet) ctx.getSheet();

		List<XSSFRow> rowsToCopy = IntStream.range(0, anzSrcRows)
			.mapToObj(i -> sheet.getRow(startRow.getRowNum() + i))
			.collect(Collectors.toList());

		for (int i = 0; i < anzGroups - 1; i++) {
			int startGroup = startNeuerBereich + i * anzSrcRows;

			sheet.copyRows(rowsToCopy, startGroup, ctx.getCellCopyPolicy());
		}

	}

	private static void copyRows(
		@Nonnull Context ctx,
		@Nonnull Row startRow,
		@Nonnull Integer anzSrcRows,
		@Nonnull Integer anzGroups,
		int startNeuerBereich) {

		for (int rowNum = 0; rowNum < anzSrcRows; rowNum++) {
			Row srcRow = getRow(ctx.getSheet(), startRow.getRowNum() + rowNum);

			for (int i = 0; i < anzGroups - 1; i++) {
				int startGroup = startNeuerBereich + i * anzSrcRows;
				Row newRow = getRow(ctx.getSheet(), startGroup + rowNum);

				copyCells(srcRow, newRow);
			}
		}
	}

	/**
	 * Issues
	 * <ul>
	 * <li>Does not shift formula references</li>
	 * <li>Does not copy merged cells</li>
	 * </ul>
	 */
	static void copyCells(@Nonnull Row srcRow, @Nonnull Row newRow) {
		IntStream.range(0, srcRow.getLastCellNum())
			.mapToObj(srcRow::getCell)
			.filter(Objects::nonNull)
			.forEach(srcCell -> copyCell(newRow, srcCell));
	}

	private static void copyCell(@Nonnull Row newRow, @Nonnull Cell srcCell) {
		Cell newCell = getCell(newRow, srcCell.getAddress().getColumn());
		newCell.setCellStyle(srcCell.getCellStyle());

		switch (srcCell.getCellTypeEnum()) {
		case STRING:
			newCell.setCellValue(srcCell.getStringCellValue());
			break;
		case NUMERIC:
			newCell.setCellValue(srcCell.getNumericCellValue());
			break;
		case FORMULA:
			String cellFormula = srcCell.getCellFormula();
			//noinspection StringConcatenationMissingWhitespace
			String s = SAME_ROW_CELL_REF.matcher(cellFormula).replaceAll("$1" + (newRow.getRowNum() + 1));
			newCell.setCellFormula(s);
			break;
		case BLANK:
			// nop
			break;
		default:
			LOG.warn("Cell type not supported: {} @{}/{}", srcCell.getCellTypeEnum(), srcCell.getRowIndex(),
				srcCell.getColumnIndex());
		}
	}

	@Nonnull
	private static Row getRow(@Nonnull Sheet sheet, int index) {
		Row row = sheet.getRow(index);

		return row == null ? sheet.createRow(index) : row;
	}

	@Nonnull
	private static Cell getCell(@Nonnull Row row, int column) {
		Cell cell = row.getCell(column);

		return cell == null ? row.createCell(column) : cell;
	}

}

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
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Objects;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import ch.dvbern.oss.lib.excelmerger.StringColorCellDTO;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static ch.dvbern.oss.lib.excelmerger.converters.ConverterUtil.BD_HUNDRED;
import static ch.dvbern.oss.lib.excelmerger.converters.ConverterUtil.BOOLEAN_VALUE;
import static ch.dvbern.oss.lib.excelmerger.converters.ConverterUtil.DEFAULT_DATETIME_FORMAT;
import static ch.dvbern.oss.lib.excelmerger.converters.ConverterUtil.DEFAULT_DATE_FORMAT;
import static ch.dvbern.oss.lib.excelmerger.converters.ConverterUtil.EMPTY_STRING;
import static ch.dvbern.oss.lib.excelmerger.converters.ConverterUtil.SCALE;
import static ch.dvbern.oss.lib.excelmerger.converters.ConverterUtil.writeLocalDate;
import static ch.dvbern.oss.lib.excelmerger.converters.ConverterUtil.writerNumber;

@SuppressWarnings("PMD.ClassNamingConventions")
public final class StandardConverters {

	public static final Converter<String> STRING_CONVERTER =
		(@Nonnull Cell cell, @Nonnull String pattern, @Nullable String value) -> {
			String stringVal = value == null ? EMPTY_STRING : value;
			cell.setCellValue(cell.getStringCellValue().replace(pattern, stringVal));
		};

	public static final Converter<StringColorCellDTO> STRING_COLORED_CONVERTER =
		(@Nonnull Cell cell, @Nonnull String pattern, @Nullable StringColorCellDTO dto) -> {
			if (dto != null) {
				String stringVal = dto.getValue();
				cell.setCellValue(cell.getStringCellValue().replace(pattern, stringVal));

				if (dto.getColor() != null || dto.getFontColor() != null) {
					applyColors(cell, dto);
				}
			}
		};

	@SuppressWarnings("PMD.CloseResource")
	private static void applyColors(@Nonnull Cell cell, @Nonnull StringColorCellDTO dto) {

		XSSFWorkbook wb = (XSSFWorkbook) cell.getSheet().getWorkbook();
		XSSFCellStyle newCellStyle = wb.getStylesSource().createCellStyle();
		newCellStyle.cloneStyleFrom(cell.getCellStyle());

		if (dto.getColor() != null) {
			newCellStyle.setFillForegroundColor(dto.getColor());
			newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}

		if (dto.getFontColor() != null) {
			XSSFFont font = wb.createFont();
			font.setColor(dto.getFontColor());
			newCellStyle.setFont(font);
		}

		if (dto.getBorderColor() != null) {
			newCellStyle.setTopBorderColor(dto.getBorderColor());
			newCellStyle.setRightBorderColor(dto.getBorderColor());
			newCellStyle.setBottomBorderColor(dto.getBorderColor());
			newCellStyle.setLeftBorderColor(dto.getBorderColor());
		}

		if (dto.getBorderStyle() != null) {
			newCellStyle.setBorderTop(dto.getBorderStyle());
			newCellStyle.setBorderRight(dto.getBorderStyle());
			newCellStyle.setBorderBottom(dto.getBorderStyle());
			newCellStyle.setBorderLeft(dto.getBorderStyle());
		}

		cell.setCellStyle(newCellStyle);
	}

	public static final ParametrisedConverter<DateTimeFormatter, LocalDate> LOCAL_DATE_CONVERTER =
		formatter -> (@Nonnull Cell cell, @Nonnull String pattern, @Nullable LocalDate value) ->
			writeLocalDate(cell, pattern, value, formatter);

	public static final Converter<LocalDate> DATE_CONVERTER =
		(@Nonnull Cell cell, @Nonnull String pattern, @Nullable LocalDate value) ->
			LOCAL_DATE_CONVERTER.apply(DEFAULT_DATE_FORMAT).setCellValueImpl(cell, pattern, value);

	public static final ParametrisedConverter<DateTimeFormatter, LocalDateTime> LOCAL_DATETIME_CONVERTER =
		formatter -> (@Nonnull Cell cell, @Nonnull String pattern, @Nullable LocalDateTime dateVal) -> {
			if (pattern.equals(cell.getStringCellValue())) {
				if (dateVal == null) {
					// schade... bei setCellValue(Date) darf kein null uebergeben werden
					cell.setCellValue(EMPTY_STRING);
				} else {
					// ganze Zelle ist Datum -> die Zelle auch als Datum setzen
					Date date = Date.from(dateVal.atZone(ZoneId.systemDefault()).toInstant());
					cell.setCellValue(date);
				}
			} else {
				// nur ein Ausschnitt
				if (dateVal == null) {
					cell.setCellValue(cell.getStringCellValue().replace(pattern, EMPTY_STRING));
				} else {
					cell.setCellValue(cell.getStringCellValue().replace(pattern, dateVal.format(formatter)));
				}
			}
		};

	public static final Converter<LocalDateTime> DATETIME_CONVERTER =
		(@Nonnull Cell cell, @Nonnull String pattern, @Nullable LocalDateTime value) ->
			LOCAL_DATETIME_CONVERTER.apply(DEFAULT_DATETIME_FORMAT).setCellValueImpl(cell, pattern, value);

	public static final Converter<Integer> INTEGER_CONVERTER =
		(@Nonnull Cell cell, @Nonnull String pattern, @Nullable Integer value) ->
			writerNumber(cell, pattern, value, true);

	public static final Converter<Long> LONG_CONVERTER =
		(@Nonnull Cell cell, @Nonnull String pattern, @Nullable Long value) ->
			writerNumber(cell, pattern, value, true);

	public static final Converter<BigDecimal> BIGDECIMAL_CONVERTER =
		(@Nonnull Cell cell, @Nonnull String pattern, @Nullable BigDecimal value) ->
			writerNumber(cell, pattern, value, false);

	public static final Converter<BigDecimal> PERCENT_CONVERTER =
		(@Nonnull Cell cell, @Nonnull String pattern, @Nullable BigDecimal value) -> {
			if (pattern.equals(cell.getStringCellValue())) {
				if (value != null) {
					double doubleValue = value.divide(BD_HUNDRED, SCALE, RoundingMode.HALF_UP).doubleValue();
					cell.setCellValue(doubleValue);
				} else {
					cell.setCellValue(EMPTY_STRING);
				}
			} else {
				cell.setCellValue(cell.getStringCellValue().replace(pattern, value + "%"));
			}
		};

	/**
	 * Converts NULL to false
	 */
	public static final Converter<Boolean> BOOLEAN_CONVERTER =
		(@Nonnull Cell cell, @Nonnull String pattern, @Nullable Boolean value) -> {
			Boolean boolVal = value == null ? Boolean.FALSE : value;

			if (pattern.equals(cell.getStringCellValue())) {
				cell.setCellValue(boolVal);
			} else {
				cell.setCellValue(cell.getStringCellValue().replace(pattern, String.valueOf(boolVal)));
			}
		};

	/**
	 * Writes X when TRUE, otherwise writes empty string
	 */
	public static final Converter<Boolean> BOOLEAN_X_CONVERTER =
		(@Nonnull Cell cell, @Nonnull String pattern, @Nullable Boolean value) -> {
			String stringVal = Objects.equals(Boolean.TRUE, value) ? BOOLEAN_VALUE : EMPTY_STRING;

			STRING_CONVERTER.setCellValueImpl(cell, pattern, stringVal);
		};

	/**
	 * Colors the cell using the given color when TRUE, otherwise does nothing.
	 */
	public static final ParametrisedConverter<XSSFColor, Boolean> CELL_COLORING_CONVERTER =
		color -> (@Nonnull Cell cell, @Nonnull String pattern, @Nullable Boolean value) -> {

			cell.setCellValue(cell.getStringCellValue().replace(pattern, EMPTY_STRING));
			if (Objects.equals(value, Boolean.TRUE)) {
				XSSFWorkbook wb = (XSSFWorkbook) cell.getSheet().getWorkbook();
				XSSFCellStyle newCellStyle = wb.getStylesSource().createCellStyle();
				newCellStyle.cloneStyleFrom(cell.getCellStyle());
				newCellStyle.setFillForegroundColor(color);
				newCellStyle.setFillBackgroundColor(color);
				newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				cell.setCellStyle(newCellStyle);
			}
		};

	/**
	 * Uses a hack, to make Excel automatically adjust row heights (sets row height to -1).
	 * <a href="https://stackoverflow.com/a/35789927">https://stackoverflow.com/a/35789927</a>.
	 * <p>
	 * To make it work, some cells or entire sheet should use automatic wrapping.
	 * <p>
	 * Does not work with LibreOffice: the rows are using standard height.
	 *
	 * @param valueConverter any converter
	 * @param <V> type of the value that's writen into the cell
	 * @return a new converter with negative cell height to trigger automatic sizing
	 */
	public static <V> Converter<V> autoHeightConverter(Converter<V> valueConverter) {
		return (@Nonnull Cell cell, @Nonnull String pattern, @Nullable V value) -> {
			cell.getRow().setHeight((short) -1);
			valueConverter.setCellValue(cell, pattern, value);
		};
	}

	public static final Converter<String> DO_NOTHING_CONVERTER =
		(@Nonnull Cell cell, @Nonnull String pattern, @Nullable String value) -> {
			// nop
		};

	private StandardConverters() {
		// utility class
	}
}

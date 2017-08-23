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

import java.awt.Color;
import java.math.BigDecimal;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import org.apache.poi.ss.usermodel.Cell;

public final class ConverterUtil {

	static final int SCALE = 6;
	static final BigDecimal BD_HUNDRED = BigDecimal.valueOf(100);
	static final Color GAINSBORO = new Color(220, 220, 220);

	public static final DateTimeFormatter DEFAULT_DATE_FORMAT = DateTimeFormatter.ofPattern("dd.MM.yyyy");
	public static final DateTimeFormatter DATE_WITH_DAY_FORMATTER = DateTimeFormatter.ofPattern("EE, dd.MM.yyyy");
	public static final DateTimeFormatter DEFAULT_DATETIME_FORMAT = DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm:ss");
	public static final String BOOLEAN_VALUE = "X";
	public static final String EMPTY_STRING = "";

	public static void writerNumber(
		@Nonnull Cell cell,
		@Nonnull String pattern,
		@Nullable Number number,
		boolean asInteger) {

		if (pattern.equals(cell.getStringCellValue())) {
			if (number == null) {
				cell.setCellValue(EMPTY_STRING);
				return;
			}

			if (asInteger) {
				cell.setCellValue(number.longValue());
				return;
			}

			cell.setCellValue(number.doubleValue());
		} else {
			cell.setCellValue(cell.getStringCellValue().replace(pattern, String.valueOf(number)));
		}
	}

	public static void writeLocalDate(
		@Nonnull Cell cell,
		@Nonnull String pattern,
		@Nullable LocalDate dateVal,
		@Nonnull DateTimeFormatter formatter) {

		if (dateVal == null) {
			if (pattern.equals(cell.getStringCellValue())) {
				cell.setCellValue(EMPTY_STRING);
			} else {
				cell.setCellValue(cell.getStringCellValue().replace(pattern, EMPTY_STRING));
			}

			return;
		}

		Instant instant = dateVal.atStartOfDay().atZone(ZoneId.systemDefault()).toInstant();
		Date date = Date.from(instant);
		if (pattern.equals(cell.getStringCellValue())) {
			cell.setCellValue(date);
		} else {
			cell.setCellValue(cell.getStringCellValue().replace(pattern, dateVal.format(formatter)));
		}
	}

	private ConverterUtil() {
		// utility
	}
}

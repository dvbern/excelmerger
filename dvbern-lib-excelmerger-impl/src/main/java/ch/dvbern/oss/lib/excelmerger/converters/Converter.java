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

import java.io.Serializable;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import ch.dvbern.oss.lib.excelmerger.ExcelMergeRuntimeException;
import org.apache.poi.ss.usermodel.Cell;

@FunctionalInterface
public interface Converter<V> extends Serializable {

	default void setCellValue(@Nonnull Cell cell, @Nonnull String pattern, @Nullable Object o) {
		try {
			//noinspection unchecked
			setCellValueImpl(cell, pattern, (V) o);
		} catch (RuntimeException rte) {
			// Dient nur zum Debugging, damit der Entwickler an row und column rankommt
			String format = "Error converting data on cell %d/%d with pattern %s on object %s";
			String msg = String.format(format, cell.getRowIndex(), cell.getColumnIndex(), pattern, o);
			throw new ExcelMergeRuntimeException(msg, rte); // NOPMD.PreserveStackTrace
		}
	}

	void setCellValueImpl(@Nonnull Cell cell, @Nonnull String pattern, @Nullable V o);
}

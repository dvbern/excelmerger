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

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import ch.dvbern.oss.lib.excelmerger.mergefields.RepeatRowMergeField;
import com.google.common.base.MoreObjects;
import org.apache.poi.ss.usermodel.Cell;

class GroupPlaceholder extends Placeholder {

	private final int rows;

	GroupPlaceholder(
		@Nonnull Cell cell,
		@Nonnull String pattern,
		@Nonnull String key,
		@Nonnull RepeatRowMergeField field,
		@Nullable Integer rowsParsed) {

		super(cell, pattern, key, field);
		this.rows = rowsParsed == null ? 1 : rowsParsed;
	}

	public int getRows() {
		return rows;
	}

	@Nonnull
	@Override
	public RepeatRowMergeField getField() {
		return (RepeatRowMergeField) super.getField();
	}

	/**
	 * Group-Repeat-Info aus der Zelle loeschen
	 */
	public void clearPlaceholder() {
		getCell().setCellValue((String) null);
	}

	@Override
	@Nonnull
	public String toString() {
		return MoreObjects.toStringHelper(this)
			.add("pattern", getPattern())
			.add("key", getKey())
			.add("field", getField())
			.add("rows", rows)
			.toString();
	}
}

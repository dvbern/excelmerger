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

import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField;
import com.google.common.base.MoreObjects;
import org.apache.poi.ss.usermodel.Cell;

import static com.google.common.base.Preconditions.checkNotNull;

class Placeholder {
	@Nonnull
	private final Cell cell;
	@Nonnull
	private final String pattern;
	@Nonnull
	private final String key;
	@Nonnull
	private final MergeField<?> field;

	Placeholder(@Nonnull Cell cell, @Nonnull String pattern, @Nonnull String key, @Nonnull MergeField<?> field) {
		this.cell = checkNotNull(cell);
		this.pattern = checkNotNull(pattern);
		this.key = checkNotNull(key);
		this.field = field;
	}

	@Nonnull
	public Cell getCell() {
		return cell;
	}

	@Nonnull
	public String getPattern() {
		return pattern;
	}

	@Nonnull
	public String getKey() {
		return key;
	}

	@Nonnull
	public MergeField<?> getField() {
		return field;
	}

	@Override
	@Nonnull
	public String toString() {
		return MoreObjects.toStringHelper(this)
			.add("pattern", pattern)
			.add("key", key)
			.add("field", field)
			.toString();
	}
}

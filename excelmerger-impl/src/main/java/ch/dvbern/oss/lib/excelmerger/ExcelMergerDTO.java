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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField.Type;
import ch.dvbern.oss.lib.excelmerger.mergefields.MergeFieldProvider;
import ch.dvbern.oss.lib.excelmerger.mergefields.RepeatRowMergeField;
import com.google.common.base.Preconditions;

import static com.google.common.base.Preconditions.checkNotNull;

public class ExcelMergerDTO {

	/**
	 * Zellen, die im globalen Teil des Excel wiederholt werden (z.B. Ueberschriften mit Firmennamen)
	 */
	@Nonnull
	private final Map<MergeField<?>, List<Object>> values = new HashMap<>();
	private final Map<MergeField<?>, List<ExcelMergerDTO>> groups = new HashMap<>();

	@Nonnull
	public <V> ExcelMergerDTO createGroup(@Nonnull MergeFieldProvider provider) {
		return createGroup(provider.getMergeField());
	}

	@Nonnull
	public <V> ExcelMergerDTO createGroup(@Nonnull MergeField<V> group) {
		checkNotNull(group);

		Preconditions.checkArgument(
			group.getType() == Type.REPEAT_ROW,
			"Not a REPEAT_ROW type %" + group.getType());

		if (groups.containsKey(group)) {
			groups.get(group).add(new ExcelMergerDTO());
		} else {
			List<ExcelMergerDTO> l = new LinkedList<>();
			l.add(new ExcelMergerDTO());
			groups.put(group, l);
		}

		List<ExcelMergerDTO> entries = groups.computeIfAbsent(group, key -> new LinkedList<>());
		ExcelMergerDTO newGroup = new ExcelMergerDTO();
		entries.add(newGroup);

		return newGroup;
	}

	<V> void createGroup(@Nonnull MergeField<V> group, int numberOfEntries) {
		List<ExcelMergerDTO> entries = groups.computeIfAbsent(group, key -> new ArrayList<>(numberOfEntries));

		IntStream.range(0, numberOfEntries)
			.mapToObj(i -> new ExcelMergerDTO())
			.collect(Collectors.toCollection(() -> entries));
	}

	public <V> void addValue(@Nonnull MergeFieldProvider provider, @Nullable V value) {
		addValue(provider.getMergeField(), value);
	}

	public <V> void addValue(@Nonnull MergeField<V> mergeField, @Nullable V value) {
		checkNotNull(mergeField);

		List<Object> valuesList = values.computeIfAbsent(mergeField, key -> new LinkedList<>());
		valuesList.add(value);
	}

	@Nullable
	public List<ExcelMergerDTO> getGroup(@Nonnull RepeatRowMergeField groupField) {
		return groups.get(groupField);
	}

	public <V> boolean hasValue(@Nonnull MergeField<V> mergeField, int valueOffset) {
		List<Object> list = values.get(mergeField);
		if (list == null) {
			return false;
		}

		return valueOffset < list.size();
	}

	@Nullable
	public <V> V getValue(@Nonnull MergeField<V> mergeField) {
		return getValue(mergeField, 0);
	}

	@Nullable
	public <V> V getValue(@Nonnull MergeField<V> mergeField, int valueOffset) {
		List<Object> list = values.get(mergeField);
		if (list == null) {
			return null;
		}

		if (valueOffset < list.size()) {
			//noinspection unchecked
			return (V) list.get(valueOffset);
		}

		return null;
	}
}

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
package ch.dvbern.oss.lib.excelmerger.mergefields;

import javax.annotation.Nonnull;

import ch.dvbern.oss.lib.excelmerger.converters.Converter;
import com.google.common.base.MoreObjects;

public class DefaultMergeField<V> implements MergeField<V> {

	private static final long serialVersionUID = -4600129143718111847L;

	@Nonnull
	private final String key;

	@Nonnull
	private final Type type;

	@Nonnull
	private final Converter<V> converter;

	public DefaultMergeField(
		@Nonnull String key,
		@Nonnull Type type,
		@Nonnull Converter<V> converter) {

		this.key = key;
		this.type = type;
		this.converter = converter;
	}

	@Nonnull
	@Override
	public String getKey() {
		return key;
	}

	@Nonnull
	@Override
	public Type getType() {
		return type;
	}

	@Nonnull
	@Override
	public Converter<V> getConverter() {
		return converter;
	}

	@Override
	@Nonnull
	public String toString() {
		return MoreObjects.toStringHelper(this)
			.add("key", key)
			.add("type", type)
			.toString();
	}
}

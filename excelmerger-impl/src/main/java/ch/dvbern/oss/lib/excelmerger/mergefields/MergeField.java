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

import java.io.Serializable;

import javax.annotation.Nonnull;

import ch.dvbern.oss.lib.excelmerger.converters.Converter;

public interface MergeField<V> extends Serializable {
	@Nonnull
	String getKey();

	@Nonnull
	Type getType();

	@Nonnull
	Converter<V> getConverter();

	enum Type {
		/**
		 * Ein einfacher Platzhalter
		 */
		SIMPLE(true, false, false),
		/**
		 * Ein Platzhalter in den Ueberschriften, der mehrere Spalten hat (z.B. Ueberschrift mit den Kita-Namen)
		 */
		REPEAT_COL(true, true, true),
		REPEAT_VAL(true, true, false),
		/**
		 * Kennzeichnet eine Excel-Row, die wiederholt werden soll
		 */
		REPEAT_ROW(false, false, false),
		/**
		 * Fügt einen Seitenumbruch ein
		 */
		PAGE_BREAK(true, true, true);

		private final boolean mergeValue;
		private final boolean consumesValue;
		private final boolean hideColumOnEmpty;

		Type(boolean mergeValue, boolean consumesValue, boolean hideColumOnEmpty) {
			this.mergeValue = mergeValue;
			this.consumesValue = consumesValue;
			this.hideColumOnEmpty = hideColumOnEmpty;
		}

		public boolean doMergeValue() {
			return mergeValue;
		}

		public boolean doConsumeValue() {
			return consumesValue;
		}

		public boolean doHideColumnOnEmpty() {
			return hideColumOnEmpty;
		}
	}
}

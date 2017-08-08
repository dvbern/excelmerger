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

import static ch.dvbern.oss.lib.excelmerger.converters.StandardConverters.DO_NOTHING_CONVERTER;

public class RepeatRowMergeField extends DefaultMergeField<String> {

	private static final long serialVersionUID = -3896240820489518571L;

	public RepeatRowMergeField(@Nonnull String key) {
		super(key, Type.REPEAT_ROW, DO_NOTHING_CONVERTER);
	}
}

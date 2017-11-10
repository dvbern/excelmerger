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

import java.time.LocalDateTime;
import java.util.List;

import static ch.dvbern.oss.lib.excelmerger.converters.StandardConverters.BOOLEAN_CONVERTER;
import static ch.dvbern.oss.lib.excelmerger.converters.StandardConverters.DATETIME_CONVERTER;
import static ch.dvbern.oss.lib.excelmerger.converters.StandardConverters.INTEGER_CONVERTER;
import static ch.dvbern.oss.lib.excelmerger.converters.StandardConverters.STRING_CONVERTER;

public final class StaticMergeFields {
	static final MergeField<String> REPEAT_ROW = new RepeatRowMergeField("repeatRow");
	static final RepeatColMergeField<String> REPEAT_COL = new RepeatColMergeField<>("repeatCol", STRING_CONVERTER);
	static final RepeatValMergeField<Boolean> REPEAT_VAL = new RepeatValMergeField<>("repeatVal", BOOLEAN_CONVERTER);
	static final SimpleMergeField<LocalDateTime> SIMPLE = new SimpleMergeField<>("simple", DATETIME_CONVERTER);
	static final MergeField<Integer> GENERIC = new SimpleMergeField<>("generic", INTEGER_CONVERTER);

	static final List<MergeField<?>> ALL_FIELDS = MergeField.getStaticMergeFields(StaticMergeFields.class);

	private StaticMergeFields() {
	}
}

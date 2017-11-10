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

import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class StaticMergeFieldsTest {

	@Test
	public void test() {
		assertTrue(StaticMergeFields.ALL_FIELDS.contains(StaticMergeFields.REPEAT_ROW));
		assertTrue(StaticMergeFields.ALL_FIELDS.contains(StaticMergeFields.REPEAT_COL));
		assertTrue(StaticMergeFields.ALL_FIELDS.contains(StaticMergeFields.REPEAT_VAL));
		assertTrue(StaticMergeFields.ALL_FIELDS.contains(StaticMergeFields.SIMPLE));
		assertTrue(StaticMergeFields.ALL_FIELDS.contains(StaticMergeFields.GENERIC));
		assertEquals(5, StaticMergeFields.ALL_FIELDS.size());
	}
}

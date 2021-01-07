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

import org.junit.jupiter.api.Test;

import static ch.dvbern.oss.lib.excelmerger.ExcelMerger.SAME_ROW_CELL_REF;
import static org.junit.jupiter.api.Assertions.assertEquals;

public class FormulaShiftingTest {

	@Test
	public void testFormulaReferenceUpdate() {
		// all cell references for Row1 should be substituted with references to Row100
		String subst = "$1100";

		assertEquals("SUM(A100:B100)", SAME_ROW_CELL_REF.matcher("SUM(A1:B1)").replaceAll(subst));
		assertEquals("A100+A100*3", SAME_ROW_CELL_REF.matcher("A1+A1*3").replaceAll(subst));
		assertEquals("$A$1", SAME_ROW_CELL_REF.matcher("$A$1").replaceAll(subst));
		assertEquals("$A100", SAME_ROW_CELL_REF.matcher("$A1").replaceAll(subst));
		assertEquals("ABD100+A100", SAME_ROW_CELL_REF.matcher("ABD1+A1").replaceAll(subst));
		assertEquals("A1B1", SAME_ROW_CELL_REF.matcher("A1B1").replaceAll(subst));
		assertEquals("CONCAT(A100; \"A1\")", SAME_ROW_CELL_REF.matcher("CONCAT(A1; \"A1\")").replaceAll(subst));
	}
}

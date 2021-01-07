package ch.dvbern.oss.lib.excelmerger;

import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.RepeatRowMergeField;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import static org.easymock.EasyMock.expect;
import static org.easymock.EasyMock.mock;
import static org.easymock.EasyMock.niceMock;
import static org.easymock.EasyMock.replay;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ContextTest {

	@Test
	public void testParsePlaceholder() {
		Sheet sheet = mock(Sheet.class);
		Cell cell = mock(Cell.class);

		String key = "something";
		int value = 123;
		RepeatRowMergeField mergeField = new RepeatRowMergeField(key);
		Map<String, MergeField<?>> fields = new HashMap<>();
		fields.put(key, mergeField);

		expect(sheet.getFirstRowNum()).andReturn(0);
		expect(cell.getCellTypeEnum()).andReturn(CellType.STRING);
		expect(cell.getStringCellValue()).andReturn("{{" + key + ':' + value + "}}").anyTimes();

		replay(cell, sheet);

		Context context = new Context(niceMock(Workbook.class), sheet, fields);
		Optional<Placeholder> result = context.parsePlaceholder(cell);

		assertTrue(result.isPresent());
		assertTrue(result.get() instanceof GroupPlaceholder);
		GroupPlaceholder resultValue = (GroupPlaceholder) result.get();
		assertEquals(value, resultValue.getRows());
	}
}

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

import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

import ch.dvbern.oss.lib.excelmerger.mergefields.MergeFieldBelegungsplan;
import ch.dvbern.oss.lib.excelmerger.mergefields.MergeFieldProvider;
import ch.dvbern.oss.lib.excelmerger.mergefields.MergeFieldWarteliste;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.BELEGUNGSPLAN;
import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.GET_WORKBOOK;
import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.WARTELISTE;
import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.getNumVal;
import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.getVal;
import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.named;
import static ch.dvbern.oss.lib.excelmerger.ExcelMergerTestUtil.writeWorkbookToFile;
import static ch.dvbern.oss.lib.excelmerger.converters.ConverterUtil.DEFAULT_DATE_FORMAT;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ExcelMergerTest {

	@Test
	public void testWarteliste() throws ExcelMergeException, IOException, InvalidFormatException {
		Workbook wb = GET_WORKBOOK.apply(WARTELISTE);
		LocalDate stichtag = LocalDate.now();

		ExcelMergerDTO excelData = new ExcelMergerDTO();
		excelData.addValue(MergeFieldWarteliste.KITA_NAME, "Testing");
		excelData.addValue(MergeFieldWarteliste.DATUM_AUSWERTUNG, LocalDate.now());
		excelData.addValue(MergeFieldWarteliste.BETREUUNGSFAKTOR_DATUM, LocalDate.now());

		excelData.addValue(MergeFieldWarteliste.REPEAT_KITA, "Kita 1");
		excelData.addValue(MergeFieldWarteliste.REPEAT_KITA, "Kita 2");

		excelData.addValue(MergeFieldWarteliste.REPEAT_FIRMA, "Hammerschmiede");
		excelData.addValue(MergeFieldWarteliste.REPEAT_FIRMA, "Schnitzelklopfer");

		excelData.addValue(MergeFieldWarteliste.REPEAT_MONTAG, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_MONTAG, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_MONTAG, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_DIENSTAG, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_DIENSTAG, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_DIENSTAG, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_MITTWOCH, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_MITTWOCH, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_MITTWOCH, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_DONNERSTAG, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_DONNERSTAG, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_DONNERSTAG, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_FREITAG, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_FREITAG, "");
		excelData.addValue(MergeFieldWarteliste.REPEAT_FREITAG, "");

		ExcelMergerDTO kind1 = excelData.createGroup(MergeFieldWarteliste.REPEAT_KIND);
		kind1.addValue(MergeFieldWarteliste.NAME, "Tester");
		kind1.addValue(MergeFieldWarteliste.VORNAME, "Thomas");
		kind1.addValue(MergeFieldWarteliste.GEBURTSTAG, LocalDate.now());
		kind1.addValue(MergeFieldWarteliste.KITA_BESETZT, true);
		kind1.addValue(MergeFieldWarteliste.KITA_BESETZT, false);
		kind1.addValue(MergeFieldWarteliste.FIRMA, false);
		kind1.addValue(MergeFieldWarteliste.FIRMA, false);
		kind1.addValue(MergeFieldWarteliste.PENSUM_WUNSCH_MIN, BigDecimal.valueOf(0.5));
		kind1.addValue(MergeFieldWarteliste.PENSUM_WUNSCH_MAX, BigDecimal.valueOf(1));

		ExcelMergerDTO kind2 = excelData.createGroup(MergeFieldWarteliste.REPEAT_KIND);
		kind2.addValue(MergeFieldWarteliste.NAME, "Lovelace");
		kind2.addValue(MergeFieldWarteliste.VORNAME, "Ada");
		kind2.addValue(MergeFieldWarteliste.GEBURTSTAG, LocalDate.now());
		kind2.addValue(MergeFieldWarteliste.KITA_BESETZT, false);
		kind2.addValue(MergeFieldWarteliste.KITA_BESETZT, true);
		kind2.addValue(MergeFieldWarteliste.FIRMA, true);
		kind2.addValue(MergeFieldWarteliste.FIRMA, false);
		kind2.addValue(MergeFieldWarteliste.PENSUM_WUNSCH_MIN, BigDecimal.valueOf(0.1));
		kind2.addValue(MergeFieldWarteliste.PENSUM_WUNSCH_MAX, BigDecimal.valueOf(0.20));

		ExcelMergerDTO kind3 = excelData.createGroup(MergeFieldWarteliste.REPEAT_KIND);
		kind3.addValue(MergeFieldWarteliste.NAME, "Schneider");
		kind3.addValue(MergeFieldWarteliste.VORNAME, "Helge");
		kind3.addValue(MergeFieldWarteliste.GEBURTSTAG, LocalDate.now());
		kind3.addValue(MergeFieldWarteliste.KITA_BESETZT, false);
		kind3.addValue(MergeFieldWarteliste.KITA_BESETZT, true);
		kind3.addValue(MergeFieldWarteliste.FIRMA, false);
		kind3.addValue(MergeFieldWarteliste.FIRMA, true);
		kind3.addValue(MergeFieldWarteliste.PENSUM_WUNSCH_MIN, BigDecimal.valueOf(0.50));
		kind3.addValue(MergeFieldWarteliste.PENSUM_WUNSCH_MAX, BigDecimal.valueOf(0.50));

		Sheet sheet = wb.getSheet("Warteliste");
		assertNotNull(sheet);

		ExcelMerger.mergeData(sheet, MergeFieldProvider.toMergeFields(MergeFieldWarteliste.values()), excelData);

		writeWorkbookToFile(wb, "warteliste-filled.xlsx");

		assertEquals(9, sheet.getLastRowNum()); // keine ueberfluessigen Zeilen hinzugefuegt

		// globale Felder
		assertEquals("Warteliste Testing", getVal(sheet, 1, "A"));
		assertEquals("Stand: " + stichtag.format(DEFAULT_DATE_FORMAT), getVal(sheet, 2, "A"));

		// REPEAT_COL
		assertEquals("Kita 1", getVal(sheet, 5, "B"));
		assertEquals("Kita 2", getVal(sheet, 5, "C"));

		// Kinderdaten
		assertEquals("Tester", getVal(sheet, 6, "L"));
		assertEquals("Thomas", getVal(sheet, 6, "M"));
		assertEquals("X", getVal(sheet, 6, "B"));
		assertEquals("", getVal(sheet, 6, "C"));
		assertEquals("Lovelace", getVal(sheet, 7, "L"));
		assertEquals("Ada", getVal(sheet, 7, "M"));
		assertEquals("", getVal(sheet, 7, "B"));
		assertEquals("X", getVal(sheet, 7, "C"));
		assertEquals("Schneider", getVal(sheet, 8, "L"));
		assertEquals("Helge", getVal(sheet, 8, "M"));
		assertEquals("", getVal(sheet, 8, "B"));
		assertEquals("X", getVal(sheet, 8, "C"));

		// Excel-Formel: zaehlenwenn
		assertEquals(1.0, getNumVal(sheet, 10, "B"), 0.0);
		assertEquals(2.0, getNumVal(sheet, 10, "C"), 0.0);
		// Excel-Formel: Summe
		assertEquals(1.10, getNumVal(sheet, 10, "DL"), 0.0);
		assertEquals(1.70, getNumVal(sheet, 10, "DM"), 0.0);

		// ausgeblendete Spalten (REPEAT_COL)
		// wir haben 2 Kitas
		assertFalse(sheet.isColumnHidden(named("B")));
		assertFalse(sheet.isColumnHidden(named("C")));
		assertTrue(sheet.isColumnHidden(named("D")));
		assertTrue(sheet.isColumnHidden(named("E")));
		assertTrue(sheet.isColumnHidden(named("F")));
		assertTrue(sheet.isColumnHidden(named("G")));
		assertTrue(sheet.isColumnHidden(named("H")));
		assertTrue(sheet.isColumnHidden(named("I")));
		assertTrue(sheet.isColumnHidden(named("J")));
		assertTrue(sheet.isColumnHidden(named("K")));
	}

	@Test
	public void testBelegungsplan() throws IOException, ExcelMergeException {
		Workbook wb = GET_WORKBOOK.apply(BELEGUNGSPLAN);

		LocalDate stichtag = LocalDate.now();

		ExcelMergerDTO excelData = new ExcelMergerDTO();
		excelData.addValue(MergeFieldBelegungsplan.KITA_NAME, "Testkita");
		DateTimeFormatter kalenderwocheFormatter = DateTimeFormatter.ofPattern("'KW' w");
		excelData.addValue(MergeFieldBelegungsplan.KALENDERWOCHE, stichtag.format(kalenderwocheFormatter));

		{
			ExcelMergerDTO group1 = excelData.createGroup(MergeFieldBelegungsplan.REPEAT_GROUP);
			group1.addValue(MergeFieldBelegungsplan.GRUPPEN_NAME, "Helden");
			ExcelMergerDTO group1kind1 = group1.createGroup(MergeFieldBelegungsplan.REPEAT_KIND);
			group1kind1.addValue(MergeFieldBelegungsplan.NAME, "Tester");
			ExcelMergerDTO group1kind2 = group1.createGroup(MergeFieldBelegungsplan.REPEAT_KIND);
			group1kind2.addValue(MergeFieldBelegungsplan.NAME, "Lovelace");
		}

		{
			ExcelMergerDTO group2 = excelData.createGroup(MergeFieldBelegungsplan.REPEAT_GROUP);
			group2.addValue(MergeFieldBelegungsplan.GRUPPEN_NAME, "Bastler");
			ExcelMergerDTO group2kind1 = group2.createGroup(MergeFieldBelegungsplan.REPEAT_KIND);
			group2kind1.addValue(MergeFieldBelegungsplan.NAME, "Honolulu");
			ExcelMergerDTO group2kind2 = group2.createGroup(MergeFieldBelegungsplan.REPEAT_KIND);
			group2kind2.addValue(MergeFieldBelegungsplan.NAME, "Baumeister");
			ExcelMergerDTO group2kind3 = group2.createGroup(MergeFieldBelegungsplan.REPEAT_KIND);
			group2kind3.addValue(MergeFieldBelegungsplan.NAME, "Rembremmerdinger");
		}

		Sheet sheet = wb.getSheet("Belegungsplan");
		assertNotNull(sheet);
		ExcelMerger.mergeData(sheet, MergeFieldProvider.toMergeFields(MergeFieldBelegungsplan.values()), excelData);

		writeWorkbookToFile(wb, "belegungsplan-filled.xlsx");
	}
}

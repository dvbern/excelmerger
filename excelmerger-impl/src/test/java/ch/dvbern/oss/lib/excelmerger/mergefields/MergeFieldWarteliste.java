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

import ch.dvbern.oss.lib.excelmerger.mergefields.MergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.MergeFieldProvider;
import ch.dvbern.oss.lib.excelmerger.mergefields.RepeatColMergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.RepeatRowMergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.RepeatValMergeField;
import ch.dvbern.oss.lib.excelmerger.mergefields.SimpleMergeField;

import static ch.dvbern.oss.lib.excelmerger.converters.StandardConverters.BIGDECIMAL_CONVERTER;
import static ch.dvbern.oss.lib.excelmerger.converters.StandardConverters.BOOLEAN_X_CONVERTER;
import static ch.dvbern.oss.lib.excelmerger.converters.StandardConverters.DATE_CONVERTER;
import static ch.dvbern.oss.lib.excelmerger.converters.StandardConverters.LONG_CONVERTER;
import static ch.dvbern.oss.lib.excelmerger.converters.StandardConverters.STRING_CONVERTER;

public enum MergeFieldWarteliste implements MergeFieldProvider {
	REPEAT_KIND(new RepeatRowMergeField("repeatKind")),

	KITA_NAME(new SimpleMergeField<>("kitaName", STRING_CONVERTER)),
	DATUM_AUSWERTUNG(new SimpleMergeField<>("datumAuswertung", DATE_CONVERTER)),
	BETREUUNGSFAKTOR_DATUM(new SimpleMergeField<>("betreuungsfaktorDatum", DATE_CONVERTER)),

	REPEAT_KITA(new RepeatColMergeField<>("repeatKita", STRING_CONVERTER)),
	KITA_BESETZT(new RepeatValMergeField<>("kitaBesetzt", BOOLEAN_X_CONVERTER)),

	VORNAME(new SimpleMergeField<>("vorname", STRING_CONVERTER)),
	NAME(new SimpleMergeField<>("name", STRING_CONVERTER)),
	GESCHLECHT(new SimpleMergeField<>("geschlecht", STRING_CONVERTER)),
	GEBURTSTAG(new SimpleMergeField<>("geburtstag", DATE_CONVERTER)),
	GEBURTSTERMIN(new SimpleMergeField<>("geburtstermin", DATE_CONVERTER)),
	BETREUUNGSFAKTOR(new SimpleMergeField<>("betreuungsfaktor", BIGDECIMAL_CONVERTER)),
	POLITISCHE_GEMEINDE(new SimpleMergeField<>("politischeGemeinde", STRING_CONVERTER)),
	GESCHWISTER(new SimpleMergeField<>("geschwister", BOOLEAN_X_CONVERTER)),
	PENSUM_MIN(new SimpleMergeField<>("pensumMin", BOOLEAN_X_CONVERTER)),
	PENSUM_MAX(new SimpleMergeField<>("pensumMax", BOOLEAN_X_CONVERTER)),

	REPEAT_MONTAG(new RepeatColMergeField<>("repeatMontag", STRING_CONVERTER)),
	MONTAG(new RepeatValMergeField<>("montag", BOOLEAN_X_CONVERTER)),
	REPEAT_DIENSTAG(new RepeatColMergeField<>("repeatDienstag", STRING_CONVERTER)),
	DIENSTAG(new RepeatValMergeField<>("dienstag", BOOLEAN_X_CONVERTER)),
	REPEAT_MITTWOCH(new RepeatColMergeField<>("repeatMittwoch", STRING_CONVERTER)),
	MITTWOCH(new RepeatValMergeField<>("mittwoch", BOOLEAN_X_CONVERTER)),
	REPEAT_DONNERSTAG(new RepeatColMergeField<>("repeatDonnerstag", STRING_CONVERTER)),
	DONNERSTAG(new RepeatValMergeField<>("donnerstag", BOOLEAN_X_CONVERTER)),
	REPEAT_FREITAG(new RepeatColMergeField<>("repeatFreitag", STRING_CONVERTER)),
	FREITAG(new RepeatValMergeField<>("freitag", BOOLEAN_X_CONVERTER)),
	REPEAT_SAMSTAG(new RepeatColMergeField<>("repeatSamstag", STRING_CONVERTER)),
	SAMSTAG(new RepeatValMergeField<>("samstag", BOOLEAN_X_CONVERTER)),
	REPEAT_SONNTAG(new RepeatColMergeField<>("repeatSonntag", STRING_CONVERTER)),
	SONNTAG(new RepeatValMergeField<>("sonntag", BOOLEAN_X_CONVERTER)),

	ANMELDE_DATUM(new SimpleMergeField<>("anmeldedatum", DATE_CONVERTER)),
	BETREUUNGSWUNSCH_AB(new SimpleMergeField<>("betreuungswunschAb", DATE_CONVERTER)),
	PRIORITAET(new SimpleMergeField<>("prioritaet", LONG_CONVERTER)),
	SUBVENTIONIERTER_PLATZ(new SimpleMergeField<>("subventionierterPlatz", BOOLEAN_X_CONVERTER)),
	PRIVATER_PLATZ(new SimpleMergeField<>("privaterPlatz", BOOLEAN_X_CONVERTER)),
	KINDERGARTEN(new SimpleMergeField<>("kindergarten", BOOLEAN_X_CONVERTER)),

	REPEAT_FIRMA(new RepeatColMergeField<>("repeatFirma", STRING_CONVERTER)),
	FIRMA(new RepeatValMergeField<>("firma", BOOLEAN_X_CONVERTER)),

	REPEAT_BETREUUNGSGRUND(new RepeatColMergeField<>("repeatBetreuungsgrund", STRING_CONVERTER)),
	BETREUUNGSGRUND(new RepeatValMergeField<>("betreuungsGrund", BOOLEAN_X_CONVERTER)),

	PENSUM_WUNSCH_MIN(new SimpleMergeField<>("pensumWunschMin", BIGDECIMAL_CONVERTER)),
	PENSUM_WUNSCH_MAX(new SimpleMergeField<>("pensumWunschMax", BIGDECIMAL_CONVERTER)),
	AKTUELLE_BELEGUNG(new SimpleMergeField<>("aktuelleBelegung", BIGDECIMAL_CONVERTER)),

	AKTUELLE_KITA(new SimpleMergeField<>("aktuelleKita", STRING_CONVERTER)),
	AKTUELLE_KITAKIND(new SimpleMergeField<>("aktuelleKitaKind", STRING_CONVERTER)),

	REPEAT_KONTAKTPERSON(new RepeatColMergeField<>("repeatKontaktperson", STRING_CONVERTER)),
	KONTAKTPER_SONVORNAME(new SimpleMergeField<>("kontaktpersonVorname", STRING_CONVERTER)),
	KONTAKTPER_SONNAME(new SimpleMergeField<>("kontaktpersonName", STRING_CONVERTER)),
	KONTAKTPER_SONSTRASSE(new SimpleMergeField<>("kontaktpersonStrasse", STRING_CONVERTER)),
	KONTAKTPER_SONNR(new SimpleMergeField<>("kontaktpersonNr", STRING_CONVERTER)),
	KONTAKTPER_SONPLZ(new SimpleMergeField<>("kontaktpersonPlz", STRING_CONVERTER)),
	KONTAKTPER_SONORT(new SimpleMergeField<>("kontaktpersonOrt", STRING_CONVERTER)),
	KONTAKTPER_SONTELEFON(new SimpleMergeField<>("kontaktpersonTelefon", STRING_CONVERTER)),
	KONTAKTPER_SONEMAIL(new SimpleMergeField<>("kontaktpersonEmail", STRING_CONVERTER)),
	KONTAKTPER_SONFULLNAME(new SimpleMergeField<>("kontaktpersonFullName", STRING_CONVERTER)),
	BEMERKUNG(new SimpleMergeField<>("bemerkung", STRING_CONVERTER)),

	// Kitas sheet
	KITAS_REPEAT_KITA(new RepeatRowMergeField("kitasRepeatKita")),
	KITAS_KITA(new SimpleMergeField<>("kitasKita", STRING_CONVERTER));

	@Nonnull
	private final MergeField<?> mergeField;

	<V> MergeFieldWarteliste(@Nonnull MergeField<V> mergeField) {
		this.mergeField = mergeField;
	}

	@Override
	@Nonnull
	public <V> MergeField<V> getMergeField() {
		//noinspection unchecked
		return (MergeField<V>) mergeField;
	}
}

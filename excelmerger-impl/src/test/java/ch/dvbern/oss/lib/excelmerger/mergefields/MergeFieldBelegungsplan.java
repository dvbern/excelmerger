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

public enum MergeFieldBelegungsplan implements MergeFieldProvider {

	KITA_NAME(new SimpleMergeField<>("kitaName", STRING_CONVERTER)),
	KALENDERWOCHE(new SimpleMergeField<>("kalenderwoche", STRING_CONVERTER)),
	BETREUUNGSFAKTOR_DATUM(new SimpleMergeField<>("betreuungsfaktorDatum", DATE_CONVERTER)),

	REPEAT_GROUP(new RepeatRowMergeField("repeatGroup")),
	GRUPPEN_NAME(new SimpleMergeField<>("gruppenname", STRING_CONVERTER)),

	// Spalten-Repeats
	REPEATMONTAG(new RepeatColMergeField<>("repeatMontag", STRING_CONVERTER)),
	REPEATDIENSTAG(new RepeatColMergeField<>("repeatDienstag", STRING_CONVERTER)),
	REPEATMITTWOCH(new RepeatColMergeField<>("repeatMittwoch", STRING_CONVERTER)),
	REPEATDONNERSTAG(new RepeatColMergeField<>("repeatDonnerstag", STRING_CONVERTER)),
	REPEATFREITAG(new RepeatColMergeField<>("repeatFreitag", STRING_CONVERTER)),
	REPEATSAMSTAG(new RepeatColMergeField<>("repeatSamstag", STRING_CONVERTER)),
	REPEATSONNTAG(new RepeatColMergeField<>("repeatSonntag", STRING_CONVERTER)),

	// Kind
	REPEAT_KIND(new RepeatRowMergeField("repeatKind")),

	NAME(new SimpleMergeField<>("name", STRING_CONVERTER)),
	VORNAME(new SimpleMergeField<>("vorname", STRING_CONVERTER)),
	GESCHLECHT(new SimpleMergeField<>("geschlecht", STRING_CONVERTER)),
	BETREUUNGSFAKTOR(new SimpleMergeField<>("betreuungsfaktor", BIGDECIMAL_CONVERTER)),
	GEBURTSTAG(new SimpleMergeField<>("geburtstag", DATE_CONVERTER)),
	GESCHWISTER(new SimpleMergeField<>("geschwister", BOOLEAN_X_CONVERTER)),

	MONTAG(new RepeatValMergeField<>("montag", STRING_CONVERTER)),
	DIENSTAG(new RepeatValMergeField<>("dienstag", STRING_CONVERTER)),
	MITTWOCH(new RepeatValMergeField<>("mittwoch", STRING_CONVERTER)),
	DONNERSTAG(new RepeatValMergeField<>("donnerstag", STRING_CONVERTER)),
	FREITAG(new RepeatValMergeField<>("freitag", STRING_CONVERTER)),
	SAMSTAG(new RepeatValMergeField<>("samstag", STRING_CONVERTER)),
	SONNTAG(new RepeatValMergeField<>("sonntag", STRING_CONVERTER)),

	BELEGUNG(new SimpleMergeField<>("belegung", BIGDECIMAL_CONVERTER)),
	KINDERGARTEN1(new SimpleMergeField<>("kindergarten1", STRING_CONVERTER)),
	KINDERGARTEN2(new SimpleMergeField<>("kindergarten2", STRING_CONVERTER)),
	FIRMA(new SimpleMergeField<>("firma", STRING_CONVERTER)),
	ANDEREGRUPPE(new SimpleMergeField<>("andereGruppe", STRING_CONVERTER)),
	STATUS(new SimpleMergeField<>("status", STRING_CONVERTER)),
	WUNSCHOFFEN(new SimpleMergeField<>("wunschOffen", BOOLEAN_X_CONVERTER)),
	GUELTIGBIS(new SimpleMergeField<>("gueltigBis", DATE_CONVERTER)),
	BEMERKUNG(new SimpleMergeField<>("bemerkung", STRING_CONVERTER)),

	// Gruppen-Summen
	BELEGUNGGESCHWISTER(new SimpleMergeField<>("belegungGeschwister", LONG_CONVERTER)),

	PLAETZEMONTAG(new RepeatValMergeField<>("plaetzeMontag", LONG_CONVERTER)),
	PLAETZEDIENSTAG(new RepeatValMergeField<>("plaetzeDienstag", LONG_CONVERTER)),
	PLAETZEMITTWOCH(new RepeatValMergeField<>("plaetzeMittwoch", LONG_CONVERTER)),
	PLAETZEDONNERSTAG(new RepeatValMergeField<>("plaetzeDonnerstag", LONG_CONVERTER)),
	PLAETZEFREITAG(new RepeatValMergeField<>("plaetzeFreitag", LONG_CONVERTER)),
	PLAETZESAMSTAG(new RepeatValMergeField<>("plaetzeSamstag", LONG_CONVERTER)),
	PLAETZESONNTAG(new RepeatValMergeField<>("plaetzeSonntag", LONG_CONVERTER)),

	BELEGUNGMONTAG(new RepeatValMergeField<>("belegungMontag", BIGDECIMAL_CONVERTER)),
	BELEGUNGDIENSTAG(new RepeatValMergeField<>("belegungDienstag", BIGDECIMAL_CONVERTER)),
	BELEGUNGMITTWOCH(new RepeatValMergeField<>("belegungMittwoch", BIGDECIMAL_CONVERTER)),
	BELEGUNGDONNERSTAG(new RepeatValMergeField<>("belegungDonnerstag", BIGDECIMAL_CONVERTER)),
	BELEGUNGFREITAG(new RepeatValMergeField<>("belegungFreitag", BIGDECIMAL_CONVERTER)),
	BELEGUNGSAMSTAG(new RepeatValMergeField<>("belegungSamstag", BIGDECIMAL_CONVERTER)),
	BELEGUNGSONNTAG(new RepeatValMergeField<>("belegungSonntag", BIGDECIMAL_CONVERTER)),

	MAXPLAETZEMONTAG(new RepeatValMergeField<>("maxPlaetzeMontag", LONG_CONVERTER)),
	MAXPLAETZEDIENSTAG(new RepeatValMergeField<>("maxPlaetzeDienstag", LONG_CONVERTER)),
	MAXPLAETZEMITTWOCH(new RepeatValMergeField<>("maxPlaetzeMittwoch", LONG_CONVERTER)),
	MAXPLAETZEDONNERSTAG(new RepeatValMergeField<>("maxPlaetzeDonnerstag", LONG_CONVERTER)),
	MAXPLAETZEFREITAG(new RepeatValMergeField<>("maxPlaetzeFreitag", LONG_CONVERTER)),
	MAXPLAETZESAMSTAG(new RepeatValMergeField<>("maxPlaetzeSamstag", LONG_CONVERTER)),
	MAXPLAETZESONNTAG(new RepeatValMergeField<>("maxPlaetzeSonntag", LONG_CONVERTER)),

	BELEGUNGSUMME(new SimpleMergeField<>("belegungSumme", BIGDECIMAL_CONVERTER)),
	PLAETZESUMME(new SimpleMergeField<>("plaetzeSumme", BIGDECIMAL_CONVERTER)),
	MAXPLAETZESUMME(new SimpleMergeField<>("maxPlaetzeSumme", BIGDECIMAL_CONVERTER)),
	BELEGUNGKINDERGARTEN1(new SimpleMergeField<>("belegungKindergarten1", LONG_CONVERTER)),
	BELEGUNGKINDERGARTEN2(new SimpleMergeField<>("belegungKindergarten2", LONG_CONVERTER)),

	WUNSCHOFFENSUMME(new SimpleMergeField<>("wunschOffenSumme", LONG_CONVERTER)),

	PLAETZEBELEGUNG(new SimpleMergeField<>("plaetzeBelegung", BIGDECIMAL_CONVERTER)),
	MAXPLAETZEBELEGUNG(new SimpleMergeField<>("maxPlaetzeBelegung", BIGDECIMAL_CONVERTER)),

	// Kita-Summen
	BELEGUNGKITAGESCHWISTER(new SimpleMergeField<>("belegungKitaGeschwister", LONG_CONVERTER)),

	BELEGUNGKITAMONTAG(new RepeatValMergeField<>("belegungKitaMontag", BIGDECIMAL_CONVERTER)),
	BELEGUNGKITADIENSTAG(new RepeatValMergeField<>("belegungKitaDienstag", BIGDECIMAL_CONVERTER)),
	BELEGUNGKITAMITTWOCH(new RepeatValMergeField<>("belegungKitaMittwoch", BIGDECIMAL_CONVERTER)),
	BELEGUNGKITADONNERSTAG(new RepeatValMergeField<>("belegungKitaDonnerstag", BIGDECIMAL_CONVERTER)),
	BELEGUNGKITAFREITAG(new RepeatValMergeField<>("belegungKitaFreitag", BIGDECIMAL_CONVERTER)),
	BELEGUNGKITASAMSTAG(new RepeatValMergeField<>("belegungKitaSamstag", BIGDECIMAL_CONVERTER)),
	BELEGUNGKITASONNTAG(new RepeatValMergeField<>("belegungKitaSonntag", BIGDECIMAL_CONVERTER)),

	PLAETZEKITAMONTAG(new RepeatValMergeField<>("plaetzeKitaMontag", LONG_CONVERTER)),
	PLAETZEKITADIENSTAG(new RepeatValMergeField<>("plaetzeKitaDienstag", LONG_CONVERTER)),
	PLAETZEKITAMITTWOCH(new RepeatValMergeField<>("plaetzeKitaMittwoch", LONG_CONVERTER)),
	PLAETZEKITADONNERSTAG(new RepeatValMergeField<>("plaetzeKitaDonnerstag", LONG_CONVERTER)),
	PLAETZEKITAFREITAG(new RepeatValMergeField<>("plaetzeKitaFreitag", LONG_CONVERTER)),
	PLAETZEKITASAMSTAG(new RepeatValMergeField<>("plaetzeKitaSamstag", LONG_CONVERTER)),
	PLAETZEKITASONNTAG(new RepeatValMergeField<>("plaetzeKitaSonntag", LONG_CONVERTER)),

	MAXPLAETZEKITAMONTAG(new RepeatValMergeField<>("maxPlaetzeKitaMontag", LONG_CONVERTER)),
	MAXPLAETZEKITADIENSTAG(new RepeatValMergeField<>("maxPlaetzeKitaDienstag", LONG_CONVERTER)),
	MAXPLAETZEKITAMITTWOCH(new RepeatValMergeField<>("maxPlaetzeKitaMittwoch", LONG_CONVERTER)),
	MAXPLAETZEKITADONNERSTAG(new RepeatValMergeField<>("maxPlaetzeKitaDonnerstag", LONG_CONVERTER)),
	MAXPLAETZEKITAFREITAG(new RepeatValMergeField<>("maxPlaetzeKitaFreitag", LONG_CONVERTER)),
	MAXPLAETZEKITASAMSTAG(new RepeatValMergeField<>("maxPlaetzeKitaSamstag", LONG_CONVERTER)),
	MAXPLAETZEKITASONNTAG(new RepeatValMergeField<>("maxPlaetzeKitaSonntag", LONG_CONVERTER)),

	MAXBEWILLIGTEPLAETZEKITAMONTAG(new RepeatValMergeField<>("maxBewilligtePlaetzeKitaMontag", BIGDECIMAL_CONVERTER)),
	MAXBEWILLIGTEPLAETZEKITADIENSTAG(new RepeatValMergeField<>("maxBewilligtePlaetzeKitaDienstag",
		BIGDECIMAL_CONVERTER)),
	MAXBEWILLIGTEPLAETZEKITAMITTWOCH(new RepeatValMergeField<>("maxBewilligtePlaetzeKitaMittwoch",
		BIGDECIMAL_CONVERTER)),
	MAXBEWILLIGTEPLAETZEKITADONNERSTAG(new RepeatValMergeField<>("maxBewilligtePlaetzeKitaDonnerstag",
		BIGDECIMAL_CONVERTER)),
	MAXBEWILLIGTEPLAETZEKITAFREITAG(new RepeatValMergeField<>("maxBewilligtePlaetzeKitaFreitag",
		BIGDECIMAL_CONVERTER)),
	MAXBEWILLIGTEPLAETZEKITASAMSTAG(new RepeatValMergeField<>("maxBewilligtePlaetzeKitaSamstag",
		BIGDECIMAL_CONVERTER)),
	MAXBEWILLIGTEPLAETZEKITASONNTAG(new RepeatValMergeField<>("maxBewilligtePlaetzeKitaSonntag",
		BIGDECIMAL_CONVERTER)),

	belegungKitaBelegung(new SimpleMergeField<>("belegungKitaBelegung", BIGDECIMAL_CONVERTER)),
	belegungKitaPlaetze(new SimpleMergeField<>("belegungKitaPlaetze", BIGDECIMAL_CONVERTER)),
	belegungKitaMaxPlaetze(new SimpleMergeField<>("belegungKitaMaxPlaetze", BIGDECIMAL_CONVERTER)),
	belegungKitaKindergarten1(new SimpleMergeField<>("belegungKitaKindergarten1", LONG_CONVERTER)),
	belegungKitaKindergarten2(new SimpleMergeField<>("belegungKitaKindergarten2", LONG_CONVERTER)),
	wunschOffenKitaSumme(new SimpleMergeField<>("wunschOffenKitaSumme", LONG_CONVERTER)),
	plaetzeKitaBelegung(new SimpleMergeField<>("plaetzeKitaBelegung", LONG_CONVERTER)),
	maxPlaetzeKitaBelegung(new SimpleMergeField<>("maxPlaetzeKitaBelegung", LONG_CONVERTER)),
	bewilligtePlaetzeKitaBelegung(new SimpleMergeField<>("bewilligtePlaetzeKitaBelegung", BIGDECIMAL_CONVERTER));

	@Nonnull
	private final MergeField<?> mergeField;

	<V> MergeFieldBelegungsplan(@Nonnull MergeField<V> mergeField) {
		this.mergeField = mergeField;
	}

	@Override
	@Nonnull
	public <V> MergeField<V> getMergeField() {
		//noinspection unchecked
		return (MergeField<V>) mergeField;
	}

}

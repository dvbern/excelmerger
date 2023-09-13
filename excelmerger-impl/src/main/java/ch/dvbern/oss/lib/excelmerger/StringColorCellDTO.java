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

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

import org.apache.poi.xssf.usermodel.XSSFColor;

public class StringColorCellDTO {

	@Nonnull
	private String value;

	@Nullable
	private XSSFColor color;

	@Nullable
	private XSSFColor fontColor;

	public StringColorCellDTO(@Nonnull String value, @Nullable XSSFColor color) {
		this.value = value;
		this.color = XSSFColor.toXSSFColor(color);
	}

	public StringColorCellDTO(@Nonnull String value, @Nullable XSSFColor color, @Nullable XSSFColor fontColor) {
		this.value = value;
		this.color = XSSFColor.toXSSFColor(color);
		this.fontColor = XSSFColor.toXSSFColor(fontColor);
	}

	@Nonnull
	public String getValue() {
		return value;
	}

	public void setValue(@Nonnull String value) {
		this.value = value;
	}

	@Nullable
	public XSSFColor getColor() {
		return color;
	}

	public void setColor(@Nullable XSSFColor color) {
		this.color = color;
	}

	public XSSFColor getFontColor() {
		return fontColor;
	}

	public void setFontColor(XSSFColor fontColor) {
		this.fontColor = fontColor;
	}
}

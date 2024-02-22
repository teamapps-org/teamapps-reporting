/*-
 * ========================LICENSE_START=================================
 * TeamApps Reporting
 * ---
 * Copyright (C) 2020 - 2024 TeamApps.org
 * ---
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * =========================LICENSE_END==================================
 */
package org.teamapps.reporting.style.run;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import java.math.BigInteger;

public class StyledText implements StyledElement {

	private String text;
	private Boolean bold;
	private Boolean italic;
	private Boolean underline;
	private Boolean caps;
	private Integer size;
	private String font;
	private String color;
	
	

	public String getText() {
		return text;
	}

	public StyledText setText(String text) {
		this.text = text;
		return this;
	}

	public Boolean getBold() {
		return bold;
	}

	public StyledText setBold(Boolean bold) {
		this.bold = bold;
		return this;
	}

	public Boolean getItalic() {
		return italic;
	}

	public StyledText setItalic(Boolean italic) {
		this.italic = italic;
		return this;
	}

	public Boolean getUnderline() {
		return underline;
	}

	public StyledText setUnderline(Boolean underline) {
		this.underline = underline;
		return this;
	}

	public Boolean getCaps() {
		return caps;
	}

	public StyledText setCaps(Boolean caps) {
		this.caps = caps;
		return this;
	}

	public Integer getSize() {
		return size;
	}

	public StyledText setSize(Integer size) {
		this.size = size;
		return this;
	}

	public String getFont() {
		return font;
	}

	public StyledText setFont(String font) {
		this.font = font;
		return this;
	}

	public String getColor() {
		return color;
	}

	public StyledText setColor(String color) {
		this.color = color;
		return this;
	}

	@Override
	public R getRun() {
		ObjectFactory factory = Context.getWmlObjectFactory();
		Text text = factory.createText();
		text.setValue(this.text);
		R run = factory.createR();
		run.getContent().add(text);
		RPr runProperties = factory.createRPr();
		BooleanDefaultTrue defaultTrue = new BooleanDefaultTrue();
		if (bold != null && bold) {
			runProperties.setB(defaultTrue);
		}
		if (italic != null && italic) {
			runProperties.setI(defaultTrue);
		}
		if (caps != null && caps) {
			runProperties.setCaps(defaultTrue);
		}
		if (size != null) {
			HpsMeasure size = factory.createHpsMeasure();
			size.setVal(BigInteger.valueOf(this.size));
			runProperties.setSz(size);
		}
		if (underline != null && underline) {
			U underline = factory.createU();
			underline.setVal(UnderlineEnumeration.SINGLE);
			runProperties.setU(underline);
		}

		if (color != null) {
			Color color = factory.createColor();
			color.setVal(this.color);
			runProperties.setColor(color);
		}

		if (font != null) {
			RFonts rFonts = factory.createRFonts();
			rFonts.setAscii(font);
			runProperties.setRFonts(rFonts);
		}

		run.setRPr(runProperties);
		return run;
	}


}

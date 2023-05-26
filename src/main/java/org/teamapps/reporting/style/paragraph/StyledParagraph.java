package org.teamapps.reporting.style.paragraph;

import org.docx4j.jaxb.Context;
import org.docx4j.model.PropertyResolver;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;
import org.teamapps.reporting.builder.ReportBuilder;
import org.teamapps.reporting.convert.DocumentFormat;
import org.teamapps.reporting.style.run.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class StyledParagraph {

	public final List<StyledElement> elements = new ArrayList<>();

	private WordprocessingMLPackage wordPackage;
	private ParagraphStyle paragraphStyle;

	public void addElement(StyledElement element) {
		elements.add(element);
	}

	public void setStyle(ParagraphStyle paragraphStyle, WordprocessingMLPackage wordPackage) {
		this.wordPackage = wordPackage;
		this.paragraphStyle = paragraphStyle;
	}

	public P getParagraph() {
		ObjectFactory factory = Context.getWmlObjectFactory();
		P paragraph = factory.createP();
		setStyle(paragraph);
		elements.forEach(e -> paragraph.getContent().add(e.getRun()));
		return paragraph;
	}

	public void replaceContent(P paragraph) {
		paragraph.getContent().clear();
		setStyle(paragraph);
		elements.forEach(e -> paragraph.getContent().add(e.getRun()));
	}

	public void setStyle(P paragraph) {
		if (wordPackage == null || paragraphStyle == null) {
			return;
		}
		PropertyResolver propertyResolver = wordPackage.getMainDocumentPart().getPropertyResolver();
		String styleId = paragraphStyle.getStyleId();
		if (propertyResolver.activateStyle(styleId)) {
			ObjectFactory factory = Context.getWmlObjectFactory();
			PPr pPr = factory.createPPr();
			paragraph.setPPr(pPr);
			PPrBase.PStyle pStyle = factory.createPPrBasePStyle();
			pPr.setPStyle(pStyle);
			pStyle.setVal(styleId);
		}
	}

}

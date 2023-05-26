package org.teamapps.reporting.style.run;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.R;

public interface StyledElement {

	R getRun();

	default ObjectFactory getFactory() {
		return Context.getWmlObjectFactory();
	}

	default R createRun() {
		return getFactory().createR();
	}
}

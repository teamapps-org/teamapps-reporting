package org.teamapps.reporting.style.run;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.Br;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.R;

public class LineBreak implements StyledElement {

	@Override
	public R getRun() {
		ObjectFactory factory = Context.getWmlObjectFactory();
		Br br = factory.createBr();
		R run = factory.createR();
		run.getContent().add(br);
		return run;
	}
}

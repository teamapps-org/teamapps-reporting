package org.teamapps.reporting.style.run;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

public class TextSpace implements StyledElement {

	@Override
	public R getRun() {
		ObjectFactory factory = Context.getWmlObjectFactory();
		R run = factory.createR();
		Text space = new Text();
		space.setSpace("preserve");
		space.setValue(" ");
		run.getContent().add(space);
		return run;
	}
}

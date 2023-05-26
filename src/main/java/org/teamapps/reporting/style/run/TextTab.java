package org.teamapps.reporting.style.run;

import org.docx4j.wml.R;

public class TextTab implements StyledElement {

	@Override
	public R getRun() {
		R run = createRun();
		R.Tab tab = getFactory().createRTab();
		run.getContent().add(tab);
		return run;
	}


}

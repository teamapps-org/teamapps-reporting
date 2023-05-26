package org.teamapps.reporting.style.run;

import org.docx4j.UnitsOfMeasurement;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.R;
import org.docx4j.wml.*;

import java.io.File;
import java.nio.file.Files;

public class ImageElement implements StyledElement{

	private final WordprocessingMLPackage wordPackage;
	private final File file;
	private Integer width;
	private Integer height;

	public ImageElement(WordprocessingMLPackage wordPackage, File file) {
		this.wordPackage = wordPackage;
		this.file = file;
	}

	public ImageElement setSizeInMm(Integer width, Integer height) {
		this.width = width;
		this.height = height;
		return this;
	}

	@Override
	public R getRun() {
		try {
			ObjectFactory factory = new ObjectFactory();
			byte[] fileContent = Files.readAllBytes(file.toPath());
			BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordPackage, fileContent);
			Inline inline;
			if (width != null && height != null) {
				long widthEmu = UnitsOfMeasurement.twipToEMU(UnitsOfMeasurement.mmToTwip(width));
				long heightEmu = UnitsOfMeasurement.twipToEMU(UnitsOfMeasurement.mmToTwip(height));
				inline = imagePart.createImageInline("", "", 1, 2, widthEmu, heightEmu, false);
			} else {
				inline = imagePart.createImageInline("", "", 1, 2, false);
			}

			R run = factory.createR();
			Drawing drawing = factory.createDrawing();
			run.getContent().add(drawing);
			drawing.getAnchorOrInline().add(inline);
			return run;
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

}

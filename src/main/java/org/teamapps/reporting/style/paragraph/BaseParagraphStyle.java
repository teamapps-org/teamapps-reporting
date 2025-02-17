/*-
 * ========================LICENSE_START=================================
 * TeamApps Reporting
 * ---
 * Copyright (C) 2020 - 2025 TeamApps.org
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
package org.teamapps.reporting.style.paragraph;

public enum BaseParagraphStyle implements ParagraphStyle{

	BALLOON_TEXT("BalloonText"),
	BIBLIOGRAPHY("Bibliography"),
	BLOCK_TEXT("BlockText"),
	BODY_TEXT("BodyText"),
	BODY_TEXT2("BodyText2"),
	BODY_TEXT3("BodyText3"),
	BODY_TEXT_FIRST_INDENT("BodyTextFirstIndent"),
	BODY_TEXT_FIRST_INDENT2("BodyTextFirstIndent2"),
	BODY_TEXT_INDENT("BodyTextIndent"),
	BODY_TEXT_INDENT2("BodyTextIndent2"),
	BODY_TEXT_INDENT3("BodyTextIndent3"),
	CAPTION("Caption"),
	CLOSING("Closing"),
	COMMENT_SUBJECT("CommentSubject"),
	COMMENT_TEXT("CommentText"),
	DATE("Date"),
	DOCUMENT_MAP("DocumentMap"),
	E_MAIL_SIGNATURE("E-mailSignature"),
	END_NOTE_TEXT("EndnoteText"),
	ENVELOPE_ADDRESS("EnvelopeAddress"),
	ENVELOPE_RETURN("EnvelopeReturn"),
	FOOTER("Footer"),
	FOOTNOTE_TEXT("FootnoteText"),
	HTML_ADDRESS("HTMLAddress"),
	HTML_PREFORMATTED("HTMLPreformatted"),
	HEADER("Header"),
	HEADING1("Heading1"),
	HEADING2("Heading2"),
	HEADING3("Heading3"),
	HEADING4("Heading4"),
	HEADING5("Heading5"),
	HEADING6("Heading6"),
	HEADING7("Heading7"),
	HEADING8("Heading8"),
	HEADING9("Heading9"),
	INDEX1("Index1"),
	INDEX2("Index2"),
	INDEX3("Index3"),
	INDEX4("Index4"),
	INDEX5("Index5"),
	INDEX6("Index6"),
	INDEX7("Index7"),
	INDEX8("Index8"),
	INDEX9("Index9"),
	INDEX_HEADING("IndexHeading"),
	INTENSE_QUOTE("IntenseQuote"),
	LIST("List"),
	LIST2("List2"),
	LIST3("List3"),
	LIST4("List4"),
	LIST5("List5"),
	LIST_BULLET("ListBullet"),
	LIST_BULLET2("ListBullet2"),
	LIST_BULLET3("ListBullet3"),
	LIST_BULLET4("ListBullet4"),
	LIST_BULLET5("ListBullet5"),
	LIST_CONTINUE("ListContinue"),
	LIST_CONTINUE2("ListContinue2"),
	LIST_CONTINUE3("ListContinue3"),
	LIST_CONTINUE4("ListContinue4"),
	LIST_CONTINUE5("ListContinue5"),
	LIST_NUMBER("ListNumber"),
	LIST_NUMBER2("ListNumber2"),
	LIST_NUMBER3("ListNumber3"),
	LIST_NUMBER4("ListNumber4"),
	LIST_NUMBER5("ListNumber5"),
	LIST_PARAGRAPH("ListParagraph"),
	MACRO_TEXT("MacroText"),
	MESSAGE_HEADER("MessageHeader"),
	NO_SPACING("NoSpacing"),
	NORMAL("Normal"),
	NORMAL_INDENT("NormalIndent"),
	NORMAL_WEB("NormalWeb"),
	NOTE_HEADING("NoteHeading"),
	PLAIN_TEXT("PlainText"),
	QUOTE("Quote"),
	SALUTATION("Salutation"),
	SIGNATURE("Signature"),
	SUBTITLE("Subtitle"),
	T_OA_HEADING("TOAHeading"),
	T_OC1("TOC1"),
	T_OC2("TOC2"),
	T_OC3("TOC3"),
	T_OC4("TOC4"),
	T_OC5("TOC5"),
	T_OC6("TOC6"),
	T_OC7("TOC7"),
	T_OC8("TOC8"),
	T_OC9("TOC9"),
	T_OC_HEADING("TOCHeading"),
	TABLE_OF_AUTHORITIES("TableofAuthorities"),
	TABLE_OF_FIGURES("TableofFigures"),
	TITLE("Title"),
	
	;
	private final String name;


	BaseParagraphStyle(String name) {
		this.name = name;
	}

	@Override
	public String getStyleId() {
		return name;
	}
}

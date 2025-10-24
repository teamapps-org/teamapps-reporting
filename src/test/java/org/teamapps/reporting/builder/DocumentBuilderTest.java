package org.teamapps.reporting.builder;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.junit.BeforeClass;
import org.junit.Test;
import org.teamapps.reporting.convert.DocumentFormat;

import javax.print.Doc;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

import static org.junit.Assert.*;

public class DocumentBuilderTest {
	static Locale printLocale = Locale.GERMANY;
	static DecimalFormat numberFormat;
	static Currency usd = Currency.getInstance(Locale.US);
	static Currency chf = Currency.getInstance(Locale.forLanguageTag("de-CH"));
	static Currency eur = Currency.getInstance(Locale.GERMANY);
	static Map<Currency, Double> rateToEur = new HashMap<>();

	@BeforeClass
	public static void init() {
		rateToEur.put(usd, 1.039);
		rateToEur.put(chf, 0.944);
		rateToEur.put(eur, 1.0);

		DecimalFormatSymbols decimalSymbols = DecimalFormatSymbols.getInstance(printLocale);
		Currency currency = decimalSymbols.getCurrency();
		numberFormat = new DecimalFormat();
		numberFormat.setDecimalFormatSymbols(decimalSymbols);
		numberFormat.setMaximumFractionDigits(currency.getDefaultFractionDigits());
		numberFormat.setMinimumFractionDigits(currency.getDefaultFractionDigits());
	}

	private double addRow(TableBuilder tableBuilder, String name, String customer, int number, double amount, Currency currency) {
		RowBuilder row = tableBuilder.addRow();
		row.setColumnValue("<name>", name)
				.setColumnValue("<customer>", customer)
				.setColumnValue("<number>", String.valueOf(number))
				.setColumnValue("<currency>", currency.getCurrencyCode())
				.setColumnValue("<currencySymbol>", currency.getSymbol())
				.setColumnValue("<amount>", numberFormat.format(amount))
				.setColumnValue("<amountEUR>", numberFormat.format(amount/rateToEur.getOrDefault(currency, 1.0)))
		;
		return amount;
	}

	private TableBuilder fillTable(ReportBuilder reportBuilder, String tableKey, String tableName) {
		TableBuilder tableBuilder = reportBuilder.createTableBuilder(tableKey);

		tableBuilder.addRow().setColumnValue(tableKey, tableName);
		tableBuilder.addRow()
				.setColumnValue("<NameHeader>", "Produkt")
				.setColumnValue("<CustomerHeader>", "Kunde")
				.setColumnValue("<NumberHeader>", "Kostenart")
				.setColumnValue("<CurrencyHeader>", "Währung")
				.setColumnValue("<AmountHeader>", "Betrag")
				.setColumnValue("<AmountEurHeader>", "Betrag EUR")
		;

		double sum = 0.0;
		double totalEur = 0.0;

		sum += addRow(tableBuilder, "chocolate", "Henry", 15, 13.2, usd);
		sum += addRow(tableBuilder, "flowers", "Susan", 7, 25.99, usd);
		sum += addRow(tableBuilder, "ice cream", "Peggy", 23, 7.5, usd);

		RowBuilder partSumUS = tableBuilder.addRow();
		partSumUS.setColumnValue("<partSum>", "Summe")
				.setColumnValue("<currency>", usd.getCurrencyCode())
				.setColumnValue("<currencySymbol>", usd.getSymbol())
				.setColumnValue("<amount>", numberFormat.format(sum))
				.setColumnValue("<amountEUR>", numberFormat.format(sum/rateToEur.getOrDefault(usd, 1.0)))
		;
		totalEur += sum/rateToEur.getOrDefault(usd, 1.0);
		tableBuilder.addRow().setColumnValue("<empty>", "");

		sum = 0.0;
		sum += addRow(tableBuilder, "chocolate", "for Henry", 15, 13.2, chf);
		sum += addRow(tableBuilder, "flowers", "for Susan", 7, 25.99, chf);
		sum += addRow(tableBuilder, "ice cream", "for Peggy", 23, 7.5, chf);

		RowBuilder partSum = tableBuilder.addRow();
		partSum.setColumnValue("<partSum>", "Summe")
				.setColumnValue("<currency>", chf.getCurrencyCode())
				.setColumnValue("<currencySymbol>", chf.getSymbol())
				.setColumnValue("<amount>", numberFormat.format(sum))
				.setColumnValue("<amountEUR>", numberFormat.format(sum/rateToEur.getOrDefault(chf, 1.0)))
		;

		totalEur += sum/rateToEur.getOrDefault(chf, 1.0);
		tableBuilder.addRow().setColumnValue("<empty>", "");

		RowBuilder totalSum = tableBuilder.addRow();
		totalSum.setColumnValue("<totalSum>", "Total")
				.setColumnValue("<amountEUR>", numberFormat.format(totalEur))
		;
		tableBuilder.removeUnusedTemplateRow("<ExampleTable1>");
		tableBuilder.removeUnusedTemplateRow("<empty>");
		tableBuilder.removeUnusedTemplateRow("<NameHeader>");
		tableBuilder.removeUnusedTemplateRow("<name>");
		tableBuilder.removeUnusedTemplateRow("<partSum>");
		tableBuilder.removeUnusedTemplateRow("<totalSum>");
		return tableBuilder;
	}

	public String getContent(Object element) {
		DocumentBuilder documentBuilder = new DocumentBuilder();
		List<P> paragraphs = documentBuilder.getAllElements(element, new P());
		StringBuilder sb = new StringBuilder();
		for (P paragraph : paragraphs) {
			List<Text> texts = documentBuilder.getAllElements(paragraph, new Text());
			texts.forEach(text -> sb.append(text.getValue()));
			sb.append(" | ");
		}
		if (sb.isEmpty()) {
			return "";
		}
		return sb.substring(0,sb.length()-3);
	}

	private String checkTable(MainDocumentPart document, TableBuilder tableBuilder, String tableKey) {
		DocumentBuilder documentBuilder = new DocumentBuilder();
		Tbl table1 = documentBuilder.findTable(document, List.of(tableKey));

		if (table1==null) return "table " + tableKey + " not found";
		List<Tr> rows = documentBuilder.getAllElements(table1, new Tr());
		Set<String> variables = tableBuilder.getVariables();
		for (Tr row : rows) {
			for (String variable : variables) {
				if (documentBuilder.getParagraphWithText(row, variable) != null) {
					return "variable " + variable + " is not replaced in table " + tableKey + " and row: " + getContent(row);
				}
			}
		}
		return null;
	}

	 @Test
	 public void tableTestVersion1() throws Exception {
		 String templatePath = "templates/exampleReport.docx";

		 InputStream reportTemplateInput = DocumentBuilderTest.class.getResourceAsStream(templatePath);
		 ReportBuilder reportBuilder = ReportBuilder.create(DocumentFormat.DOCX, reportTemplateInput);

		 SimpleDateFormat sdf = (SimpleDateFormat)SimpleDateFormat.getDateInstance(DateFormat.SHORT, printLocale);
		 DateTimeFormatter simpleDateFormat = DateTimeFormatter.ofPattern(sdf.toPattern());
		 reportBuilder.addReplacement("<currentDate>", simpleDateFormat.format(LocalDate.now()));

		 TableBuilder table1Builder = fillTable(reportBuilder, "<ExampleTable1>", "Version 1");

		 reportBuilder.removeTable("<ExampleTable3>");

		 String tempDir = System.getProperty("java.io.tmpdir");
		 File resultDirectory = new File(tempDir, "testResults");
		 resultDirectory.mkdirs();
		 File resultingFile = new File(resultDirectory, "exampleReportResult1.docx");
		 if (resultingFile.exists()) resultingFile.delete();
		 reportBuilder.setOutputFile(resultingFile);
		 File wordReport = reportBuilder.build();
		 System.out.println("report " + wordReport.getPath() + " created!");

		 InputStream resultInput = new FileInputStream(wordReport);
		 MainDocumentPart result = DocumentTemplateLoader.getTemplate(resultInput).getMainDocumentPart();

		 assertNull(checkTable(result, table1Builder, "Version 1"));
		 DocumentBuilder documentBuilder = new DocumentBuilder();
		 assertNull(documentBuilder.findTable(result, List.of("<ExampleTable3>")));
	 }

	 @Test
	 public void tableTestVersion2() throws Exception {
		 String templatePath = "templates/exampleReport.docx";

		 InputStream reportTemplateInput = DocumentBuilderTest.class.getResourceAsStream(templatePath);
		 ReportBuilder reportBuilder = ReportBuilder.create(DocumentFormat.DOCX, reportTemplateInput);

		 SimpleDateFormat sdf = (SimpleDateFormat)SimpleDateFormat.getDateInstance(DateFormat.SHORT, printLocale);
		 DateTimeFormatter simpleDateFormat = DateTimeFormatter.ofPattern(sdf.toPattern());
		 reportBuilder.addReplacement("<currentDate>", simpleDateFormat.format(LocalDate.now()));

		 TableBuilder table2Builder = fillTable(reportBuilder, "<ExampleTable2>", "Version 2");

		 reportBuilder.removeTable("<ExampleTable3>");

		 String tempDir = System.getProperty("java.io.tmpdir");
		 File resultDirectory = new File(tempDir, "testResults");
		 resultDirectory.mkdirs();
		 File resultingFile = new File(resultDirectory, "exampleReportResult2.docx");
		 if (resultingFile.exists()) resultingFile.delete();
		 reportBuilder.setOutputFile(resultingFile);
		 File wordReport = reportBuilder.build();
		 System.out.println("report " + wordReport.getPath() + " created!");

		 InputStream resultInput = new FileInputStream(wordReport);
		 MainDocumentPart result = DocumentTemplateLoader.getTemplate(resultInput).getMainDocumentPart();

		 assertNull(checkTable(result, table2Builder, "Version 2"));
		 DocumentBuilder documentBuilder = new DocumentBuilder();
		 assertNull(documentBuilder.findTable(result, List.of("<ExampleTable3>")));
	 }
}
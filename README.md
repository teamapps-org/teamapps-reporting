
# TeamApps Reporting

TeamApps reporting utilities:

* Design beautiful report templates with Microsoft Word, LibreOffice, OpenOffice or any RTF-Editor
* Programmatically populate report templates with data of your application
* Render the reports as Microsoft Word Document, OpenDocument, RTF or Image
* Convert between file formats using this library (DOCX, PDF and PNG only) or via REST-Service (Docker image coming soon - for now use: https://github.com/EugenMayer/docker-image-jodconverter)

### Setting up the dependencies

Current TeamApps Reporting version: [![Maven Central](https://maven-badges.herokuapp.com/maven-central/org.teamapps/teamapps-reporting/badge.svg)](https://maven-badges.herokuapp.com/maven-central/org.teamapps/teamapps-reporting)
```xml
<dependency>
  <groupId>org.teamapps</groupId>
  <artifactId>teamapps-reporting</artifactId>
  <version>1.7</version>
</dependency>
```


## Inspecting a prepared report

Use `ReportBuilder.getDefinition()` to inspect the data supplied to the generator
without reading the template, generating a document or invoking a converter. This
lets application tests verify their prepared report data independently of DOCX rendering:

```java
ReportBuilder report = ReportBuilder.create(DocumentFormat.DOCX, templateInput);
report.addReplacement("<title>", "Annual report");
TableBuilder table = report.createTableBuilder("<items>");
table.addRow().setColumnValue("<name>", "Donation").setColumnValue("<amount>", "100.00");
table.removeUnusedTemplateRow("<name>");
report.removeTable("<unused>");
report.removeParagraph("<draft>");

ReportDefinition snapshot = report.getDefinition();
assertEquals("Annual report", snapshot.replacements().get("<title>"));
TableDefinition items = snapshot.tables().get(0);
assertEquals(List.of("<items>"), items.keys());
assertEquals(List.of(Map.of("<name>", "Donation", "<amount>", "100.00")), items.rows());
assertEquals(List.of(List.of("<name>")), items.templateRowsToRemove());
assertEquals(List.of(List.of("<unused>")), snapshot.tablesToRemove());
assertEquals(List.of("<draft>"), snapshot.paragraphsToRemove());
```

The assertions use JUnit's `assertEquals`; `List` and `Map` are from `java.util`.
`TableBuilder.getDefinition()` also returns an individual table snapshot.

`ReportDefinition` includes `inputFormat()` and `strictMode()`.
Each `TableDefinition` includes `copyTable()`.
Table order, input row order and removal-request order are preserved. Both records
support value equality, so tests can compare complete expected definitions.

All exposed collections, including individual row maps and nested removal keys,
are unmodifiable copies. Later changes to builders or caller-supplied lists and
arrays do not change an earlier snapshot. Null replacement values are preserved
for inspection; snapshots do not validate whether the supplied data can be rendered.

A snapshot describes the requested report, not the generated output: it cannot tell
whether a matching template table or row exists, or whether replacements succeed.
Template streams, output files and diagnostic listeners are not included. Snapshot
creation must not run concurrently with builder modifications. Snapshots contain
actual replacement values and should be handled with the same care as report data.

## License

The TeamApps Reporting is released under version 2.0 of the [Apache License](https://www.apache.org/licenses/LICENSE-2.0).

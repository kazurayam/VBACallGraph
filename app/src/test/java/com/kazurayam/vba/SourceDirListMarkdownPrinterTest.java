package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class SourceDirListMarkdownPrinterTest {

    private TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(SourceDirListMarkdownPrinterTest.class)
                    .subOutputDirectory(SourceDirListMarkdownPrinterTest.class)
                    .build();
    private Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");
    private Path classOutputDir;
    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
    }
    @Test
    public void test_printAllSourceDirs() throws IOException {
        SourceDirListMarkdownPrinter printer = new SourceDirListMarkdownPrinter();
        printer.add(new Workbook(
                MyWorkbook.Backbone.getId(),
                MyWorkbook.Backbone.resolveWorkbookUnder(baseDir),
                MyWorkbook.Backbone.resolveSourceDirUnder(baseDir)));
        printer.add(new Workbook(
                MyWorkbook.Member.getId(),
                MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                MyWorkbook.Member.resolveSourceDirUnder(baseDir)));
        printer.add(new Workbook(
                MyWorkbook.Cashbook.getId(),
                MyWorkbook.Cashbook.resolveWorkbookUnder(baseDir),
                MyWorkbook.Cashbook.resolveSourceDirUnder(baseDir)));
        printer.add(new Workbook(
                MyWorkbook.Settlement.getId(),
                MyWorkbook.Settlement.resolveWorkbookUnder(baseDir),
                MyWorkbook.Settlement.resolveSourceDirUnder(baseDir)));
        printer.add(new Workbook(
                MyWorkbook.FeePaymentCheck.getId(),
                MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir)));
        printer.add(new Workbook(
                MyWorkbook.PleasePayFeeLetter.getId(),
                MyWorkbook.PleasePayFeeLetter.resolveWorkbookUnder(baseDir),
                MyWorkbook.PleasePayFeeLetter.resolveSourceDirUnder(baseDir)));
        printer.add(new Workbook(
                MyWorkbook.WebCredentials.getId(),
                MyWorkbook.WebCredentials.resolveWorkbookUnder(baseDir),
                MyWorkbook.WebCredentials.resolveSourceDirUnder(baseDir)));
        //
        Path report = classOutputDir.resolve("MyVBASourceDirs.md");
        Writer writer = Files.newBufferedWriter(report);
        printer.printAllSourceDirs(writer);
        assertThat(report).exists();
    }

}

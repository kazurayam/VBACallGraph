package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
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
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.Backbone));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.Member));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.Cashbook));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.Settlement));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.FeePaymentCheck));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.PleasePayFeeLetter));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.WebCredentials));
        //
        Path report = classOutputDir.resolve("MyVBASourceDirs.md");
        Writer writer = Files.newBufferedWriter(report);
        printer.printAllSourceDirs(writer);
        assertThat(report).exists();
    }

}

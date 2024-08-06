package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class VBASourceMarkdownPrinterTest {

    private TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(VBASourceMarkdownPrinterTest.class)
                    .subOutputDirectory(VBASourceMarkdownPrinterTest.class)
                    .build();
    private Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");
    private Path classOutputDir;
    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
    }
    @Test
    public void test_printAllVBASourceDirs() throws IOException {
        VBASourceMarkdownPrinter printer = new VBASourceMarkdownPrinter();
        printer.add(new ResolvedMyWorkbook(baseDir, MyWorkbook.Backbone));
        printer.add(new ResolvedMyWorkbook(baseDir, MyWorkbook.Member));
        printer.add(new ResolvedMyWorkbook(baseDir, MyWorkbook.Cashbook));
        printer.add(new ResolvedMyWorkbook(baseDir, MyWorkbook.Settlement));
        printer.add(new ResolvedMyWorkbook(baseDir, MyWorkbook.FeePaymentCheck));
        printer.add(new ResolvedMyWorkbook(baseDir, MyWorkbook.PleasePayFeeLetter));
        printer.add(new ResolvedMyWorkbook(baseDir, MyWorkbook.WebCredentials));
        //
        Path report = classOutputDir.resolve("MyVBASourceDirs.md");
        Writer writer = Files.newBufferedWriter(report);
        printer.printAllVBASourceDirs(writer);
        assertThat(report).exists();
    }

}

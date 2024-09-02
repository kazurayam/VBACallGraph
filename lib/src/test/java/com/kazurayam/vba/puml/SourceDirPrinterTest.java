package com.kazurayam.vba.puml;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.example.MyWorkbook;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class SourceDirPrinterTest {

    private TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(SourceDirPrinterTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(SourceDirPrinterTest.class)
                    .build();

    private Path classOutputDir;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
    }
    @Test
    public void test_printAllSourceDirs() throws IOException {
        SourceDirPrinter printer = new SourceDirPrinter();
        printer.add(new ModelWorkbook(
                MyWorkbook.Backbone.resolveWorkbookUnder(),
                MyWorkbook.Backbone.resolveSourceDirUnder())
                .id(MyWorkbook.Backbone.getId()));
        printer.add(new ModelWorkbook(
                MyWorkbook.Member.resolveWorkbookUnder(),
                MyWorkbook.Member.resolveSourceDirUnder())
                .id(MyWorkbook.Member.getId()));
        printer.add(new ModelWorkbook(
                MyWorkbook.Cashbook.resolveWorkbookUnder(),
                MyWorkbook.Cashbook.resolveSourceDirUnder())
                .id(MyWorkbook.Cashbook.getId()));
        printer.add(new ModelWorkbook(
                MyWorkbook.Settlement.resolveWorkbookUnder(),
                MyWorkbook.Settlement.resolveSourceDirUnder())
                .id(MyWorkbook.Settlement.getId()));
        printer.add(new ModelWorkbook(
                MyWorkbook.FeePaymentControl.resolveWorkbookUnder(),
                MyWorkbook.FeePaymentControl.resolveSourceDirUnder())
                .id(MyWorkbook.FeePaymentControl.getId()));
        printer.add(new ModelWorkbook(
                MyWorkbook.WebCredentials.resolveWorkbookUnder(),
                MyWorkbook.WebCredentials.resolveSourceDirUnder())
                .id(MyWorkbook.WebCredentials.getId()));
        //
        Path report = classOutputDir.resolve("MyVBASourceDirs.md");
        Writer writer = Files.newBufferedWriter(report);
        printer.printAllSourceDirs(writer);
        assertThat(report).exists();
    }

}

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

    private final Path baseDir = too.getProjectDirectory().resolve("src/test/fixture/hub");

    private Path classOutputDir;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
    }
    @Test
    public void test_printAllSourceDirs() throws IOException {
        SourceDirPrinter printer = new SourceDirPrinter();
        printer.add(new ModelWorkbook(
                MyWorkbook.Backbone.resolveWorkbookUnder(baseDir),
                MyWorkbook.Backbone.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Backbone.getId()));
        printer.add(new ModelWorkbook(
                MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                MyWorkbook.Member.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Member.getId()));
        printer.add(new ModelWorkbook(
                MyWorkbook.Cashbook.resolveWorkbookUnder(baseDir),
                MyWorkbook.Cashbook.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Cashbook.getId()));
        printer.add(new ModelWorkbook(
                MyWorkbook.Settlement.resolveWorkbookUnder(baseDir),
                MyWorkbook.Settlement.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Settlement.getId()));
        printer.add(new ModelWorkbook(
                MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.FeePaymentCheck.getId()));
        printer.add(new ModelWorkbook(
                MyWorkbook.PleasePayFeeLetter.resolveWorkbookUnder(baseDir),
                MyWorkbook.PleasePayFeeLetter.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.PleasePayFeeLetter.getId()));
        printer.add(new ModelWorkbook(
                MyWorkbook.WebCredentials.resolveWorkbookUnder(baseDir),
                MyWorkbook.WebCredentials.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.WebCredentials.getId()));
        //
        Path report = classOutputDir.resolve("MyVBASourceDirs.md");
        Writer writer = Files.newBufferedWriter(report);
        printer.printAllSourceDirs(writer);
        assertThat(report).exists();
    }

}

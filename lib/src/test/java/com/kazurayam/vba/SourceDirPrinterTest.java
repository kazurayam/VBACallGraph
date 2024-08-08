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
        printer.add(new SensibleWorkbook(
                MyWorkbook.Backbone.getId(),
                MyWorkbook.Backbone.resolveWorkbookUnder(baseDir),
                MyWorkbook.Backbone.resolveSourceDirUnder(baseDir)));
        printer.add(new SensibleWorkbook(
                MyWorkbook.Member.getId(),
                MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                MyWorkbook.Member.resolveSourceDirUnder(baseDir)));
        printer.add(new SensibleWorkbook(
                MyWorkbook.Cashbook.getId(),
                MyWorkbook.Cashbook.resolveWorkbookUnder(baseDir),
                MyWorkbook.Cashbook.resolveSourceDirUnder(baseDir)));
        printer.add(new SensibleWorkbook(
                MyWorkbook.Settlement.getId(),
                MyWorkbook.Settlement.resolveWorkbookUnder(baseDir),
                MyWorkbook.Settlement.resolveSourceDirUnder(baseDir)));
        printer.add(new SensibleWorkbook(
                MyWorkbook.FeePaymentCheck.getId(),
                MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir)));
        printer.add(new SensibleWorkbook(
                MyWorkbook.PleasePayFeeLetter.getId(),
                MyWorkbook.PleasePayFeeLetter.resolveWorkbookUnder(baseDir),
                MyWorkbook.PleasePayFeeLetter.resolveSourceDirUnder(baseDir)));
        printer.add(new SensibleWorkbook(
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

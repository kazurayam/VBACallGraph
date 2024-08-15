package com.kazurayam.vba.puml;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.example.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class FindUsageAppTest {

    private static final Logger logger =
            LoggerFactory.getLogger(FindUsageAppTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(FindUsageAppTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(FindUsageAppTest.class)
                    .build();

    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");
    private FindUsageApp app;
    private Path classOutputDir;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        app = new FindUsageApp();
        app.add(new SensibleWorkbook(
                MyWorkbook.FeePaymentCheck.getId(),
                MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir)
        ));

        app.add(new SensibleWorkbook(
                MyWorkbook.Cashbook.getId(),
                MyWorkbook.Cashbook.resolveWorkbookUnder(baseDir),
                MyWorkbook.Cashbook.resolveSourceDirUnder(baseDir)
        ));

        app.add(new SensibleWorkbook(
                MyWorkbook.Member.getId(),
                MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                MyWorkbook.Member.resolveSourceDirUnder(baseDir)
        ));

        app.add(new SensibleWorkbook(
                MyWorkbook.Backbone.getId(),
                MyWorkbook.Backbone.resolveWorkbookUnder(baseDir),
                MyWorkbook.Backbone.resolveSourceDirUnder(baseDir)
        ));
        app.setOptions(Options.KAZURAYAM);
    }

    @Test
    public void test_toString() throws IOException {
        String str = app.toString();
        assertThat(str).isNotNull();
        logger.info("[test_toString] " + str);
        Path file = classOutputDir.resolve("test_toString.json");
        Files.writeString(file, str);
    }

    @Test
    public void test_writeDiagram_Options_KAZURAYAM() throws IOException {
        Path file = classOutputDir.resolve("test_writeDiagram_Options_KAZURAYAM.pu");
        app.writeDiagram(file);
        assertThat(file).exists();
        assertThat(file.toFile().length()).isGreaterThan(0);
    }

    @Test
    public void test_writeDiagram_Options_RELAXED() throws IOException {
        Path file = classOutputDir.resolve("test_writeDiagram_Options_RELAXED.pu");
        app.setOptions(Options.RELAXED);
        app.writeDiagram(file);
        assertThat(file).exists();
        assertThat(file.toFile().length()).isGreaterThan(0);
    }

    @Test
    public void test_writeDiagram_Options_DEFAULT() throws IOException {
        Path file = classOutputDir.resolve("test_writeDiagram_Options_DEFAULT.pu");
        app.setOptions(Options.DEFAULT);
        app.writeDiagram(file);
        assertThat(file).exists();
        assertThat(file.toFile().length()).isGreaterThan(0);
    }
}

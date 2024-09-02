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

public class CallGraphAppTest {

    private static final Logger logger =
            LoggerFactory.getLogger(CallGraphAppTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(CallGraphAppTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(CallGraphAppTest.class)
                    .build();

    private CallGraphApp app;
    private Path classOutputDir;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        app = new CallGraphApp();
        app.add(new ModelWorkbook(
                MyWorkbook.FeePaymentControl.resolveWorkbookUnder(),
                MyWorkbook.FeePaymentControl.resolveSourceDirUnder())
                .id(MyWorkbook.FeePaymentControl.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.Cashbook.resolveWorkbookUnder(),
                MyWorkbook.Cashbook.resolveSourceDirUnder())
                .id(MyWorkbook.Cashbook.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.Member.resolveWorkbookUnder(),
                MyWorkbook.Member.resolveSourceDirUnder())
                .id(MyWorkbook.Member.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.Backbone.resolveWorkbookUnder(),
                MyWorkbook.Backbone.resolveSourceDirUnder())
                .id(MyWorkbook.Backbone.getId()));
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
        Path file = classOutputDir.resolve("test_writeDiagram_Options_KAZURAYAM.puml");
        app.writeDiagram(file);
        assertThat(file).exists();
        assertThat(file.toFile().length()).isGreaterThan(0);
    }

    @Test
    public void test_writeDiagram_Options_RELAXED() throws IOException {
        Path file = classOutputDir.resolve("test_writeDiagram_Options_RELAXED.puml");
        app.setOptions(Options.RELAXED);
        app.writeDiagram(file);
        assertThat(file).exists();
        assertThat(file.toFile().length()).isGreaterThan(0);
    }

    @Test
    public void test_writeDiagram_Options_DEFAULT() throws IOException {
        Path file = classOutputDir.resolve("test_writeDiagram_Options_DEFAULT.puml");
        app.setOptions(Options.DEFAULT);
        app.writeDiagram(file);
        assertThat(file).exists();
        assertThat(file.toFile().length()).isGreaterThan(0);
    }
}

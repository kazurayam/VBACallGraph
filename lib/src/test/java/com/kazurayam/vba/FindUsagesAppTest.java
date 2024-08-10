package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class FindUsagesAppTest {

    private static final Logger logger =
            LoggerFactory.getLogger(FindUsagesAppTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(FindUsagesAppTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(FindUsagesAppTest.class)
                    .build();

    private static final Path baseDir = too.getProjectDirectory().resolve("src/test/fixture/hub");
    private FindUsagesApp app;
    private Path classOutputDir;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        app = new FindUsagesApp();
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
        app.setExcludeUnittestModules(true);
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
    public void test_writeDiagram() throws IOException {
        Path file = classOutputDir.resolve("test_writeDiagram.pu");
        app.writeDiagram(file);
        assertThat(file).exists();
        assertThat(file.toFile().length()).isGreaterThan(0);
    }
}

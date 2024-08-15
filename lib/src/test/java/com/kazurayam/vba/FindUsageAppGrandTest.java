package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.FindUsageAppFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class FindUsageAppGrandTest {

    private static final Logger logger =
            LoggerFactory.getLogger(FindUsageAppGrandTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(FindUsageAppGrandTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(FindUsageAppGrandTest.class)
                    .build();

    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");
    private FindUsageApp app;
    private Path classOutputDir;
    private Path original;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        app = FindUsageAppFactory.createKazurayamSeven();
    }

    @Test
    public void test_writeDiagram_Options_KAZURAYAM() throws IOException {
        original = classOutputDir.resolve("test_writeDiagram_Options_KAZURAYAM.pu");
        app.writeDiagram(original);
        assertThat(original).exists();
        assertThat(original.toFile().length()).isGreaterThan(0);
    }

    @AfterMethod
    public void afterMethod() throws IOException, InterruptedException {
        Path image = classOutputDir.resolve("test_writeDiagram_Options_KAZURAYAM.png");
        assertThat(image).exists();
        // create a PDF from a PNG
        Path originalFileName = PDFFromImageGenerator.resolvePDFFileNameFromImage(image);
        Path original = classOutputDir.resolve(originalFileName);
        PDFFromImageGenerator.generate(image, original);
        // modify the original PDF to a poster PDF
        MutoolPosterRunner runner =
                new MutoolPosterRunner.Builder()
                        .x(2)
                        .y(2)
                        .original(original)
                        .build();
        runner.run();
    }
}

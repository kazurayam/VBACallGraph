package com.kazurayam.vba.printing;

import com.kazurayam.unittest.TestOutputOrganizer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class PDFFromImageGeneratorTest {

    private static final Logger logger =
            LoggerFactory.getLogger(PDFFromImageGeneratorTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(PDFFromImageGeneratorTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(PDFFromImageGeneratorTest.class)
                    .build();

    private static final String pngFileName =
            "FindUsageAppGrandTest.testWriteDiagram_Options_KAZURAYAM.png";

    private Path classOutputDir;
    private Path pngFile;

    @BeforeTest
    public void beforeTest() throws IOException {
        Path fixtureDir = too.getProjectDirectory().resolve("src/test/fixture");
        pngFile = fixtureDir.resolve("diagram").resolve(pngFileName);
        classOutputDir = too.cleanClassOutputDirectory();
    }

    @Test
    public void test_generate() throws IOException {
        Path pdfFileName = PDFFromImageGenerator.resolvePDFFileNameFromImage(pngFile);
        Path pdfFile = classOutputDir.resolve((pdfFileName));
        PDFFromImageGenerator.generate(pngFile, pdfFile);
        assertThat(pdfFile).exists();
    }


}

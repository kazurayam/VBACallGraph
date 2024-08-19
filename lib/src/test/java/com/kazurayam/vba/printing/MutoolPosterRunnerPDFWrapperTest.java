package com.kazurayam.vba.printing;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.printing.MutoolPosterRunner.PDFWrapper;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;


public class MutoolPosterRunnerPDFWrapperTest {

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(MutoolPosterRunnerPDFWrapperTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(MutoolPosterRunnerPDFWrapperTest.class)
                    .build();

    private Path classOutputDir;
    private Path original;
    private static final String originalFileName =
            "CallGraphAppGrandTest.testWriteDiagram_Options_KAZURAYAM.pdf";

    @BeforeTest
    public void beforeTest() throws IOException {
        Path fixtureDir = too.getProjectDirectory().resolve("src/test/fixture");
        original = fixtureDir.resolve("diagram").resolve(originalFileName);
        classOutputDir = too.cleanClassOutputDirectory();
    }

    @Test
    public void test_constructor() throws IOException {
        PDFWrapper instance = new PDFWrapper(original);
        assertThat(instance).isNotNull();
    }

    @Test
    public void test_getPage0() throws IOException {
        PDFWrapper instance = new PDFWrapper(original);
        PDPage page = instance.getPage(0);
        assertThat(page).isNotNull();
    }

    @Test
    public
    void test_point2mm() {
        float points = 100.0f;
        float mm = PDFWrapper.point2mm(points);
        assertThat(Math.floor(mm)).isEqualTo(35.0f);  // 35.2777777778 mm to be exact
    }

    @Test
    public void test_getRectangle0() throws IOException {
        PDFWrapper instance = new PDFWrapper(original);
        PDRectangle rectangle = instance.getRectangle(0);
        assertThat(Math.floor(rectangle.getWidth())).isEqualTo(765.0f);
        assertThat(Math.floor(rectangle.getHeight())).isEqualTo(1173.0f);
    }
}

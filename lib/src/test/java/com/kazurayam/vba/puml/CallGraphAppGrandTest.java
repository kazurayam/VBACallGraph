package com.kazurayam.vba.puml;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.printing.MutoolPosterRunner;
import com.kazurayam.vba.printing.PDFFromImageGenerator;
import com.kazurayam.vba.example.CallGraphAppFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.kazurayam.vba.printing.PlantUMLRunner;
import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class CallGraphAppGrandTest {

    private static final Logger logger =
            LoggerFactory.getLogger(CallGraphAppGrandTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(CallGraphAppGrandTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(CallGraphAppGrandTest.class)
                    .build();

    private CallGraphApp app;
    private Path classOutputDir;
    private Path puml;
    private static String FILE_NAME_BODY = "test_writeDiagram_Options_KAZURAYAM";

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        app = CallGraphAppFactory.createKazurayamSevenPlus();
    }

    @Test
    public void test_writeDiagram_Options_KAZURAYAM() throws IOException {
        puml = classOutputDir.resolve(FILE_NAME_BODY + ".puml");
        app.writeDiagram(puml);
        assertThat(puml).exists();
        assertThat(puml.toFile().length()).isGreaterThan(0);
    }

    @AfterTest
    public void afterTest() throws IOException, InterruptedException {
        // create a PNG from a puml by PlantUML
        puml = classOutputDir.resolve(FILE_NAME_BODY + ".puml");
        PlantUMLRunner plantuml =
                new PlantUMLRunner.Builder()
                        .workingDirectory(classOutputDir)
                        .puml(puml)
                        .outdir(classOutputDir)
                        .build();
        plantuml.run();
        Path image = classOutputDir.resolve(FILE_NAME_BODY + ".png");
        assertThat(image).exists();

        // create a PDF from a PNG
        Path originalFileName = PDFFromImageGenerator.resolvePDFFileNameFromImage(image);
        Path original = classOutputDir.resolve(originalFileName);
        PDFFromImageGenerator.generate(image, original);

        // modify the original PDF to a poster PDF
        MutoolPosterRunner mutool =
                new MutoolPosterRunner.Builder()
                        .original(original)
                        .pieceSize("A2")
                        .build();
        mutool.run();
        Path poster = classOutputDir.resolve(FILE_NAME_BODY + "-poster.pdf");
        assertThat(poster).exists();
    }
}

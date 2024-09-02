package com.kazurayam.vba.printing;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.puml.CallGraphApp;
import com.kazurayam.vba.puml.Options;
import com.kazurayam.vba.puml.ModelWorkbook;
import com.kazurayam.vba.example.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class PlantUMLRunnerTest {
    private static final Logger logger =
            LoggerFactory.getLogger(PlantUMLRunnerTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(PlantUMLRunnerTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(PlantUMLRunnerTest.class)
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
        app.setOptions(Options.DEFAULT);
        //
    }

    @Test
    public void test_smoke() throws IOException, InterruptedException {
        Path puml = classOutputDir.resolve("test_smoke.puml");
        app.writeDiagram(puml);
        assertThat(puml).exists();
        //
        PlantUMLRunner runner =
                new PlantUMLRunner.Builder()
                        .workingDirectory(classOutputDir)
                        .puml(puml)
                        .outdir(classOutputDir)
                        .build();
        runner.run();
        Path out = classOutputDir.resolve("test_smoke.png");
        assertThat(out).exists();
        assertThat(out.toFile().length()).isGreaterThan(15000);
    }
}
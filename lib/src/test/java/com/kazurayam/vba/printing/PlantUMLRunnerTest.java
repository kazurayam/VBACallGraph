package com.kazurayam.vba.printing;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.puml.FindUsageApp;
import com.kazurayam.vba.puml.Options;
import com.kazurayam.vba.puml.SensibleWorkbook;
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
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir))
        );
        app.setOptions(Options.DEFAULT);
        //
    }

    @Test
    public void test_smoke() throws IOException, InterruptedException {
        Path pu = classOutputDir.resolve("test_smoke.puml");
        app.writeDiagram(pu);
        assertThat(pu).exists();
        //
        PlantUMLRunner runner = new PlantUMLRunner();
        runner.workingDirectory(classOutputDir);
        runner.setDiagram(pu);
        runner.run();
        Path out = classOutputDir.resolve("out.pdf");
        assertThat(out).exists();
        assertThat(out.toFile().length()).isGreaterThan(1000000);
    }
}
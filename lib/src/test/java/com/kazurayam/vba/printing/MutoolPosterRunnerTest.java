package com.kazurayam.vba.printing;


import com.kazurayam.unittest.TestOutputOrganizer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class MutoolPosterRunnerTest {


    private static final Logger logger =
            LoggerFactory.getLogger(MutoolPosterRunnerTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(MutoolPosterRunnerTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(MutoolPosterRunnerTest.class)
                    .build();

    private Path classOutputDir;
    private Path original;
    private static final String originalFileName =
            "FindUsageAppGrandTest.testWriteDiagram_Options_KAZURAYAM.pdf";

    @BeforeTest
    public void beforeTest() throws IOException {
        Path fixtureDir = too.getProjectDirectory().resolve("src/test/fixture");
        original = fixtureDir.resolve("diagram").resolve(originalFileName);
        classOutputDir = too.cleanClassOutputDirectory();
    }

    @Test
    public void test_run() throws IOException, InterruptedException {
        MutoolPosterRunner runner =
                new MutoolPosterRunner.Builder()
                        .x(2)
                        .y(2)
                        .original(original)
                        .build();
        runner.run();
        Path poster = runner.getPoster();
        assertThat(poster).exists();
        assertThat(poster.toFile().length()).isGreaterThan(0);
        assertThat(poster.getFileName().toString()).endsWith("-poster.pdf");
    }


}

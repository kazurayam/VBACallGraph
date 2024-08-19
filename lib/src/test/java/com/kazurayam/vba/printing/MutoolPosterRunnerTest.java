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
            "CallGraphAppGrandTest.testWriteDiagram_Options_KAZURAYAM.pdf";

    @BeforeTest
    public void beforeTest() throws IOException {
        Path fixtureDir = too.getProjectDirectory().resolve("src/test/fixture");
        original = fixtureDir.resolve("diagram").resolve(originalFileName);
        classOutputDir = too.cleanClassOutputDirectory();
    }

    @Test
    public void test_run_xy() throws IOException, InterruptedException {
        MutoolPosterRunner runner =
                new MutoolPosterRunner.Builder()
                        .original(original)
                        .x(2)
                        .y(2)
                        .build();
        assertThat(runner.getX()).isEqualTo(2);
        assertThat(runner.getY()).isEqualTo(2);
        runner.run();
        Path poster = runner.getPoster();
        assertThat(poster).exists();
        assertThat(poster.toFile().length()).isGreaterThan(0);
        assertThat(poster.getFileName().toString()).endsWith("-poster.pdf");
    }

    @Test
    public void test_run_pieceSize() throws IOException, InterruptedException {
        String pieceSize = "A3";
        Path output = classOutputDir.resolve(pieceSize + "-poster.pdf");
        MutoolPosterRunner runner =
                new MutoolPosterRunner.Builder()
                        .original(original)
                        .pieceSize(pieceSize)
                        .poster(output)
                        .build();
        assertThat(runner.getX()).isEqualTo(3);
        assertThat(runner.getY()).isEqualTo(3);
        runner.run();
        Path poster = runner.getPoster();
        assertThat(poster).exists();
        assertThat(poster.toFile().length()).isGreaterThan(0);
        assertThat(poster.getFileName().toString()).endsWith("-poster.pdf");
    }
}

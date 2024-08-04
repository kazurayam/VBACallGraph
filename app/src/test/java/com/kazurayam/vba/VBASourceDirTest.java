package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import org.testng.annotations.Test;

import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class VBASourceDirTest {

    private TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(VBASourceDirTest.class)
                    .subOutputDirectory(VBASourceDirTest.class).build();

    private Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");

    @Test
    public void test_all_resolveBasedOn() {
        VBASourceDir[] values = VBASourceDir.values();
        for (int i = 0; i < values.length; i++) {
            VBASourceDir ex = values[i];
            Path path = ex.resolveBasedOn(baseDir);
            assertThat(path).exists();
        }
    }
}

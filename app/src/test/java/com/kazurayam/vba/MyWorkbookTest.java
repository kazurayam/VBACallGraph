package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import org.testng.annotations.Test;

import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class MyWorkbookTest {

    private TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(MyWorkbookTest.class)
                    .subOutputDirectory(MyWorkbookTest.class).build();

    private Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");

    @Test
    public void test_all_resolveBasedOn() {
        MyWorkbook[] values = MyWorkbook.values();
        for (int i = 0; i < values.length; i++) {
            MyWorkbook ex = values[i];
            assertThat(ex.resolveWorkbookBasedOn(baseDir)).exists();
            assertThat(ex.resolveVBASourceDirBasedOn(baseDir)).exists();
        }
    }
}

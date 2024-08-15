package com.kazurayam.vba.example;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import org.assertj.core.api.Assertions;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class MyWorkbookTest {

    Logger logger = LoggerFactory.getLogger(MyWorkbookTest.class);

    private TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(MyWorkbookTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(MyWorkbookTest.class)
                    .build();

    private Path baseDir = too.getProjectDirectory().resolve("src/test/fixture/hub");

    @Test
    public void test_all_resolveWorkbookUnder() {
        MyWorkbook[] values = MyWorkbook.values();
        for (MyWorkbook ex : values) {
            Assertions.assertThat(ex.resolveWorkbookUnder(baseDir)).exists();
        }
    }
    @Test
    public void test_all_resolveSourceDirUnder() {
        MyWorkbook[] values = MyWorkbook.values();
        for (MyWorkbook ex : values) {
            Assertions.assertThat(ex.resolveWorkbookUnder(baseDir)).exists();
            Assertions.assertThat(ex.resolveSourceDirUnder(baseDir)).exists();
        }
    }


    @Test
    public void test_toJson() throws JsonProcessingException {
        String json = MyWorkbook.Cashbook.toJson();
        assertThat(json).isNotNull();
        logger.info("[test_toJson] " + json);
        assertThat(json).contains("id");
        assertThat(json).contains("Cashbook");
        assertThat(json).contains("repositoryName");
        assertThat(json).contains("workbookSubPath");
        assertThat(json).contains("sourceDirSubPath");
    }

    @Test
    public void test_toString() throws IOException {
        Path methodOutputDir = too.cleanMethodOutputDirectory("test_toString");
        MyWorkbook[] values = MyWorkbook.values();
        for (MyWorkbook ex : values) {
            String json = ex.toString();
            assertThat(json).isNotNull();
            Path p = methodOutputDir.resolve(ex.getId() + ".json");
            Files.writeString(p, json);
        }
    }
}

package com.kazurayam.vbaexample;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import org.assertj.core.api.Assertions;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.Test;

import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class MyWorkbookInstanceLocationTest {

    Logger logger = LoggerFactory.getLogger(MyWorkbookInstanceLocationTest.class);

    private TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(MyWorkbookInstanceLocationTest.class)
                    .subOutputDirectory(MyWorkbookInstanceLocationTest.class).build();

    private Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");

    @Test
    public void test_all_resolveBasedOn() {
        MyWorkbook[] values = MyWorkbook.values();
        for (int i = 0; i < values.length; i++) {
            MyWorkbook ex = values[i];
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
    public void test_toString() {
        String json = MyWorkbook.Cashbook.toString();
        assertThat(json).isNotNull();
        logger.info("[test_toString] " + json);
    }
}

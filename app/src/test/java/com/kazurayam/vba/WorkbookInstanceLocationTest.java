package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.Test;

import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class WorkbookInstanceLocationTest {

    Logger logger = LoggerFactory.getLogger(WorkbookInstanceLocationTest.class);

    private TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(WorkbookInstanceLocationTest.class)
                    .subOutputDirectory(WorkbookInstanceLocationTest.class).build();

    private Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");

    @Test
    public void test_all_resolveBasedOn() {
        WorkbookInstanceLocation[] values = WorkbookInstanceLocation.values();
        for (int i = 0; i < values.length; i++) {
            WorkbookInstanceLocation ex = values[i];
            assertThat(ex.resolveWorkbookBasedOn(baseDir)).exists();
            assertThat(ex.resolveSourceDirBasedOn(baseDir)).exists();
        }
    }

    @Test
    public void test_toJson() throws JsonProcessingException {
        String json = WorkbookInstanceLocation.Cashbook.toJson();
        assertThat(json).isNotNull();
        logger.info("[test_toJson] " + json);
    }

    @Test
    public void test_toString() {
        String json = WorkbookInstanceLocation.Cashbook.toString();
        assertThat(json).isNotNull();
        logger.info("[test_toString] " + json);
    }
}

package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class WorkbookInstanceTest {

    private Logger logger = LoggerFactory.getLogger(WorkbookInstanceTest.class);
    private TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(WorkbookInstanceTest.class)
                    .subOutputDirectory(WorkbookInstanceTest.class)
                    .build();
    private Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");
    private WorkbookInstance wbi;
    @BeforeTest
    public void beforeTest() {
        wbi = new WorkbookInstance(baseDir, WorkbookInstanceLocation.Cashbook);
    }
    @Test
    public void test_Cashbook_isNotNull() {
        assertThat(wbi).isNotNull();
    }

    @Test
    public void test_toJson() throws JsonProcessingException {
        String json = wbi.toJson();
        assertThat(json).isNotNull();
        logger.info("[test_toJson] " + json);
        assertThat(json).contains("id");
        assertThat(json).contains("Cashbook");
        assertThat(json).contains("repositoryName");
        assertThat(json).contains("workbookSubPath");
        assertThat(json).contains("sourceDirSubPath");
    }
}

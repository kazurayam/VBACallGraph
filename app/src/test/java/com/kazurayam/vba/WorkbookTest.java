package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.SortedMap;

import static org.assertj.core.api.Assertions.assertThat;

public class WorkbookTest {

    private static final Logger logger = LoggerFactory.getLogger(WorkbookTest.class);
    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(WorkbookTest.class)
                    .subOutputDirectory(WorkbookTest.class)
                    .build();
    private static final Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");
    private Workbook wb;
    @BeforeTest
    public void beforeTest() throws IOException {
        wb = new Workbook(
                MyWorkbook.Cashbook.getId(),
                MyWorkbook.Cashbook.resolveWorkbookUnder(baseDir),
                MyWorkbook.Cashbook.resolveSourceDirUnder(baseDir));
    }
    @Test
    public void test_Cashbook_isNotNull() {
        assertThat(wb).isNotNull();
    }

    @Test
    public void test_toJson() throws JsonProcessingException {
        String json = wb.toJson();
        assertThat(json).isNotNull();
        logger.info("[test_toJson] " + json);
        assertThat(json).contains("id");
        assertThat(json).contains("Cashbook");
        assertThat(json).contains("workbookPath");
        assertThat(json).contains("sourceDirPath");
    }

    @Test
    public void test_getModuleProcedures() throws IOException {
        Workbook wb = new Workbook(
                MyWorkbook.Member.getId(),
                MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                MyWorkbook.Member.resolveSourceDirUnder(baseDir));
        SortedMap<VBAModule, List<Procedure>> moduleProcedures =
                wb.getModuleProcedures();
        assertThat(moduleProcedures.keySet().size()).isEqualTo(3);
    }
}

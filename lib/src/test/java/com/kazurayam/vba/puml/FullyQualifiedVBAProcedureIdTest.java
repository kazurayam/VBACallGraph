package com.kazurayam.vba.puml;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.example.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;


public class FullyQualifiedVBAProcedureIdTest {
    private static final Logger logger = LoggerFactory.getLogger(FullyQualifiedVBAProcedureIdTest.class);
    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(FullyQualifiedVBAProcedureIdTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(FullyQualifiedVBAProcedureIdTest.class)
                    .build();
    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");

    private Path classOutputDir;
    private FullyQualifiedVBAProcedureId fqpi;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        SensibleWorkbook wb = new SensibleWorkbook(
                MyWorkbook.Member.getId(),
                MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                MyWorkbook.Member.resolveSourceDirUnder(baseDir));
        VBAModule module = wb.getModule("AoMemberUtils");
        VBAProcedure procedure = module.getProcedure("FetchMemberTable");
        fqpi = new FullyQualifiedVBAProcedureId(wb, module, procedure);
    }

    @Test
    public void test_getWorkbook() {
        assertThat(fqpi.getWorkbook().getId()).isEqualTo(MyWorkbook.Member.getId());
    }
    @Test
    public void test_getWorkbookId() {
        assertThat(fqpi.getWorkbookId()).isEqualTo(MyWorkbook.Member.getId());
    }
    @Test
    public void test_getModule() {
        assertThat(fqpi.getModule().getName()).isEqualTo("AoMemberUtils");
    }
    @Test
    public void test_getModuleName() {
        assertThat(fqpi.getModuleName()).isEqualTo("AoMemberUtils");
    }
    @Test
    public void test_getProcedure() {
        assertThat(fqpi.getProcedure().getName()).isEqualTo("FetchMemberTable");
    }
    @Test
    public void test_getProcedureName() {
        assertThat(fqpi.getProcedureName()).isEqualTo("FetchMemberTable");
    }
    @Test
    public void test_toJson() throws JsonProcessingException {
        String json = fqpi.toJson();
        assertThat(json).isNotNull();
        logger.info("[test_toJson] " + json);
    }
    @Test
    public void test_toString() throws IOException {
        Path out = classOutputDir.resolve("test_toString.json");
        String json = fqpi.toString();
        Files.writeString(out, json);
    }
}

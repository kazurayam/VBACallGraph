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

public class FullyQualifiedVBAModuleIdTest {
    private static final Logger logger = LoggerFactory.getLogger(FullyQualifiedVBAModuleIdTest.class);
    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(FullyQualifiedVBAModuleIdTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(FullyQualifiedVBAModuleIdTest.class)
                    .build();

    private Path classOutputDir;
    private FullyQualifiedVBAModuleId fqmi;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        ModelWorkbook wb = new ModelWorkbook(
                MyWorkbook.Member.resolveWorkbookUnder(),
                MyWorkbook.Member.resolveSourceDirUnder())
                .id(MyWorkbook.Member.getId());
        VBAModule module = wb.getModule("MbMemberTableUtil");
        fqmi = new FullyQualifiedVBAModuleId(wb, module);
    }

    @Test
    public void test_getWorkbook() {
        assertThat(fqmi.getWorkbook().getId()).isEqualTo(MyWorkbook.Member.getId());
    }
    @Test
    public void test_getWorkbookId() {
        assertThat(fqmi.getWorkbookId()).isEqualTo(MyWorkbook.Member.getId());
    }
    @Test
    public void test_getModule() {
        assertThat(fqmi.getModule().getName()).isEqualTo("MbMemberTableUtil");
    }
    @Test
    public void test_getModuleName() {
        assertThat(fqmi.getModuleName()).isEqualTo("MbMemberTableUtil");
    }
    @Test
    public void test_toJson() throws JsonProcessingException {
        String json = fqmi.toJson();
        assertThat(json).isNotNull();
        logger.info("[test_toJson] " + json);
    }
    @Test
    public void test_toString() throws IOException {
        Path out = classOutputDir.resolve("test_toString.json");
        String json = fqmi.toString();
        Files.writeString(out, json);
    }
}

package com.kazurayam.vba;


import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
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
    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");

    private Path classOutputDir;
    private FullyQualifiedVBAModuleId fqmi;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        SensibleWorkbook wb = new SensibleWorkbook(
                MyWorkbook.Member.getId(),
                MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                MyWorkbook.Member.resolveSourceDirUnder(baseDir));
        VBAModule module = wb.getModule("AoMemberUtils");
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
        assertThat(fqmi.getModule().getName()).isEqualTo("AoMemberUtils");
    }
    @Test
    public void test_getModuleName() {
        assertThat(fqmi.getModuleName()).isEqualTo("AoMemberUtils");
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

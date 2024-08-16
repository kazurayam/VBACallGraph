package com.kazurayam.vba.puml;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class VBAProcedureTest {

    private static final Logger logger = LoggerFactory.getLogger(VBAProcedureTest.class);
    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(VBAProcedureTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(VBAProcedureTest.class)
                    .build();
    private Path classOutputDir;
    private VBAProcedure proc;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        proc = new VBAProcedure.Builder()
                .project("Cashbook")
                .moduleType("Class")
                .module("Account")
                .scope("Public")
                .procKind("Sub")
                .procedure("AccountName")
                .lineNo(68)
                .source("Public Property Get AccountName() as String")
                .comment("Sun also rises")
                .build();
    }

    @Test
    public void test_getProject() {
        assertThat(proc.getProject()).isEqualTo("Cashbook");
    }

    @Test
    public void test_getModuleType() {
        assertThat(proc.getModuleType()).isEqualTo(VBAModule.ModuleType.Class);
    }

    @Test
    public void test_getModule() {
        assertThat(proc.getModule()).isEqualTo("Account");
    }

    @Test
    public void test_getScope() {
        assertThat(proc.getScope()).isEqualTo(VBAProcedure.Scope.Public);
    }

    @Test
    public void test_getProcKind() {
        assertThat(proc.getProcKind()).isEqualTo(VBAProcedure.ProcKind.Sub);
    }

    @Test
    public void test_getProcedure() {
        assertThat(proc.getProcedure()).isEqualTo("AccountName");
    }

    @Test
    public void test_getSourceFileName() {
        assertThat(proc.getSourceFileName()).isEqualTo("Account.cls");
    }

    @Test
    public void test_getLineNo() {
        assertThat(proc.getLineNo()).isEqualTo(68);
    }

    @Test
    public void test_getSource() {
        assertThat(proc.getSource()).contains("Public Property Get AccountName() as String");
    }
    @Test
    public void test_getComment() {
        assertThat(proc.getComment()).contains("Sun also rises");
    }

    @Test
    public void test_toJson() throws JsonProcessingException {
        logger.info("[test_toJson] " + proc.toJson());
        assertThat(proc.toJson()).contains("Sun also rises");
    }
    @Test
    public void test_toString() throws IOException {
        String json = proc.toString();
        logger.info("[test_toString] " + json);
        assertThat(json).contains("Sun also rises");
        Files.writeString(classOutputDir.resolve("test_toString.json"), json);
    }
}

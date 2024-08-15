package com.kazurayam.vba.puml;

import com.kazurayam.unittest.TestOutputOrganizer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

public class VBAModuleTest {

    private static final Logger logger = LoggerFactory.getLogger(VBAModuleTest.class);
    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(VBAModuleTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(VBAModuleTest.class)
                    .build();

    private VBAModule module;
    private Path classOutputDir;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        module = new VBAModule("Account", VBAModule.ModuleType.Class);
        VBAProcedure proc = new VBAProcedure.Builder()
                .name("AccountName")
                .module("Account")
                .type("Class")
                .scope("Public")
                .subOrFunc("Sub")
                .lineNo(68)
                .source("Public Property Get AccountName() as String")
                .comment("Sun also rises")
                .build();
        module.add(proc);
    }

    @Test
    public void test_getName() {
        assertThat(module.getName()).isEqualTo("Account");
    }

    @Test
    public void test_toString() throws IOException {
        String prettyJson = module.toString();
        logger.info("[test_toString] " + prettyJson);
        assertThat(prettyJson).contains("module");
        assertThat(prettyJson).contains("Account");
        assertThat(prettyJson).contains("procedures");
        Files.writeString(classOutputDir.resolve("test_toString.json"), prettyJson);
    }
}

package com.kazurayam.vba;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import static org.assertj.core.api.Assertions.assertThat;

public class VBAModuleTest {

    private Logger logger = LoggerFactory.getLogger(VBAModuleTest.class);

    private VBAModule module;

    @BeforeTest
    public void beforeTest() {
        module = new VBAModule("Account");
        VBAProcedure proc = new VBAProcedure.Builder()
                .name("AccountName")
                .module("Account")
                .scope(VBAProcedure.Scope.Public)
                .subOrFunc(VBAProcedure.SubOrFunc.Sub)
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
    public void test_toString() {
        String prettyJson = module.toString();
        logger.info("[test_toString] " + prettyJson);
        assertThat(prettyJson).contains("module");
        assertThat(prettyJson).contains("Account");
        assertThat(prettyJson).contains("procedures");
    }
}

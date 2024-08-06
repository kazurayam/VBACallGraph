package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonProcessingException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import static org.assertj.core.api.Assertions.assertThat;

public class VBAProcedureTest {

    Logger logger = LoggerFactory.getLogger(VBAProcedureTest.class);

    private VBAProcedure proc;

    @BeforeTest
    public void beforeTest() {
        proc = new VBAProcedure.Builder()
                .name("AccountName")
                .module("Account")
                .scope(Scope.Public)
                .subOrFunc(SubOrFunc.Sub)
                .lineNo(68)
                .source("Public Property Get AccountName() as String")
                .comment("Sun also rises")
                .build();
    }

    @Test
    public void test_name() {
        assertThat(proc.getName()).isEqualTo("AccountName");
    }
    @Test
    public void test_module() {
        assertThat(proc.getModule()).isEqualTo("Account");
    }
    @Test
    public void test_scope() {
        assertThat(proc.getScope()).isEqualTo(Scope.Public);
    }
    @Test
    public void test_subOrFunc() {
        assertThat(proc.getSubOrFunc()).isEqualTo(SubOrFunc.Sub);
    }
    @Test
    public void test_lineNo() {
        assertThat(proc.getLineNo()).isEqualTo(68);
    }
    @Test
    public void test_source() {
        assertThat(proc.getSource()).contains("Public Property Get AccountName() as String");
    }
    @Test
    public void test_comment() {
        assertThat(proc.getComment()).contains("Sun also rises");
    }
    @Test
    public void test_toJson() throws JsonProcessingException {
        logger.info("[test_toJson] " + proc.toJson());
        assertThat(proc.toJson()).contains("Sun also rises");
    }
    @Test
    public void test_toString() {
        logger.info("[test_toString] " + proc.toString());
        assertThat(proc.toString()).contains("Sun also rises");
    }
}

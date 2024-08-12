package com.kazurayam.vba;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.Test;

import static org.assertj.core.api.Assertions.assertThat;

import java.lang.reflect.Method;
import java.util.regex.Pattern;
import java.util.List;

public class PatternManagerTest {

    private static Logger logger = LoggerFactory.getLogger(PatternManagerTest.class);

    @Test
    public void test_escapeAsRegex() {
        String escaped = PatternManager.escapeAsRegex(".FetchMemberTable(");
        assertThat(escaped).isEqualTo("\\.FetchMemberTable\\(");
    }

    @Test
    public void test_createPatterns_dot_parenthesis() {
        List<Pattern> patterns =
                PatternManager.createPatterns(".FetchMemberTable(");
        assertThat(patterns).isNotEmpty();
        Method m = new Object(){}.getClass().getEnclosingMethod();
        logger.debug(String.format("[%s] %s", m.getName(), patterns));
    }
}

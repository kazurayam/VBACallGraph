package com.kazurayam.vba;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.Test;

import static org.assertj.core.api.Assertions.assertThat;

import java.lang.reflect.Method;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.List;

public class ProcedureNamePatternManagerTest {

    private static Logger logger = LoggerFactory.getLogger(ProcedureNamePatternManagerTest.class);

    @Test
    public void test_escapeAsRegex() {
        String escaped = ProcedureNamePatternManager.escapeAsRegex(".FetchMemberTable(");
        assertThat(escaped).isEqualTo("\\.FetchMemberTable\\(");
    }

    @Test
    public void test_createPatterns_dot_parenthesis() {
        List<Pattern> patterns =
                ProcedureNamePatternManager.createPatterns(".FetchMemberTable(");
        assertThat(patterns).isNotEmpty();
        Method m = new Object(){}.getClass().getEnclosingMethod();
        logger.debug(String.format("[%s] %s", m.getName(), patterns));
    }

    @Test
    public void test_createPatterns_Test() {
        List<Pattern> patterns =
                ProcedureNamePatternManager.createPatterns("Test");
        assertThat(patterns).isNotEmpty();
        Method m = new Object(){}.getClass().getEnclosingMethod();
        logger.debug(String.format("[%s] %s", m.getName(), patterns));
        //
        tryPattern(patterns.get(0), "    Call Test", true);
        tryPattern(patterns.get(0), "    Test()", true);
        tryPattern(patterns.get(0), "    Call Test_Foo", false);
    }

    private void tryPattern(Pattern ptn, String source, Boolean expect) {
        Matcher matcher = ptn.matcher(source);
        Boolean actual = matcher.find();
        assertThat(actual)
                .as(String.format(
                        "Pattern \"%s\" was expected to be %s in the source \"%s\" but ...",
                        ptn.pattern(),
                        (expect) ? "found" : "not found",
                        source))
                .isEqualTo(expect);
    }
}

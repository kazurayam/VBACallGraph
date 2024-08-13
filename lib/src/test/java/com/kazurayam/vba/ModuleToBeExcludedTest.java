package com.kazurayam.vba;

import org.testng.annotations.Test;

import static org.assertj.core.api.Assertions.assertThat;

public class ModuleToBeExcludedTest {

    @Test
    public void test_matches_STARTS_WITH_TEST() {
        applyMatches(ModuleToBeExcluded.STARTS_WITH_TEST,
                "Test_Foo");
    }

    @Test
    public void test_matches_ENDS_WITH_TEST() {
        applyMatches(ModuleToBeExcluded.ENDS_WITH_TEST,
                "FooTest");
    }

    @Test
    public void test_matches_プロシージャ一覧を作る() {
        applyMatches(ModuleToBeExcluded.プロシージャ一覧を作る,
                "プロシージャ一覧を作る");
    }

    void applyMatches(ModuleToBeExcluded instance, String target) {
        Boolean result = instance.find(target);
        assertThat(result)
                .as(String.format("%s matches %s: %b",
                        instance.getPattern().toString(),
                        target, result))
                .isTrue();
    }
}

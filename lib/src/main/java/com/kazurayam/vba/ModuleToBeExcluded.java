package com.kazurayam.vba;

import java.util.regex.Pattern;

public class ModuleToBeExcluded {

    public static ModuleToBeExcluded STARTS_WITH_TEST =
            new ModuleToBeExcluded(Pattern.compile(
                    "^Test\\.*", Pattern.CASE_INSENSITIVE));

    public static ModuleToBeExcluded ENDS_WITH_TEST =
            new ModuleToBeExcluded(Pattern.compile(
                    "\\.*Test$", Pattern.CASE_INSENSITIVE));

    public static ModuleToBeExcluded プロシージャー一覧を作る =
            new ModuleToBeExcluded(Pattern.compile(
                    "^プロシージャー一覧を作る$", Pattern.CASE_INSENSITIVE));

    public static ModuleToBeExcluded プロシージャ一覧を作る =
            new ModuleToBeExcluded(Pattern.compile(
                    "^プロシージャ一覧を作る$", Pattern.CASE_INSENSITIVE));

    public static ModuleToBeExcluded プロシジャ一覧を作る =
            new ModuleToBeExcluded(Pattern.compile(
                    "^プロシジャ一覧を作る$", Pattern.CASE_INSENSITIVE));

    private final Pattern pattern;

    public ModuleToBeExcluded(Pattern pattern) {
        this.pattern = pattern;
    }

    public Pattern getPattern() {
        return pattern;
    }

    public Boolean find(String moduleName) {
        return pattern.matcher(moduleName).find();
    }
}

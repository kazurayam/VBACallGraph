package com.kazurayam.vba.puml;

import java.util.regex.Pattern;

public class ModuleToBeExcluded {

    public static ModuleToBeExcluded STARTS_WITH_TEST =
            new ModuleToBeExcluded(Pattern.compile(
                    "^Test\\.*", Pattern.CASE_INSENSITIVE));

    public static ModuleToBeExcluded ENDS_WITH_TEST =
            new ModuleToBeExcluded(Pattern.compile(
                    "\\.*Test$", Pattern.CASE_INSENSITIVE));

    public static ModuleToBeExcluded XPORT =
            new ModuleToBeExcluded(Pattern.compile(
                    "^Xport$", Pattern.CASE_INSENSITIVE));

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

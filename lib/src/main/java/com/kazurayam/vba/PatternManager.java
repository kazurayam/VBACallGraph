package com.kazurayam.vba;

import java.util.HashSet;
import java.util.List;
import java.util.regex.Pattern;
import java.util.regex.PatternSyntaxException;
import java.util.Set;
import java.util.ArrayList;


public class PatternManager {

    private static final Set<Character> CHARS_TOBE_ESCAPED;

    static {
        char[] specialChars = ".[]{}()<>*+-=!?^$|".toCharArray();
        CHARS_TOBE_ESCAPED = new HashSet<>();
        for (char c : specialChars) {
            CHARS_TOBE_ESCAPED.add(c);
        }
    }

    public static String escapeAsRegex(String pattern) {
        char[] chars = pattern.toCharArray();
        StringBuilder sb = new StringBuilder();
        for (char c : chars) {
            if (CHARS_TOBE_ESCAPED.contains(c)) {
                sb.append("\\");
                sb.append(c);
            } else {
                sb.append(c);
            }
        }
        return sb.toString();
    }

    public static List<Pattern> createPatterns(String patternString) {
        String ptn = escapeAsRegex(patternString);
        List<Pattern> patterns = new ArrayList<>();
        try {
            patterns.add(Pattern.compile(ptn));
            return patterns;
        } catch (PatternSyntaxException e) {
            System.err.println("Pattern " + ptn + " could not be parsed");
        }
        return patterns;
    }

}

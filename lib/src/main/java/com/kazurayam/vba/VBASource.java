package com.kazurayam.vba;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * VBA Source code file (*.bas, *.cls)
 */
public class VBASource {

    private final String moduleName;
    private final Path sourcePath;

    private List<String> code;
    private boolean codeLoaded;

    public VBASource(String moduleName, Path sourcePath) {
        this.moduleName = moduleName;
        this.sourcePath = sourcePath;
        code = new ArrayList<>();
        codeLoaded = false;
    }

    public String getModuleName() {
        return moduleName;
    }

    public Path getSourcePath() {
        return sourcePath;
    }

    public List<VBASourceLine> find(String pattern) {
        Pattern ptn = Pattern.compile(escapeAsRegex(pattern));
        List<VBASourceLine> linesFound = new ArrayList<>();
        cache();
        for (int i = 0; i < code.size(); i++) {
            String line = code.get(i);
            Matcher m = ptn.matcher(line);
            if (m.find()) {
                VBASourceLine vbaSourceLine = new VBASourceLine(i, line);
                vbaSourceLine.setMatcher(m);
                linesFound.add(vbaSourceLine);
            }
        }
        return linesFound;
    }

    private static final Set<Character> CHARS_TOBE_ESCAPED;
    static {
        char[] specialChars = "\\.[]{}()<>*+-=!?^$|".toCharArray();
        CHARS_TOBE_ESCAPED = new HashSet<>();
        for (char c : specialChars) {
            CHARS_TOBE_ESCAPED.add(c);
        }
    }

    private String escapeAsRegex(String pattern) {
        char[] chars = pattern.toCharArray();
        StringBuilder sb = new StringBuilder();
        for (char c : chars) {
            if (CHARS_TOBE_ESCAPED.contains(c)) {
                sb.append(c);
            }
        }
        return sb.toString();
    }

    private void cache() {
        if (!codeLoaded) {
            try {
                code = loadCode();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    private List<String> loadCode() throws IOException {
        return Files.readAllLines(sourcePath);
    }

    /**
     *
     */
    public static class VBASourceLine {
        private final int lineNo;
        private final String line;
        private Matcher matcher;
        public VBASourceLine(int lineNo, String line) {
            this.lineNo = lineNo;
            this.line = line;
            this.matcher = null;
        }
        public int getLineNo() {
            return lineNo;
        }
        public String getLine() {
            return line;
        }
        public void setMatcher(Matcher matcher) {
            this.matcher = matcher;
        }
        /**
         * @return may be null
         */
        public Matcher getMatcher() {
            return this.matcher;
        }
    }
}

package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;
import java.nio.charset.Charset;
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
public class VBASource implements Comparable<VBASource> {

    private final String moduleName;
    private final Path sourcePath;

    private List<String> code;
    private boolean codeLoaded;

    private final static ObjectMapper mapper;
    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(VBASource.class, new VBASourceSerializer());
        module.addSerializer(VBASourceLine.class, new VBASourceLine.VBASourceLineSerializer());
        mapper.registerModule(module);
    }

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

    public List<String> getCode() {
        return code;
    }

    /**
     *
     */
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
                vbaSourceLine.setFound(true);
                linesFound.add(vbaSourceLine);
            }
        }
        return linesFound;
    }

    private static final Set<Character> CHARS_TOBE_ESCAPED;
    static {
        char[] specialChars = ".[]{}()<>*+-=!?^$|".toCharArray();
        CHARS_TOBE_ESCAPED = new HashSet<>();
        for (char c : specialChars) {
            CHARS_TOBE_ESCAPED.add(c);
        }
    }

    static String escapeAsRegex(String pattern) {
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

    private void cache() {
        if (!codeLoaded) {
            try {
                code = loadCode();
                codeLoaded = true;
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    /**
     * Read all lines in a .bas file (or a .cls file), which is encoded in Shift_JIS
     * on kazurayam's machine
     * @return List of all lines in a .bas file
     */
    List<String> loadCode() throws IOException {
        return Files.readAllLines(sourcePath, Charset.forName("MS932"));
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        VBASource other = (VBASource) o;
        if (moduleName.equals(other.moduleName)) {
            return sourcePath.equals(other.sourcePath);
        } else {
            return false;
        }
    }

    @Override
    public int hashCode() {
        int result = moduleName.hashCode();
        result = 31 * result + sourcePath.hashCode();
        return result;
    }

    @Override
    public String toString() {
        // pretty printed
        try {
            Object json = mapper.readValue(this.toJson(), Object.class);
            return mapper.writerWithDefaultPrettyPrinter().writeValueAsString(json);
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public int compareTo(VBASource other) {
        int moduleNameComparison = moduleName.compareTo(other.moduleName);
        if (moduleNameComparison == 0) {
            return sourcePath.compareTo(other.sourcePath);
        } else {
            return moduleNameComparison;
        }
    }

    public String toJson() throws JsonProcessingException {
        // no indent
        return mapper.writeValueAsString(this);
    }


    public static class VBASourceSerializer extends StdSerializer<VBASource> {
        public VBASourceSerializer() {
            this(null);
        }

        public VBASourceSerializer(Class<VBASource> t) {
            super(t);
        }

        @Override
        public void serialize(
                VBASource vbaSource, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject();     // {
            jgen.writeStringField("moduleName",
                    vbaSource.getModuleName());
            jgen.writeStringField("sourcePath",
                    vbaSource.getSourcePath().toString());
            // toString()の結果のJSONにcodeを含めるとJSONが大きくなるが、役に立たない。
            // だからcodeを含めない。
            /*
            if (vbaSource.codeLoaded) {
                jgen.writeArrayFieldStart("code");
                for (String line : vbaSource.getCode()) {
                    jgen.writeString(line);
                }
                jgen.writeEndArray();
            }
             */
            jgen.writeEndObject();       // }
        }
    }
}

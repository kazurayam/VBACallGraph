package com.kazurayam.vba.puml;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.util.regex.Matcher;
import java.io.IOException;

/**
 *
 */
public class VBASourceLine implements Comparable<VBASourceLine> {

    private final int lineNo;
    private final String line;

    private Boolean found;
    private Matcher matcher;

    private final static ObjectMapper mapper;
    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(VBASourceLine.class, new VBASourceLineSerializer());
        mapper.registerModule(module);
    }

    public VBASourceLine(int lineNo, String line) {
        this.lineNo = lineNo;
        this.line = line;
        this.found = false;
        this.matcher = null;
    }

    public int getLineNo() {
        return lineNo;
    }

    public String getLine() {
        return line;
    }

    public void setFound(Boolean found) { this.found = found; }
    public void setMatcher(Matcher matcher) {
        this.matcher = matcher;
    }

    public Boolean getFound() { return found; }

    /**
     * @return may be null
     */
    public Matcher getMatcher() {
        return this.matcher;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        VBASourceLine other = (VBASourceLine) o;
        if (lineNo == other.lineNo) {
            return line.equals(other.line);
        } else {
            return false;
        }
    }

    @Override
    public int hashCode() {
        int result = lineNo;
        result = 31 * result + line.hashCode();
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
    public int compareTo(VBASourceLine other) {
        int lineNoComparison = this.lineNo - other.lineNo;
        if (lineNoComparison == 0) {
            return this.line.compareTo(other.line);
        } else {
            return lineNoComparison;
        }
    }

    public String toJson() throws JsonProcessingException {
        // no indent
        return mapper.writeValueAsString(this);
    }

    /**
     *
     */
    public static class VBASourceLineSerializer extends StdSerializer<VBASourceLine> {
        public VBASourceLineSerializer() { this(null); }
        public VBASourceLineSerializer(Class<VBASourceLine> t) { super(t); }

        @Override
        public void serialize(
                VBASourceLine sl, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject();                                 // {
            jgen.writeNumberField("lineNo", sl.getLineNo()); // "lineNo":58,
            jgen.writeStringField("line", sl.getLine());     // "line":".........",
            jgen.writeEndObject();   // }
        }
    }
}

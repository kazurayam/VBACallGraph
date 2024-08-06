package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;
import java.nio.file.Path;

public class WorkbookInstance implements Comparable<WorkbookInstance> {
    private final Path baseDir;
    private final WorkbookInstanceLocation wbInstanceLocation;

    private static final ObjectMapper mapper;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(WorkbookInstance.class, new WorkbookInstanceSerializer());
        mapper.registerModule(module);
    }
    public WorkbookInstance(Path baseDir, WorkbookInstanceLocation myExcelFile) {
        this.baseDir = baseDir;
        this.wbInstanceLocation = myExcelFile;
    }
    public Path getBaseDir() {
        return baseDir;
    }
    public WorkbookInstanceLocation getWorkbookInstanceLocation() {
        return wbInstanceLocation;
    }
    @Override
    public boolean equals(Object obj) {
        if (!(obj instanceof WorkbookInstance other)) {
            return false;
        }
        if (this.baseDir == other.baseDir) {
            return (this.wbInstanceLocation == other.wbInstanceLocation);
        } else {
            return false;
        }
    }

    @Override
    public int hashCode() {
        int hash = 7;
        hash = 31 * hash + baseDir.hashCode();
        hash = 31 * hash + wbInstanceLocation.hashCode();
        return hash;
    }

    @Override
    public String toString() {
        //pretty print
        try {
            Object json = mapper.readValue(this.toJson(), Object.class);
            return mapper.writerWithDefaultPrettyPrinter().writeValueAsString(json);
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
    }

    public String toJson() throws JsonProcessingException {
        // without indent
        return mapper.writeValueAsString(this);
    }

    @Override
    public int compareTo(WorkbookInstance other) {
        int baseDirComparison = this.baseDir.compareTo(other.baseDir);
        if (baseDirComparison == 0) {
            return this.wbInstanceLocation.compareTo(other.wbInstanceLocation);
        } else {
            return baseDirComparison;
        }
    }

    private static class WorkbookInstanceSerializer
            extends StdSerializer<WorkbookInstance> {
        public WorkbookInstanceSerializer() { this(null); }
        public WorkbookInstanceSerializer(Class<WorkbookInstance> t) { super(t); }
        @Override
        public void serialize(
                WorkbookInstance wbi, JsonGenerator jgen, SerializerProvider provider)
        throws IOException {
            jgen.writeStartObject();
            jgen.writeStringField("baseDir", wbi.getBaseDir().toString());
            jgen.writeObjectField("workbookInstanceLocation",
                    wbi.getWorkbookInstanceLocation());
            jgen.writeEndObject();
        }
    }
}

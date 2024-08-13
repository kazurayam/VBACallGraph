package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;

public class FullyQualifiedVBAModuleId implements Comparable<FullyQualifiedVBAModuleId> {

    private final SensibleWorkbook workbook;
    private final VBAModule module;

    private final static ObjectMapper mapper;
    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(FullyQualifiedVBAModuleId.class,
                new FullyQualifiedVBAModuleIdSerializer());
        mapper.registerModule(module);
    }

    public FullyQualifiedVBAModuleId(SensibleWorkbook workbook, VBAModule module) {
        this.workbook = workbook;
        this.module = module;
    }

    public SensibleWorkbook getWorkbook() { return workbook; }

    public String getWorkbookId() { return workbook.getId(); }

    public VBAModule getModule() { return module; }

    public String getModuleName() { return module.getName(); }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        FullyQualifiedVBAModuleId that = (FullyQualifiedVBAModuleId) o;
        if (!workbook.equals(that.workbook)) return false;
        return module.equals(that.module);
    }

    @Override
    public int hashCode() {
        int result = workbook.hashCode();
        result = 31 * result + module.hashCode();
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

    public String toJson() throws JsonProcessingException {
        // no indent
        return mapper.writeValueAsString(this);
    }

    @Override
    public int compareTo(FullyQualifiedVBAModuleId other) {
        int workbookIdComparison =
                getWorkbookId().compareTo(other.getWorkbookId());
        if (workbookIdComparison == 0) {
            return getModuleName().compareTo(other.getModuleName());
        } else {
            return workbookIdComparison;
        }
    }

    public static class FullyQualifiedVBAModuleIdSerializer extends StdSerializer<FullyQualifiedVBAModuleId> {
        public FullyQualifiedVBAModuleIdSerializer() { this(null); }
        public FullyQualifiedVBAModuleIdSerializer(Class<FullyQualifiedVBAModuleId> t) { super(t); }
        @Override
        public void serialize(FullyQualifiedVBAModuleId fqmi, JsonGenerator jgen, SerializerProvider provider) throws IOException {
            jgen.writeStartObject();
            jgen.writeStringField("workbookId", fqmi.getWorkbookId());
            jgen.writeStringField("moduleName", fqmi.getModuleName());
            jgen.writeEndObject();

        }
    }
}

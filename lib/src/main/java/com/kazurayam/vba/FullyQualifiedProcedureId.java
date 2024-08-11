package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;

public class FullyQualifiedProcedureId implements Comparable<FullyQualifiedProcedureId> {

    private final SensibleWorkbook workbook;
    private final VBAModule module;
    private final VBAProcedure procedure;

    private final static ObjectMapper mapper;
    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(FullyQualifiedProcedureId.class,
                new FullyQualifiedProcedureIdSerializer());
        mapper.registerModule(module);
    }

    public FullyQualifiedProcedureId(SensibleWorkbook workbook, VBAModule module, VBAProcedure procedure) {
        this.workbook = workbook;
        this.module = module;
        this.procedure = procedure;
    }

    public SensibleWorkbook getWorkbook() {
        return workbook;
    }

    public String getWorkbookId() {
        return workbook.getId();
    }

    public VBAModule getModule() {
        return module;
    }

    public String getModuleName() {
        return module.getName();
    }

    public VBAProcedure getProcedure() {
        return procedure;
    }

    public String getProcedureName() {
        return procedure.getName();
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        FullyQualifiedProcedureId that = (FullyQualifiedProcedureId) o;
        if (!getWorkbookId().equals(that.getWorkbookId())) return false;
        if (!getModuleName().equals(that.getModuleName())) return false;
        return getProcedureName().equals(that.getProcedureName());
    }

    @Override
    public int hashCode() {
        int result = getWorkbookId().hashCode();
        result = 31 * result + getModuleName().hashCode();
        result = 31 * result + getProcedureName().hashCode();
        return result;
    }

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
        // no indentations
        return mapper.writeValueAsString(this);
    }

    @Override
    public int compareTo(FullyQualifiedProcedureId that) {
        int workbookIdComparison = getWorkbookId().compareTo(that.getWorkbookId());
        if (workbookIdComparison == 0) {
            int moduleNameComparison = getModuleName().compareTo(that.getModuleName());
            if (moduleNameComparison == 0) {
                return getProcedureName().compareTo(that.getProcedureName());
            } else {
                return moduleNameComparison;
            }
        } else {
            return workbookIdComparison;
        }
    }

    /**
     *
     */
    public static class FullyQualifiedProcedureIdSerializer extends StdSerializer<FullyQualifiedProcedureId> {
        public FullyQualifiedProcedureIdSerializer() { this(null); }
        public FullyQualifiedProcedureIdSerializer(Class<FullyQualifiedProcedureId> t) { super(t); }
        @Override
        public void serialize(
                FullyQualifiedProcedureId fqpi, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject();
            jgen.writeStringField("workbookId", fqpi.getWorkbookId());
            jgen.writeStringField("moduleName", fqpi.getModuleName());
            jgen.writeStringField("procedureName", fqpi.getProcedureName());
            jgen.writeEndObject();
        }
    }
}


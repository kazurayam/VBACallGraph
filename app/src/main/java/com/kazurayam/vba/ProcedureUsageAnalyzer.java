package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ProcedureUsageAnalyzer {

    private List<SensibleWorkbook> workbooks;

    private static final ObjectMapper mapper;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(ProcedureUsageAnalyzer.class,
                new ProcedureUsageAnalyzerSerializer());
        mapper.registerModule(module);
    }

    public ProcedureUsageAnalyzer() {
        workbooks = new ArrayList<>();
    }

    public void add(SensibleWorkbook workbook) {
        this.workbooks.add(workbook);
    }

    public Iterator<SensibleWorkbook> iterator() {
        return workbooks.iterator();
    }

    public SensibleWorkbook get(int index) {
        return workbooks.get(index);
    }

    public int size() {
        return workbooks.size();
    }

    @Override
    public String toString() {
        //pretty printed
        try {
            Object json = mapper.readValue(this.toJson(), Object.class);
            return mapper.writerWithDefaultPrettyPrinter().writeValueAsString(json);
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
    }

    public String toJson() throws JsonProcessingException {
        // no indentation
        return mapper.writeValueAsString(this);
    }

    private static class ProcedureUsageAnalyzerSerializer extends StdSerializer<ProcedureUsageAnalyzer> {
        public ProcedureUsageAnalyzerSerializer() { this(null); }
        public ProcedureUsageAnalyzerSerializer(Class<ProcedureUsageAnalyzer> t) {
            super(t);
        }
        @Override
        public void serialize(
                ProcedureUsageAnalyzer domain, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject();
            jgen.writeArrayFieldStart("workbooks");
            Iterator<SensibleWorkbook> iter = domain.iterator();
            while(iter.hasNext()) {
                SensibleWorkbook wb = iter.next();
                jgen.writeObject(wb);
            }
            jgen.writeEndArray();
            jgen.writeEndObject();
        }
    }
}

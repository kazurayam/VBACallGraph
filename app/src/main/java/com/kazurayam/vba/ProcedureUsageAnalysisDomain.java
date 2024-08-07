package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ProcedureUsageAnalysisDomain {

    private List<Workbook> workbooks;

    private static final ObjectMapper mapper;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(ProcedureUsageAnalysisDomain.class,
                new ProcedureUsageAnalysisDomainSerializer());
        mapper.registerModule(module);
    }

    public ProcedureUsageAnalysisDomain() {
        workbooks = new ArrayList<>();
    }

    public void add(Workbook workbook) {
        this.workbooks.add(workbook);
    }

    public Iterator<Workbook> iterator() {
        return workbooks.iterator();
    }

    public Workbook get(int index) {
        return workbooks.get(index);
    }

    public int size() {
        return workbooks.size();
    }



    private static class ProcedureUsageAnalysisDomainSerializer extends StdSerializer<ProcedureUsageAnalysisDomain> {
        public ProcedureUsageAnalysisDomainSerializer() { this(null); }
        public ProcedureUsageAnalysisDomainSerializer(Class<ProcedureUsageAnalysisDomain> t) {
            super(t);
        }
        @Override
        public void serialize(
                ProcedureUsageAnalysisDomain domain, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeArrayFieldStart("workbooks");
            Iterator<Workbook> iter = domain.iterator();
            while(iter.hasNext()) {
                Workbook wb = iter.next();
                jgen.writeObject(wb);
            }
            jgen.writeEndArray();
        }
    }
}

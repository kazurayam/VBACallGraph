package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class FindUsagesApp {

    private final List<SensibleWorkbook> workbooks;

    private static final ObjectMapper mapper;

    private boolean EXCLUDE_UNITTEST_MODULES;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(FindUsagesApp.class,
                new FindUsageAppSerializer());
        module.addSerializer(SensibleWorkbook.class,
                new SensibleWorkbook.SensibleWorkbookSerializer());
        mapper.registerModule(module);
    }

    public FindUsagesApp() {
        workbooks = new ArrayList<>();
        EXCLUDE_UNITTEST_MODULES = true;
    }

    public void setExcludeUnittestModules(boolean excludeUnittestModules) {
        this.EXCLUDE_UNITTEST_MODULES = excludeUnittestModules;
    }

    public void add(SensibleWorkbook workbook) {
        this.workbooks.add(workbook);
    }

    public SensibleWorkbook get(int index) {
        return workbooks.get(index);
    }

    public Iterator<SensibleWorkbook> iterator() {
        return workbooks.iterator();
    }

    public int size() {
        return workbooks.size();
    }

    public void writeDiagram(Path pu) throws IOException {
        BufferedWriter bw = Files.newBufferedWriter(pu);
        writeDiagram(bw);
    }

    public void writeDiagram(Writer writer) throws IOException {
        ProcedureUsageDiagramGenerator pudgen =
                new ProcedureUsageDiagramGenerator();
        pudgen.writeStartUml();
        for (SensibleWorkbook wb : workbooks) {
            pudgen.writeStartWorkbook(wb);
            for (String key : wb.getModules().keySet()) {
                VBAModule module = wb.getModule(key);
                if (!shouldIgnore(module)) {
                    pudgen.writeStartModule(module);
                    for (VBAProcedure procedure : module.getProcedures()) {
                        pudgen.writeProcedure(module, procedure);
                    }
                    pudgen.writeEndModule();
                }
            }
            pudgen.writeEndWorkbook();
        }
        pudgen.writeEndUml();
        writer.write(pudgen.toString());
        writer.flush();
        writer.close();
    }

    private boolean shouldIgnore(VBAModule module) {
        if (EXCLUDE_UNITTEST_MODULES) {
            String moduleNameLC = module.getName().toLowerCase();
            return (moduleNameLC.startsWith("test") || moduleNameLC.endsWith("test"));
        }
        return false;
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

    private static class FindUsageAppSerializer extends StdSerializer<FindUsagesApp> {
        public FindUsageAppSerializer() { this(null); }
        public FindUsageAppSerializer(Class<FindUsagesApp> t) {
            super(t);
        }
        @Override
        public void serialize(
                FindUsagesApp app, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject();                             //{
            jgen.writeFieldName("VBAProcedureUsageAnalyzer"); //"VBAProcedureUsageAnalyzer":
            jgen.writeStartObject();                             //  {
            jgen.writeArrayFieldStart("workbooks");     //    "workbooks": [
            Iterator<SensibleWorkbook> iter = app.iterator();
            while(iter.hasNext()) {
                SensibleWorkbook wb = iter.next();
                jgen.writeObject(wb);                            //      { ... },
            }
            jgen.writeEndArray();                                //    ]
            jgen.writeEndObject();                               //  }
            jgen.writeEndObject();                               //}
        }
    }
}

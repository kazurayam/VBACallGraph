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
import java.util.SortedSet;

public class FindUsagesApp {

    private final List<SensibleWorkbook> workbooks;
    private final Indexer indexer;
    private Options options = Options.DEFAULT;

    private static final ObjectMapper mapper;

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
        indexer = new Indexer();
        options = Options.DEFAULT;
    }

    public void setOptions(Options options) {
        this.options = options;
    }

    public void add(SensibleWorkbook workbook) {
        workbooks.add(workbook);
        indexer.add(workbook);
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
        // build the index
        SortedSet<VBAProcedureReference> memo = indexer.findAllProcedureReferences();
        //
        ProcedureUsageDiagramGenerator pudgen =
                new ProcedureUsageDiagramGenerator();
        pudgen.writeStartUml();
        for (SensibleWorkbook wb : workbooks) {
            pudgen.writeStartWorkbook(wb);
            for (String key : wb.getModules().keySet()) {
                VBAModule module = wb.getModule(key);
                if (!options.shouldExclude(module)) {
                    pudgen.writeStartModule(module);
                    for (VBAProcedure procedure : module.getProcedures()) {
                        pudgen.writeProcedure(module, procedure);
                    }
                    pudgen.writeEndModule();
                }
            }
            pudgen.writeEndWorkbook();
        }

        // write the directed arrows between Modules
        SortedSet<VBAProcedureReference> moduleReferences = indexer.findAllProcedureReferences();
        for (VBAProcedureReference reference : moduleReferences) {
            if (!options.shouldExclude(reference)) {
                pudgen.writeProcedureReference(reference);
            }
        }

        //
        pudgen.writeEndUml();
        //
        writer.write(pudgen.toString());
        writer.flush();
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

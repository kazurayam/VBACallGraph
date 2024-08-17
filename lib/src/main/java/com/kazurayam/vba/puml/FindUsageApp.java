package com.kazurayam.vba.puml;

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

public class FindUsageApp {

    private final List<ModelWorkbook> workbooks;
    private final Indexer indexer;
    private Options options = Options.DEFAULT;

    private static final ObjectMapper mapper;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(FindUsageApp.class,
                new FindUsageAppSerializer());
        module.addSerializer(ModelWorkbook.class,
                new ModelWorkbook.ModelWorkbookSerializer());
        mapper.registerModule(module);
    }

    public FindUsageApp() {
        workbooks = new ArrayList<>();
        indexer = new Indexer();
        options = Options.DEFAULT;
    }

    public void setOptions(Options options) {
        this.options = options;
    }

    public void add(ModelWorkbook workbook) {
        workbooks.add(workbook);
        indexer.add(workbook);
    }

    public ModelWorkbook get(int index) {
        return workbooks.get(index);
    }

    public Iterator<ModelWorkbook> iterator() {
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
        indexer.setOptions(this.options);
        SortedSet<VBAProcedureReference> memo = indexer.findAllProcedureReferences();
        //
        FindUsageDiagramGenerator pudgen =
                new FindUsageDiagramGenerator();
        pudgen.writeStartUml();
        for (ModelWorkbook wb : workbooks) {
            pudgen.writeStartWorkbook(wb);
            for (String key : wb.getModules().keySet()) {
                VBAModule module = wb.getModule(key);
                if (!options.shouldExcludeModule(module)) {
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
            if (!options.shouldExcludeModule(reference)) {
                // we do not like to draw arrows between Module-A and Module-A
                // just to simplify the diagram
                if (!reference.isReferringToSameModule()) {
                    pudgen.writeProcedureReference(reference);
                }
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

    private static class FindUsageAppSerializer extends StdSerializer<FindUsageApp> {
        public FindUsageAppSerializer() { this(null); }
        public FindUsageAppSerializer(Class<FindUsageApp> t) {
            super(t);
        }
        @Override
        public void serialize(
                FindUsageApp app, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject();                             //{
            jgen.writeFieldName("VBAProcedureUsageAnalyzer"); //"VBAProcedureUsageAnalyzer":
            jgen.writeStartObject();                             //  {
            jgen.writeArrayFieldStart("workbooks");     //    "workbooks": [
            Iterator<ModelWorkbook> iter = app.iterator();
            while(iter.hasNext()) {
                ModelWorkbook wb = iter.next();
                jgen.writeObject(wb);                            //      { ... },
            }
            jgen.writeEndArray();                                //    ]
            jgen.writeEndObject();                               //  }
            jgen.writeEndObject();                               //}
        }
    }
}

package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.SortedMap;
import java.util.TreeMap;

public class VBAModule implements Comparable<VBAModule> {

    private final String name;
    private final List<VBAProcedure> procedureList;
    private final SortedMap<String, VBASource> vbaSources;

    private final static ObjectMapper mapper;
    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(VBAModule.class, new VBAModuleSerializer());
        module.addSerializer(VBAProcedure.class, new VBAProcedure.VBAProcedureSerializer());
        mapper.registerModule(module);
    }

    public VBAModule(String name) {
        this.name = name;
        this.procedureList = new ArrayList<>();
        this.vbaSources = new TreeMap<>();
    }

    public String getName() {
        return name;
    }

    public void add(VBAProcedure procedure) {
        procedureList.add(procedure);
    }

    public List<VBAProcedure> getProcedureList() {
        return procedureList;
    }

    @Override
    public boolean equals(Object obj) {
        if (!(obj instanceof VBAModule other)) {
            return false;
        }
        return this.getName() == other.getName();
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
        // no indentations
        return mapper.writeValueAsString(this);
    }

    @Override
    public int compareTo(VBAModule other) {
        return this.getName().compareTo(other.getName());
    }


    public static class VBAModuleSerializer extends StdSerializer<VBAModule> {
        public VBAModuleSerializer() {
            this(null);
        }
        public VBAModuleSerializer(Class<VBAModule> t) {
            super(t);
        }

        @Override
        public void serialize(
                VBAModule module, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject(); // {
            jgen.writeStringField("module", module.getName()); // "module": "NAME",
            // creating JSON Array with key
            jgen.writeFieldName("procedures");  // "procedures":
            // creating Array of Procedure objects
            jgen.writeStartArray();  // [
            //
            List<VBAProcedure> list = module.getProcedureList();
            for (VBAProcedure procedure : list) {
                jgen.writeObject(procedure);
            }
            jgen.writeEndArray();    // ]
            jgen.writeEndObject();   // }
        }
    }
}

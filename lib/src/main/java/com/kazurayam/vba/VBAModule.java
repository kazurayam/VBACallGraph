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
import java.util.Objects;

public class VBAModule implements Comparable<VBAModule> {

    private final String name;
    private final ModuleType type;
    private final List<VBAProcedure> procedures;

    private VBASource vbaSource;

    private final static ObjectMapper mapper;
    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(VBAModule.class, new VBAModuleSerializer());
        module.addSerializer(VBAProcedure.class, new VBAProcedure.VBAProcedureSerializer());
        module.addSerializer(VBASource.class, new VBASource.VBASourceSerializer());
        mapper.registerModule(module);
    }

    public VBAModule(String name, ModuleType type) {
        this.name = name;
        this.type = type;
        this.procedures = new ArrayList<>();
        this.vbaSource = null;
    }

    public String getName() {
        return name;
    }

    public ModuleType getType() {
        return type;
    }

    public boolean isClass() {
        return type.equals(ModuleType.Class);
    }

    public boolean isStandard() {
        return type.equals(ModuleType.Standard);
    }

    public void add(VBAProcedure procedure) {
        procedures.add(procedure);
    }

    public List<VBAProcedure> getProcedures() {
        return procedures;
    }

    public boolean hasProcedure(String procedureName) {
        VBAProcedure procedure = this.getProcedure(procedureName);
        return procedure != null;
    }

    public VBAProcedure getProcedure(String procedureName) {
        for (VBAProcedure procedure : procedures) {
            if (procedure.getName().equals(procedureName)) {
                return procedure;
            }
        }
        return null;
    }

    public void setVBASource(VBASource vbaSource) {
        this.vbaSource = vbaSource;
    }
    public VBASource getVBASource() {
      return vbaSource;
    }

    @Override
    public boolean equals(Object obj) {
        if (!(obj instanceof VBAModule other)) {
            return false;
        }
        return Objects.equals(this.getName(), other.getName());
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

    public static enum ModuleType {
        Standard(".bas"),
        Class(".cls"),
        Unspecified(".unspecified");
        private String extension;
        private ModuleType(String extension) {
            this.extension = extension;
        }
        public String getFileExtension() {
            return extension;
        }
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
            jgen.writeStringField("type", module.getType().toString());
            // creating JSON Array with key
            jgen.writeFieldName("procedures");  // "procedures":
            // creating Array of Procedure objects
            jgen.writeStartArray();  // [
            //
            List<VBAProcedure> list = module.getProcedures();
            for (VBAProcedure procedure : list) {
                jgen.writeObject(procedure);
            }
            jgen.writeEndArray();    // ]
            jgen.writeObjectField("source", module.getVBASource());
            jgen.writeEndObject();   // }
        }
    }
}

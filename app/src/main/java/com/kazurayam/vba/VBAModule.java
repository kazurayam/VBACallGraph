package com.kazurayam.vba;
import com.fasterxml.jackson.core.JsonGenerator;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;

public class VBAModule implements Comparable<VBAModule> {

    private final static ObjectMapper mapper;
    private final String name;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(VBAModule.class, new VBAModuleSerializer());
        mapper.registerModule(module);
    }

    public VBAModule(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }
    @Override
    public boolean equals(Object obj) {
        if (!(obj instanceof VBAModule)) {
            return false;
        }
        VBAModule other = (VBAModule) obj;
        return this.getName() == other.getName();
    }

    @Override
    public String toString() {
        return this.getName();
    }

    public String toJson() throws JsonProcessingException {
        return mapper.writeValueAsString(this);
    }
    @Override
    public int compareTo(VBAModule other) {
        return this.getName().compareTo(other.getName());
    }


    private static class VBAModuleSerializer extends StdSerializer<VBAModule> {
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
            jgen.writeString(module.getName());
        }
    }
}

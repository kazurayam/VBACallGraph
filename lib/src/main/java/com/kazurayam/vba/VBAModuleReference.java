package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;

public class VBAModuleReference implements Comparable<VBAModuleReference> {

    private final FullyQualifiedVBAModuleId referrer;

    private final FullyQualifiedVBAModuleId referee;

    private final static ObjectMapper mapper;
    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(VBAModuleReference.class,
                new VBAModuleReferenceSerializer());
        module.addSerializer(FullyQualifiedVBAModuleId.class,
                new FullyQualifiedVBAModuleId.FullyQualifiedVBAModuleIdSerializer());
        mapper.registerModule(module);
    }

    public VBAModuleReference(FullyQualifiedVBAModuleId referrer,
                              FullyQualifiedVBAModuleId referee) {
        this.referrer = referrer;
        this.referee = referee;
    }

    public FullyQualifiedVBAModuleId getReferrer() {
        return referrer;
    }

    public FullyQualifiedVBAModuleId getReferee() {
        return referee;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        VBAModuleReference other = (VBAModuleReference) o;
        if (referrer.equals(other.referrer)) {
            return referee.equals(other.referee);
        } else {
            return false;
        }
    }

    @Override
    public int hashCode() {
        int result = referrer.hashCode();
        result = 31 * result + referee.hashCode();
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
        // no indentation
        return mapper.writeValueAsString(this);
    }

    @Override
    public int compareTo(VBAModuleReference o) {
        int referrerComparison = referrer.compareTo(o.referrer);
        if (referrerComparison == 0) {
            return referee.compareTo(o.referee);
        } else {
            return referrerComparison;
        }
    }

    public static class VBAModuleReferenceSerializer extends StdSerializer<VBAModuleReference> {
        public VBAModuleReferenceSerializer() { this(null); }
        public VBAModuleReferenceSerializer(Class<VBAModuleReference> t) { super(t); }
        @Override
        public void serialize(
                VBAModuleReference mr, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject();
            jgen.writeObjectField("referrer", mr.getReferrer());
            jgen.writeObjectField("referee", mr.getReferee());
            jgen.writeEndObject();
        }
    }
}

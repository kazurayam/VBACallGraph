package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;

public class ProcedureReference implements Comparable<ProcedureReference> {

    /**
     * The id of Procedure that refers the referee
     */
    private final FullyQualifiedProcedureId referrer;

    /**
     * The id of Procedure that is referred by the referrer
     */
    private final FullyQualifiedProcedureId referee;

    private final static ObjectMapper mapper;
    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(ProcedureReference.class,
                new ProcedureReferenceSerializer());
        module.addSerializer(FullyQualifiedProcedureId.class,
                new FullyQualifiedProcedureId.FullyQualifiedProcedureIdSerializer());
        mapper.registerModule(module);
    }

    public ProcedureReference(FullyQualifiedProcedureId referrer,
                              FullyQualifiedProcedureId referee) {
        this.referrer = referrer;
        this.referee = referee;
    }

    public FullyQualifiedProcedureId getReferrer() {
        return referrer;
    }

    public FullyQualifiedProcedureId getReferee() {
        return referee;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        ProcedureReference other = (ProcedureReference) o;
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
    public int compareTo(ProcedureReference other) {
        int referrerComparison = referrer.compareTo(other.referrer);
        if (referrerComparison == 0) {
            return referee.compareTo(other.referee);
        } else {
            return referrerComparison;
        }
    }

    public static class ProcedureReferenceSerializer extends StdSerializer<ProcedureReference> {
        public ProcedureReferenceSerializer() { this(null); }
        public ProcedureReferenceSerializer(Class<ProcedureReference> t) { super(t); }
        @Override
        public void serialize(
                ProcedureReference pr, JsonGenerator jgen, SerializerProvider provider)
            throws IOException {
            jgen.writeStartObject();
            jgen.writeObjectField("referrer", pr.getReferrer());
            jgen.writeObjectField("referee", pr.getReferee());
            jgen.writeEndObject();
        }
    }
}

package com.kazurayam.vba.puml;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;

public class VBAProcedureReference implements Comparable<VBAProcedureReference> {

    /**
     * The id of Module that refers the referee Procedure
     */
    private final FullyQualifiedVBAModuleId referrer;
    private final VBASource referrerSource;
    private final VBASourceLine referrerLine;

    /**
     * The id of Procedure that is referred by the referrer
     */
    private final FullyQualifiedVBAProcedureId referee;

    private final static ObjectMapper mapper;
    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(VBAProcedureReference.class,
                new VBAProcedureReferenceSerializer());
        module.addSerializer(FullyQualifiedVBAModuleId.class,
                new FullyQualifiedVBAModuleId.FullyQualifiedVBAModuleIdSerializer());
        module.addSerializer(FullyQualifiedVBAProcedureId.class,
                new FullyQualifiedVBAProcedureId.FullyQualifiedVBAProcedureIdSerializer());
        module.addSerializer(VBASource.class,
                new VBASource.VBASourceSerializer());
        module.addSerializer(VBASourceLine.class,
                new VBASourceLine.VBASourceLineSerializer());
        mapper.registerModule(module);
    }

    public VBAProcedureReference(FullyQualifiedVBAModuleId referrer,
                                 VBASource referrerSource,
                                 VBASourceLine referrerLine,
                                 FullyQualifiedVBAProcedureId referee) {
        this.referrer = referrer;
        this.referrerSource = referrerSource;
        this.referrerLine = referrerLine;
        this.referee = referee;
    }

    public FullyQualifiedVBAModuleId getReferrer() {
        return referrer;
    }

    public VBASource getReferrerSource() { return referrerSource; }

    public VBASourceLine getReferrerLine() { return referrerLine; }

    public FullyQualifiedVBAProcedureId getReferee() {
        return referee;
    }

    public Boolean isReferringToSameModule() {
        return (referee.getModule().equals(referrer.getModule()));
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        VBAProcedureReference other = (VBAProcedureReference) o;
        if (referrer.equals(other.referrer)) {
            if (referrerSource.equals(other.referrerSource)) {
                if (referrerLine.equals(other.referrerLine)) {
                    return referee.equals(other.referee);
                } else {
                    return false;
                }
            } else {
                return false;
            }
        } else {
            return false;
        }
    }

    @Override
    public int hashCode() {
        int result = referrer.hashCode();
        result = 31 * result + referrerSource.hashCode();
        result = 31 * result + referrerLine.hashCode();
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
    public int compareTo(VBAProcedureReference other) {
        int referrerComparison = referrer.compareTo(other.referrer);
        if (referrerComparison == 0) {
            int referrerSourceComparison =
                    referrerSource.compareTo(other.referrerSource);
            return referee.compareTo(other.referee);
        } else {
            return referrerComparison;
        }
    }

    public static class VBAProcedureReferenceSerializer extends StdSerializer<VBAProcedureReference> {
        public VBAProcedureReferenceSerializer() { this(null); }
        public VBAProcedureReferenceSerializer(Class<VBAProcedureReference> t) { super(t); }
        @Override
        public void serialize(
                VBAProcedureReference pr, JsonGenerator jgen, SerializerProvider provider)
            throws IOException {
            jgen.writeStartObject();
            jgen.writeObjectField("referrer", pr.getReferrer());
            jgen.writeObjectField("referrerSource", pr.getReferrerSource());
            jgen.writeObjectField("referrerLine", pr.getReferrerLine());
            jgen.writeObjectField("referee", pr.getReferee());
            jgen.writeEndObject();
        }
    }
}

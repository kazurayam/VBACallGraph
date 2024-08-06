package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;

public class VBAProcedureSerializer extends StdSerializer<VBAProcedure> {
    public VBAProcedureSerializer() {
        this(null);
    }
    public VBAProcedureSerializer(Class<VBAProcedure> t) {
        super(t);
    }
    @Override
    public void serialize(
            VBAProcedure proc, JsonGenerator jgen, SerializerProvider provider)
            throws IOException, JsonProcessingException {
        jgen.writeStartObject();
        jgen.writeStringField("name", proc.getName());
        jgen.writeStringField("module", proc.getModule());
        jgen.writeStringField("scope", proc.getScope().toString());
        jgen.writeStringField("subOrFunc", proc.getSubOrFunc().toString());
        jgen.writeNumberField("lineNo", proc.getLineNo());
        jgen.writeStringField("source", proc.getSource());
        jgen.writeStringField("comment", proc.getComment());
        jgen.writeEndObject();
    }
}

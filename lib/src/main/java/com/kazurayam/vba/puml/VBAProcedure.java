package com.kazurayam.vba.puml;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;

public class VBAProcedure {
    // |Project|ModuleType|Module|Scope|ProcKind|Procedure|LineNo|Source|Comment|
    private final String project;
    private final VBAModule.ModuleType moduleType;
    private final String module;
    private final Scope scope;
    private final ProcKind procKind;
    private final String procedure;
    private final int lineNo;
    private final String source;
    private final String comment;

    private final static ObjectMapper mapper;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(VBAProcedure.class, new VBAProcedureSerializer());
        mapper.registerModule(module);
    }

    private VBAProcedure(Builder builder) {
        project = builder.project;
        moduleType = builder.moduleType;
        module = builder.module;
        scope = builder.scope;
        procKind = builder.procKind;
        procedure = builder.procedure;
        lineNo = builder.lineNo;
        source = builder.source;
        comment = builder.comment;
    }
    public String getProject() { return project; }
    public VBAModule.ModuleType getModuleType() { return moduleType; }
    public String getModule() { return module; }
    public Scope getScope() { return scope; }
    public ProcKind getProcKind() { return procKind; }
    public String getProcedure() { return procedure; }
    public String getSourceFileName() {
        if (this.getModuleType().equals(VBAModule.ModuleType.Class)) {
            return getModule() + ".cls";
        } else if (this.getModuleType().equals(VBAModule.ModuleType.Standard)) {
            return getModule() + ".bas";
        } else {
            return getModule() + ".unknown";
        }
    }
    public int getLineNo() { return lineNo; }
    public String getSource() { return source; }
    public String getComment() { return comment; }

    @Override
    public String toString() {
        //pretty print
        try {
            Object json = mapper.readValue(this.toJson(), Object.class);
            return mapper.writerWithDefaultPrettyPrinter().writeValueAsString(json);
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
    }
    public String toJson() throws JsonProcessingException {
        // without indent
        return mapper.writeValueAsString(this);
    }

    /**
     *
     */
    public static class Builder {
        // |Project|ModuleType|Module|Scope|ProcKind|Procedure|LineNo|Source|Comment|
        private String project;
        private VBAModule.ModuleType moduleType;
        private String module;
        private Scope scope;
        private ProcKind procKind;
        private String procedure;
        private int lineNo;
        private String source;
        private String comment;
        public Builder() {
            project = "";
            moduleType = VBAModule.ModuleType.Unspecified;
            module = "";
            scope = Scope.Unspecified;
            procKind = ProcKind.Unspecified;
            procedure = "";
            lineNo = 0;
            source = "";
            comment = "";
        }
        public Builder project(String project) {
            this.project = project;
            return this;
        }
        public Builder moduleType(String moduleType) {
            try {
                this.moduleType = VBAModule.ModuleType.valueOf(moduleType);
            } catch (IllegalArgumentException e) {
                this.moduleType = VBAModule.ModuleType.Unspecified;
            }
            return this;
        }
        public Builder module(String module) {
            this.module = module;
            return this;
        }
        public Builder scope(String scope) {
            try {
                this.scope = VBAProcedure.Scope.valueOf(scope);
            } catch (IllegalArgumentException e) {
                this.scope = VBAProcedure.Scope.Unspecified;;
            }
            return this;
        }
        public Builder procKind(String procKind) {
            switch (procKind) {
                case "Sub":
                    this.procKind = ProcKind.Sub;
                    break;
                case "Function":
                    this.procKind = ProcKind.Function;
                    break;
                case "Property Let":
                    this.procKind = ProcKind.PropertyLet;
                    break;
                case "Property Get":
                    this.procKind = ProcKind.PropertyGet;
                    break;
                case "Property Set":
                    this.procKind = ProcKind.PropertySet;
                    break;
                default:
                    this.procKind = ProcKind.Unspecified;
                    break;
            }
            return this;
        }
        public Builder procedure(String procedure) {
            this.procedure = procedure;
            return this;
        }
        public Builder lineNo(int lineNo) {
            this.lineNo = lineNo;
            return this;
        }
        public Builder source(String source) {
            this.source = source;
            return this;
        }
        public Builder comment(String comment) {
            this.comment = comment;
            return this;
        }
        public VBAProcedure build() {
            return new VBAProcedure(this);
        }
    }


    /**
     *
     */
    public static class VBAProcedureSerializer extends StdSerializer<VBAProcedure> {
        public VBAProcedureSerializer() {
            this(null);
        }
        public VBAProcedureSerializer(Class<VBAProcedure> t) {
            super(t);
        }
        @Override
        public void serialize(
                VBAProcedure proc, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject();
            jgen.writeStringField("Project", proc.getProject());
            jgen.writeStringField("ModuleType", proc.getModuleType().name());
            jgen.writeStringField("Module", proc.getModule());
            jgen.writeStringField("Scope", proc.getScope().toString());
            jgen.writeStringField("ProcKind", proc.getProcKind().toString());
            jgen.writeStringField("Procedure", proc.getProcedure());
            jgen.writeNumberField("LineNo", proc.getLineNo());
            jgen.writeStringField("Source", proc.getSource());
            jgen.writeStringField("Comment", proc.getComment());
            jgen.writeEndObject();
        }
    }

    public static enum Scope {
        Public,
        Private,
        Friend,
        Static,
        Unspecified;
    }

    public static enum ProcKind {
        Sub,
        Function,
        PropertyLet,
        PropertyGet,
        PropertySet,
        Unspecified;
    }

}

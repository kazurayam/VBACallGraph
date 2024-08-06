package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.module.SimpleModule;

public class VBAProcedure {
    private String name;
    private String module;
    private Scope scope;
    private SubOrFunc subOrFunc;
    private int lineNo;
    private String source;
    private String comment;

    private static ObjectMapper mapper;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(VBAProcedure.class, new VBAProcedureSerializer());
        mapper.registerModule(module);
    }

    private VBAProcedure(Builder builder) {
        name = builder.name;
        module = builder.module;
        scope = builder.scope;
        subOrFunc = builder.subOrFunc;
        lineNo = builder.lineNo;
        source = builder.source;
        comment = builder.comment;

    }
    public String getName() { return name; }
    public String getModule() { return module; }
    public Scope getScope() { return scope; }
    public SubOrFunc getSubOrFunc() { return subOrFunc; }
    public int getLineNo() { return lineNo; }
    public String getSource() { return source; }
    public String getComment() { return comment; }
    @Override
    public String toString() {
        try {
            return this.toJson();
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
    }
    public String toJson() throws JsonProcessingException {
        return mapper.writeValueAsString(this);
    }

    public static class Builder {
        private String name;
        private String module;
        private Scope scope;
        private SubOrFunc subOrFunc;
        private int lineNo;
        private String source;
        private String comment;
        public Builder() {
            name = "";
            module = "";
            scope = Scope.Unspecified;
            subOrFunc = SubOrFunc.Unspecified;
            lineNo = 0;
            source = "";
            comment = "";
        }
        public Builder name(String name) {
            this.name = name;
            return this;
        }
        public Builder module(String module) {
            this.module = module;
            return this;
        }
        public Builder scope(Scope scope) {
            this.scope = scope;
            return this;
        }
        public Builder subOrFunc(SubOrFunc subOrFunc) {
            this.subOrFunc = subOrFunc;
            return this;
        };
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

}

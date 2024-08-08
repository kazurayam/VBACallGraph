package com.kazurayam.vbaexample;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;
import java.nio.file.Path;

public enum MyWorkbook {
    Backbone("Backboneライブラリ",
            "kazurayam-vba-lib",
            "office/kazurayam-vba-lib.xlsm",
            "office/exported-vba-source/kazurayam-vba-lib"),
    Member("Member会員名簿のためのVBAライブラリ",
            "aogan-vba-lib",
            "office/Member会員名簿のためのVBAライブラリ.xlsm",
            "office/exported-vba-source/Member会員名簿のためのVBAライブラリ"),
    Cashbook("Cashbook現金出納帳のためのVBAライブラリ",
            "aogan-vba-lib",
            "office/Cashbook現金出納帳のためのVBAライブラリ.xlsm",
            "office/exported-vba-source/Cashbook現金出納帳のためのVBAライブラリ"),
    Settlement("決算算出ワークブック",
            "aogan-jimukyoku",
            "office/決算算出ワークブック_令和5年度.xlsm",
            "office/exported-vba-source/決算算出ワークブック_令和5年度"),
    FeePaymentCheck("会費納入状況チェック",
            "aogan-jimukyoku",
            "office/会費納入状況チェック_R6年度.xlsm",
            "office/exported-vba-source/会費納入状況チェック_R6年度"),
    PleasePayFeeLetter("会費納入をお願いするletterを作成する",
            "aogan-jimukyoku",
            "office/会費納入のお願いletterを作成する.xlsm",
            "office/exported-vba-source/会費納入のお願いletterを作成する"),
    WebCredentials("会員名簿からIDパスワード管理情報を生成する",
            "aogan-jimukyoku",
            "office/会員名簿からIDパスワード管理情報を生成する.xlsm",
            "office/exported-vba-source/会員名簿からIDパスワード管理情報を生成する");

    private final String id;
    private final String repositoryName;
    private final String workbookSubPath;
    private final String sourceDirSubPath;

    private static final ObjectMapper mapper;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(MyWorkbook.class, new WorkbookInstanceLocationSerializer());
        mapper.registerModule(module);
    }

    MyWorkbook(String id, String repositoryName,
               String workbookSubPath,
               String vbaSourceDirSubPath) {
        this.id = id;
        this.repositoryName = repositoryName;
        this.workbookSubPath = workbookSubPath;
        this.sourceDirSubPath = vbaSourceDirSubPath;
    }

    public String getId() {
        return id;
    }

    public String getRepositoryName() {
        return repositoryName;
    }

    public String getWorkbookSubPath() {
        return workbookSubPath;
    }

    public String getSourceDirSubPath() {
        return sourceDirSubPath;
    }

    public Path resolveWorkbookUnder(Path baseDir) {
        return baseDir.resolve(getRepositoryName()).resolve(getWorkbookSubPath());
    }

    public Path resolveSourceDirUnder(Path baseDir) {
        return baseDir.resolve(getRepositoryName()).resolve(getSourceDirSubPath());
    }

    @Override
    public String toString() {
        // pretty print
        try {
            Object json = mapper.readValue(this.toJson(), Object.class);
            return mapper.writerWithDefaultPrettyPrinter().writeValueAsString(json);
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
    }

    public String toJson() throws JsonProcessingException {
        return mapper.writeValueAsString(this);
    }

    /**
     * Serializer of WorkbookInstanceLocation into Json based on the Jackson Databind
     */
    private static class WorkbookInstanceLocationSerializer
            extends StdSerializer<MyWorkbook> {
        public WorkbookInstanceLocationSerializer() {
            this(null);
        }

        public WorkbookInstanceLocationSerializer(Class<MyWorkbook> t) {
            super(t);
        }

        @Override
        public void serialize(
                MyWorkbook wil, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject();
            jgen.writeStringField("id", wil.getId());
            jgen.writeStringField("repositoryName", wil.getRepositoryName());
            jgen.writeStringField("workbookSubPath", wil.getWorkbookSubPath());
            jgen.writeStringField("sourceDirSubPath", wil.getSourceDirSubPath());
            jgen.writeEndObject();
        }
    }
}

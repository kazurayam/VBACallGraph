package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;
import java.nio.file.Path;

public enum WorkbookInstanceLocation {
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
            "aomori-gankaikai-jimukyoku",
            "office/決算算出ワークブック_令和5年度.xlsm",
            "office/exported-vba-source/決算算出ワークブック_令和5年度"),
    FeePaymentCheck("会費納入状況チェック",
            "aomori-gankaikai-jimukyoku",
            "office/会費納入状況チェック_青森県眼科医会_R6年度.xlsm",
            "office/exported-vba-source/会費納入状況チェック_青森県眼科医会_R6年度"),
    PleasePayFeeLetter("会費納入をお願いするletterを作成する",
            "aomori-gankaikai-jimukyoku",
            "office/会費納入のお願いletterを作成する.xlsm",
            "office/exported-vba-source/会費納入のお願いletterを作成する"),
    WebCredentials("会員名簿からIDパスワード管理情報を生成する",
            "aomori-gankaikai-jimukyoku",
            "office/会員名簿からIDパスワード管理情報を生成する.xlsm",
            "office/exported-vba-source/会員名簿からIDパスワード管理情報を生成する");

    private final String id;
    private final String repositoryName;
    private final String workbookSubPath;
    private final String vbaSourceDirSubPath;

    private static final ObjectMapper mapper;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(WorkbookInstanceLocation.class, new WorkbookInstanceLocationSerializer());
        mapper.registerModule(module);
    }

    WorkbookInstanceLocation(String id, String repositoryName,
                                     String workbookSubPath,
                                     String vbaSourceDirSubPath) {
        this.id = id;
        this.repositoryName = repositoryName;
        this.workbookSubPath = workbookSubPath;
        this.vbaSourceDirSubPath = vbaSourceDirSubPath;
    }
    public String getId() {
        return id;
    }
    public String getRepositoryName() {
        return repositoryName;
    }
    public String getWorkbookSubPath() { return workbookSubPath; }
    public String getVbaSourceDirSubPath() {
        return vbaSourceDirSubPath;
    }
    public Path resolveWorkbookBasedOn(Path baseDir) {
        return baseDir.resolve(getRepositoryName()).resolve(getWorkbookSubPath());
    }
    public Path resolveVBASourceDirBasedOn(Path baseDir) {
        return baseDir.resolve(getRepositoryName()).resolve(getVbaSourceDirSubPath());
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
            extends StdSerializer<WorkbookInstanceLocation> {
        public WorkbookInstanceLocationSerializer() { this(null); }
        public WorkbookInstanceLocationSerializer(Class<WorkbookInstanceLocation> t) { super(t); }
        @Override
        public void serialize(
                WorkbookInstanceLocation wil, JsonGenerator jgen, SerializerProvider provider)
            throws IOException {
            jgen.writeStartObject();
            jgen.writeStringField("id", wil.getId());
            jgen.writeStringField("repositoryName", wil.getRepositoryName());
            jgen.writeStringField("workbookSubpath", wil.getWorkbookSubPath());
            jgen.writeStringField("vbaSourceDirSubPath", wil.getVbaSourceDirSubPath());
            jgen.writeEndObject();
        }
    }
}

package com.kazurayam.vba.example;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

public enum MyWorkbook {
    VBACallGraphSetup("VBACallGraphSetup",
            "src/test/fixture/hub-kazurayam/VBACallGraphSetup",
            "office/VBACallGraphSetup.xlsm",
            "office/exported-vba-source/VBACallGraphSetup"),

    Backbone("Backboneライブラリ",
            "src/test/fixture/hub-aogan/kazurayam-vba-lib",
            "office/Backbone.xlsm",
            "office/exported-vba-source/Backbone"),
    Member("Member会員名簿のためのVBAライブラリ",
            "src/test/fixture/hub-aogan/aogan-vba-lib",
            "office/Member会員名簿のためのVBAライブラリ.xlsm",
            "office/exported-vba-source/Member会員名簿のためのVBAライブラリ"),
    Cashbook("Cashbook現金出納帳のためのVBAライブラリ",
            "src/test/fixture/hub-aogan/aogan-vba-lib",
            "office/Cashbook現金出納帳のためのVBAライブラリ.xlsm",
            "office/exported-vba-source/Cashbook現金出納帳のためのVBAライブラリ"),
    Settlement("決算算出ワークブック",
            "src/test/fixture/hub-aogan/aogan-jimukyoku",
            "office/決算算出ワークブック_令和5年度.xlsm",
            "office/exported-vba-source/決算算出ワークブック_令和5年度"),
    FeePaymentCheck("会費納入状況チェック",
            "src/test/fixture/hub-aogan/aogan-jimukyoku",
            "office/会費納入状況チェック_R6年度.xlsm",
            "office/exported-vba-source/会費納入状況チェック_R6年度"),
    PleasePayFeeLetter("会費納入をお願いするletterを作成する",
            "src/test/fixture/hub-aogan/aogan-jimukyoku",
            "office/会費納入のお願いletterを作成する.xlsm",
            "office/exported-vba-source/会費納入のお願いletterを作成する"),
    WebCredentials("会員名簿からIDパスワード管理情報を生成する",
            "src/test/fixture/hub-aogan/aogan-jimukyoku",
            "office/会員名簿からIDパスワード管理情報を生成する.xlsm",
            "office/exported-vba-source/会員名簿からIDパスワード管理情報を生成する");

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(MyWorkbook.class).build();
    private static final Path PROJECT_DIR = too.getProjectDirectory();

    private final String id;
    private final String localRepository;
    private final String workbookSubPath;
    private final String sourceDirSubPath;

    private static final ObjectMapper mapper;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(MyWorkbook.class, new WorkbookInstanceLocationSerializer());
        mapper.registerModule(module);
    }

    MyWorkbook(String id,
               String localRepository,
               String workbookSubPath,
               String vbaSourceDirSubPath) {
        this.id = id;
        this.localRepository = localRepository;
        this.workbookSubPath = workbookSubPath;
        this.sourceDirSubPath = vbaSourceDirSubPath;
    }

    public String getId() {
        return id;
    }

    public String getLocalRepository() {
        return localRepository;
    }

    public String getWorkbookSubPath() {
        return workbookSubPath;
    }

    public String getSourceDirSubPath() {
        return sourceDirSubPath;
    }

    public Path resolveWorkbookUnder() {
        Path p = PROJECT_DIR.resolve(getLocalRepository()).resolve(getWorkbookSubPath());
        assert Files.exists(p) : p + " does not exist";
        return p;
    }

    public Path resolveSourceDirUnder() {
        Path p = PROJECT_DIR.resolve(getLocalRepository()).resolve(getSourceDirSubPath());
        assert Files.exists(p) : p + " does not exist";
        return p;
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
            jgen.writeStringField("localRepository", wil.getLocalRepository().toString());
            jgen.writeStringField("workbookSubPath", wil.getWorkbookSubPath());
            jgen.writeStringField("sourceDirSubPath", wil.getSourceDirSubPath());
            jgen.writeEndObject();
        }
    }
}

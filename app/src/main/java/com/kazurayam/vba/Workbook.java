package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.nio.file.Path;
import java.util.Map;
import java.util.SortedMap;
import java.util.TreeMap;

public class Workbook {

    private final String id;
    private final Path workbookPath;
    private final Path sourceDirPath;
    private final SortedMap<VBAModule, List<Procedure>> moduleProcedures;
    private static final String SHEET_NAME = "プロシージャ一覧";
    private final static ObjectMapper mapper;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(Workbook.class, new WorkbookSerializer());
        mapper.registerModule(module);
    }

    public Workbook(String id, Path workbookPath, Path sourceDirPath) throws IOException {
        this.id = id;
        this.workbookPath = workbookPath;
        this.sourceDirPath = sourceDirPath;
        InputStream is = Files.newInputStream(workbookPath);
        moduleProcedures = this.load(is);
    }

    public String getId() {
        return id;
    }

    public Path getWorkbookPath() {
        return workbookPath;
    }

    public Path getSourceDirPath() {
        return sourceDirPath;
    }

    public SortedMap<VBAModule, List<Procedure>> getModuleProcedures() {
        return this.moduleProcedures;
    }

    SortedMap<VBAModule, List<Procedure>> load(InputStream inputStream) throws IOException {
        SortedMap<VBAModule, List<Procedure>> moduleProcedures = new TreeMap<>();
        org.apache.poi.ss.usermodel.Workbook wb = new XSSFWorkbook(inputStream);
        Sheet sheet = wb.getSheet(SHEET_NAME);
        for (Row row : sheet) {
            if (row.getRowNum() > 0) { // we will skip the first row = header
                String name = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).getStringCellValue();
                String module = row.getCell(1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).getStringCellValue();
                String scope = row.getCell(2, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).getStringCellValue();
                String subOrFunc = row.getCell(3, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).getStringCellValue();
                double dvalue = row.getCell(4, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).getNumericCellValue();
                Integer lineNo = ((Double)dvalue).intValue();
                String source = row.getCell(5, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).getStringCellValue();
                String comment = row.getCell(6, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).getStringCellValue();
                Procedure proc =
                        new Procedure.Builder()
                                .name(name)
                                .module(module)
                                .scope(Procedure.Scope.valueOf(scope))
                                .subOrFunc(Procedure.SubOrFunc.valueOf(subOrFunc))
                                .source(source)
                                .comment(comment)
                                .build();
                VBAModule key = new VBAModule(module);
                //
                List<Procedure> list;
                if (moduleProcedures.containsKey(key)) {
                    list = moduleProcedures.get(key);
                } else {
                    list = new ArrayList<>();
                }
                list.add(proc);
                moduleProcedures.put(key, list);
            }
        }
        return moduleProcedures;
    }

    @Override
    public String toString() {
        //pretty print
        try {
            Object json = mapper.readValue(this.toJson(), Object.class);
            return mapper.writerWithDefaultPrettyPrinter().writeValueAsString(json);
        } catch (IOException e) {
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
    private static class WorkbookSerializer extends StdSerializer<Workbook> {
        public WorkbookSerializer() { this(null); }
        public WorkbookSerializer(Class<Workbook> t) { super(t); }
        @Override
        public void serialize(
                Workbook wb, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject();
            jgen.writeStringField("id", wb.getId());
            jgen.writeStringField("workbookPath", wb.getWorkbookPath().toString());
            jgen.writeStringField("sourceDirPath", wb.getSourceDirPath().toString());
            jgen.writeEndObject();
        }
    }
}

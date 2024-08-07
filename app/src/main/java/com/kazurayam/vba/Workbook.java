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
import java.util.List;
import java.nio.file.Path;

public class Workbook {

    private final Path baseDir;
    private final WorkbookInstanceLocation workbook;
    private final List<Procedure> procedureList;
    private static final String SHEET_NAME = "プロシージャ一覧";
    private final static ObjectMapper mapper;

    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(Workbook.class, new WorkbookSerializer());
        mapper.registerModule(module);
    }

    public Workbook(Path baseDir, WorkbookInstanceLocation workbook) throws IOException {
        this.baseDir = baseDir;
        this.workbook = workbook;
        Path xlsm = workbook.resolveWorkbookBasedOn(baseDir);
        InputStream is = Files.newInputStream(xlsm);
        procedureList = this.load(is);
    }

    public Path getBaseDir() {
        return baseDir;
    }

    public WorkbookInstanceLocation getWorkbook() {
        return workbook;
    }

    public List<Procedure> getCoppiedList() {
        return new ArrayList<>(procedureList);
    }

    List<Procedure> load(InputStream inputStream) throws IOException {
        List<Procedure> list = new ArrayList<>();
        org.apache.poi.ss.usermodel.Workbook wb = new XSSFWorkbook(inputStream);
        Sheet sheet = wb.getSheet(SHEET_NAME);
        int r = 0;
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
                list.add(proc);
            }
            r++;
        }
        return list;
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
            jgen.writeStringField("baseDir", wb.getBaseDir().toString());
            jgen.writeObjectField("workbook", wb.getWorkbook().toString());
            jgen.writeEndObject();
        }
    }
}

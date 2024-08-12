package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class ProcedureUsageDiagramGeneratorTest {

    private Logger logger = LoggerFactory.getLogger(ProcedureUsageDiagramGeneratorTest.class);

    private static TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(ProcedureUsageDiagramGeneratorTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(ProcedureUsageDiagramGeneratorTest.class)
                    .build();
    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");
    private SensibleWorkbook wb;
    private Path classOutputDir;
    private ProcedureUsageDiagramGenerator pudgen;

    @BeforeTest
    public void beforeTest() throws IOException {
        wb = new SensibleWorkbook(
                MyWorkbook.FeePaymentCheck.getId(),
                MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir));
        classOutputDir = too.cleanClassOutputDirectory();
    }

    @BeforeMethod
    public void setup() {
        pudgen = new ProcedureUsageDiagramGenerator();
    }

    @Test
    public void test_writeStartUml_writeEndUml() {
        pudgen.writeStartUml();
        pudgen.writeEndUml();
        assertThat(pudgen.toString()).contains("@startuml\n");
        assertThat(pudgen.toString()).contains("@enduml\n");
    }

    @Test
    public void test_writeStartWorkbook_writeEndWorkbook() {
        pudgen.writeStartWorkbook(wb);
        pudgen.writeEndWorkbook();
        logger.debug("[test_writeStartWorkbook_writeEndWorkbook] " +
                pudgen.toString());
        assertThat(pudgen.toString()).contains(
                "package 会費納入状況チェック {\n");
        assertThat(pudgen.toString()).contains(
                "}\n");
    }

    @Test
    public void test_writeStartModule_writeEndModule() {
        pudgen.writeStartModule(wb.getModule("年会費納入状況チェック"));
        pudgen.writeEndModule();
        logger.debug("[test_writeStartModule_writeEndModule] " +
                pudgen.toString());
        assertThat(pudgen.toString()).contains(
                "stereotype 年会費納入状況チェック {\n");
        assertThat(pudgen.toString()).contains(
                "}\n");
    }

    @Test
    public void test_writeProcedure() {
        VBAModule module = wb.getModule("年会費納入状況チェック");
        VBAProcedure procedure = module.getProcedure("FindPaymentBy");
        pudgen.writeProcedure(module, procedure);
        logger.debug("[test_writeProcedure] " +
                pudgen.toString());
        assertThat(pudgen.toString()).contains(
                "{method} FindPaymentBy\n");
    }

    @Test
    public void test_writeProcedureReference() {

    }

    @Test
    public void test_toString() throws IOException {
        Path output = classOutputDir.resolve("test_toString.pu");
        VBAModule module = wb.getModule("年会費納入状況チェック");
        VBAProcedure procedure = module.getProcedure("FindPaymentBy");
        pudgen.writeStartUml();
        pudgen.writeStartWorkbook(wb);
        pudgen.writeStartModule(module);
        pudgen.writeProcedure(module, module.getProcedure("FindPaymentBy"));
        pudgen.writeProcedure(module, module.getProcedure("Main"));
        pudgen.writeProcedure(module, module.getProcedure("OpenCashbook"));
        pudgen.writeProcedure(module, module.getProcedure("OpenMemberTable"));
        pudgen.writeProcedure(module, module.getProcedure("PrintFinding"));
        pudgen.writeProcedure(module, module.getProcedure("RecordFindingIntoMemberTable"));
        pudgen.writeEndModule();
        pudgen.writeEndWorkbook();
        pudgen.writeEndUml();
        pudgen.generate(output);
        assertThat(output).exists();
        assertThat(output.toFile().length()).isGreaterThan(0);
    }
}

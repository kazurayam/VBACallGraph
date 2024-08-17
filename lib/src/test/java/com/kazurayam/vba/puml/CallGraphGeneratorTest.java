package com.kazurayam.vba.puml;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.example.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class CallGraphGeneratorTest {

    private Logger logger = LoggerFactory.getLogger(CallGraphGeneratorTest.class);

    private static TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(CallGraphGeneratorTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(CallGraphGeneratorTest.class)
                    .build();
    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");
    private ModelWorkbook wbFeePaymentCheck;
    private ModelWorkbook wbCashbook;
    private Path classOutputDir;
    private CallGraphDiagramGenerator pudgen;

    @BeforeTest
    public void beforeTest() throws IOException {
        wbFeePaymentCheck = new ModelWorkbook(
                MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.FeePaymentCheck.getId());
        wbCashbook = new ModelWorkbook(
                MyWorkbook.Cashbook.resolveWorkbookUnder(baseDir),
                MyWorkbook.Cashbook.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Cashbook.getId());
        classOutputDir = too.cleanClassOutputDirectory();
    }

    @BeforeMethod
    public void setup() {
        pudgen = new CallGraphDiagramGenerator();
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
        pudgen.writeStartWorkbook(wbFeePaymentCheck);
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
        pudgen.writeStartModule(wbFeePaymentCheck.getModule("会費納入状況チェック"));
        pudgen.writeEndModule();
        logger.debug("[test_writeStartModule_writeEndModule] " +
                pudgen.toString());
        assertThat(pudgen.toString()).contains(
                "stereotype 会費納入状況チェック {\n");
        assertThat(pudgen.toString()).contains(
                "}\n");
    }

    @Test
    public void test_writeProcedure_as_private_method() {
        VBAModule module = wbFeePaymentCheck.getModule("会費納入状況チェック");
        VBAProcedure procedure = module.getProcedure("FindPaymentBy");
        pudgen.writeProcedure(module, procedure);
        logger.debug("[test_writeProcedure_as_method] " +
                pudgen.toString());
        assertThat(pudgen.toString()).contains(
                "{method} -FindPaymentBy\n");
    }

    @Test
    public void test_writeProcedure_as_public_field() {
        VBAModule module = wbCashbook.getModule("Account");
        VBAProcedure procedure = module.getProcedure("AccountName");
        pudgen.writeProcedure(module, procedure);
        logger.debug("[test_writeProcedure_as_field]" +
                pudgen.toString());
        assertThat(pudgen.toString()).contains(
                "{field} +AccountName\n");
    }

    @Test
    public void test_toString() throws IOException {
        Path output = classOutputDir.resolve("test_toString.puml");
        VBAModule module = wbFeePaymentCheck.getModule("会費納入状況チェック");
        VBAProcedure procedure = module.getProcedure("FindPaymentBy");
        pudgen.writeStartUml();
        pudgen.writeStartWorkbook(wbFeePaymentCheck);
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

package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
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

    private TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(ProcedureUsageDiagramGeneratorTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(ProcedureUsageDiagramGeneratorTest.class)
                    .build();
    private Path classOutputDir;
    private ProcedureUsageDiagramGenerator pudgen;

    @BeforeTest
    public void beforeTest() throws IOException {
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
        pudgen.writeStartWorkbook("会費納入状況チェック");
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
        pudgen.writeStartModule("会費納入状況チェック");
        pudgen.writeEndModule();
        logger.debug("[test_writeStartModule_writeEndModule] " +
                pudgen.toString());
        assertThat(pudgen.toString()).contains(
                "map 会費納入状況チェック {\n");
        assertThat(pudgen.toString()).contains(
                "}\n");
    }

    @Test
    public void test_writeProcedure() {
        pudgen.writeProcedure("FindPaymentBy");
        logger.debug("[test_writeProcedure] " +
                pudgen.toString());
        assertThat(pudgen.toString()).contains(
                "FindPaymentBy =>\n");
    }

    @Test
    public void test_toString() throws IOException {
        Path output = classOutputDir.resolve("test_toString.pu");
        pudgen.writeStartUml();
        pudgen.writeStartWorkbook("会費納入状況チェック");
        pudgen.writeStartModule("会費納入状況チェック");
        pudgen.writeProcedure("FindPaymentBy");
        pudgen.writeProcedure("Main");
        pudgen.writeProcedure("OpenCashbook");
        pudgen.writeProcedure("OpenMemberTable");
        pudgen.writeProcedure("PrintFinding");
        pudgen.writeProcedure("RecordFindingIntoMemberTable");
        pudgen.writeEndModule();
        pudgen.writeEndWorkbook();
        pudgen.writeEndUml();
        pudgen.save(output);
        assertThat(output).exists();
        assertThat(output.toFile().length()).isGreaterThan(0);
    }
}

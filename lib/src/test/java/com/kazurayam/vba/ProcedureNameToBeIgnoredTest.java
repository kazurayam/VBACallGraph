package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeTest;

import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class ProcedureNameToBeIgnoredTest {

    private static final Logger logger =
            LoggerFactory.getLogger(IndexerTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(ProcedureNameToBeIgnoredTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(ProcedureNameToBeIgnoredTest.class)
                    .build();

    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");
    private Path classOutputDir;

    private ProcedureNameToBeIgnored procedureNameToBeIgnored;
    private FullyQualifiedVBAProcedureId referee;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        SensibleWorkbook wbFeePaymentCheck =
                new SensibleWorkbook(
                        MyWorkbook.FeePaymentCheck.getId(),
                        MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                        MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir));
        VBAModule md = wbFeePaymentCheck.getModule("プロシージャー一覧を作る");
        VBAProcedure pr = md.getProcedure("プロシージャー一覧を作る");
        referee =
                new FullyQualifiedVBAProcedureId(wbFeePaymentCheck, md, pr);
    }

    @Test
    public void test_Class_Initialize() {
        ProcedureNameToBeIgnored entity = ProcedureNameToBeIgnored.Class_Initialize;
        assertThat(entity.getModuleType()).isEqualTo(VBAModule.ModuleType.Class);
        assertThat(entity.getProcedureName()).isEqualTo("Initialize");
    }

    @Test
    public void test_matches() {
        assertThat(ProcedureNameToBeIgnored.Class_Initialize
                .matches(referee)).isFalse();
        assertThat(ProcedureNameToBeIgnored.Class_Class_Initialize
                .matches(referee)).isFalse();
        assertThat(ProcedureNameToBeIgnored.Standard_プロシージャー一覧を作る
                .matches(referee)).isTrue();
    }
}

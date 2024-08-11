package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

public class ProcedureReferenceTest {
    private static final Logger logger = LoggerFactory.getLogger(ProcedureReferenceTest.class);
    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(ProcedureReferenceTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(ProcedureReferenceTest.class)
                    .build();
    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");

    private Path classOutputDir;
    private ProcedureReference procedureReference;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        //
        SensibleWorkbook wbFeePaymentCheck =
                new SensibleWorkbook(
                        MyWorkbook.FeePaymentCheck.getId(),
                        MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                        MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir));
        VBAModule md年会費納入状況チェック = wbFeePaymentCheck.getModule("年会費納入状況チェック");
        VBAProcedure procMain = md年会費納入状況チェック.getProcedure("Main");
        FullyQualifiedProcedureId referrer =
                new FullyQualifiedProcedureId(wbFeePaymentCheck, md年会費納入状況チェック, procMain);
        //
        SensibleWorkbook wbMember =
                new SensibleWorkbook(
                        MyWorkbook.Member.getId(),
                        MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                        MyWorkbook.Member.resolveSourceDirUnder(baseDir));
        VBAModule mdAoMemberUtils = wbMember.getModule("AoMemberUtils");
        VBAProcedure procFetchMemberTable = mdAoMemberUtils.getProcedure("FetchMemberTable");
        FullyQualifiedProcedureId referee =
                new FullyQualifiedProcedureId(wbMember, mdAoMemberUtils, procFetchMemberTable);
        //
        procedureReference = new ProcedureReference(referrer, referee);
    }

    @Test
    public void test_getReferrer() {
        FullyQualifiedProcedureId referrer = procedureReference.getReferrer();
        assertThat(referrer).isNotNull();
        VBAModule module = referrer.getModule();
        assertThat(module.getName()).isEqualTo("年会費納入状況チェック");
        VBAProcedure procedure = referrer.getProcedure();
        assertThat(procedure.getName()).isEqualTo("Main");
    }

    @Test
    public void test_getReferee() {
        FullyQualifiedProcedureId referee = procedureReference.getReferee();
        assertThat(referee).isNotNull();
        VBAModule module = referee.getModule();
        assertThat(module.getName()).isEqualTo("AoMemberUtils");
        VBAProcedure procedure = referee.getProcedure();
        assertThat(procedure.getName()).isEqualTo("FetchMemberTable");
    }

    @Test
    public void test_toJson() throws JsonProcessingException {
        String json = procedureReference.toJson();
        assertThat(json).isNotNull();
        assertThat(json).contains("referrer");
        assertThat(json).contains("referee");
        logger.debug("[test_toJson] " + json);
    }

    @Test
    public void test_toString() throws IOException {
        String json = procedureReference.toString();
        assertThat(json).isNotNull();
        logger.debug("[test_toString] " + json);
        Path out = classOutputDir.resolve("test_toString.json");
        Files.writeString(out, json);
        assertThat(out).exists();
    }
}

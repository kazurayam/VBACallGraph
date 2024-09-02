package com.kazurayam.vba.puml;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.example.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

public class VBAProcedureReferenceTest {
    private static final Logger logger = LoggerFactory.getLogger(VBAProcedureReferenceTest.class);
    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(VBAProcedureReferenceTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(VBAProcedureReferenceTest.class)
                    .build();

    private Path classOutputDir;
    private VBAProcedureReference procedureReference;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        //
        ModelWorkbook wbFeePaymentCheck =
                new ModelWorkbook(
                        MyWorkbook.FeePaymentControl.resolveWorkbookUnder(),
                        MyWorkbook.FeePaymentControl.resolveSourceDirUnder())
                        .id(MyWorkbook.FeePaymentControl.getId());
        VBAModule md会費納入状況チェック = wbFeePaymentCheck.getModule("会費納入状況チェック");
        FullyQualifiedVBAModuleId referrer =
                new FullyQualifiedVBAModuleId(wbFeePaymentCheck, md会費納入状況チェック);
        VBASource referrerModuleSource = md会費納入状況チェック.getVBASource();
        VBASourceLine referrerSourceLine =
                new VBASourceLine(52,
                        "    Set memberTable = MbMemberTableUtil.FetchMemberTable(memberFile, \"R6年度\", ThisWorkbook)\n");
        //
        ModelWorkbook wbMember =
                new ModelWorkbook(
                        MyWorkbook.Member.resolveWorkbookUnder(),
                        MyWorkbook.Member.resolveSourceDirUnder())
                        .id(MyWorkbook.Member.getId());
        VBAModule mdMbMemberTableUtil = wbMember.getModule("MbMemberTableUtil");
        VBAProcedure procFetchMemberTable =
                mdMbMemberTableUtil.getProcedure("FetchMemberTable");
        FullyQualifiedVBAProcedureId referee =
                new FullyQualifiedVBAProcedureId(wbMember, mdMbMemberTableUtil,
                        procFetchMemberTable);
        //
        procedureReference =
                new VBAProcedureReference(referrer,
                        referrerModuleSource, referrerSourceLine, referee);
    }

    @Test
    public void test_getReferrer() {
        FullyQualifiedVBAModuleId referrer = procedureReference.getReferrer();
        assertThat(referrer).isNotNull();
        VBAModule module = referrer.getModule();
        assertThat(module.getName()).isEqualTo("会費納入状況チェック");
    }

    @Test
    public void test_getReferee
            () {
        FullyQualifiedVBAProcedureId referee = procedureReference.getReferee();
        assertThat(referee).isNotNull();
        VBAModule module = referee.getModule();
        assertThat(module.getName()).isEqualTo("MbMemberTableUtil");
        VBAProcedure procedure = referee.getProcedure();
        assertThat(procedure.getProcedure()).isEqualTo("FetchMemberTable");
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

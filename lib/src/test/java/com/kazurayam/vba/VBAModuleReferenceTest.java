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

public class VBAModuleReferenceTest {
    private static final Logger logger = LoggerFactory.getLogger(VBAModuleReferenceTest.class);
    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(VBAModuleReferenceTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(VBAModuleReferenceTest.class)
                    .build();
    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");

    private Path classOutputDir;
    private VBAModuleReference moduleReference;

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
        FullyQualifiedVBAModuleId referrer =
                new FullyQualifiedVBAModuleId(wbFeePaymentCheck, md年会費納入状況チェック);
        //
        SensibleWorkbook wbMember =
                new SensibleWorkbook(
                        MyWorkbook.Member.getId(),
                        MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                        MyWorkbook.Member.resolveSourceDirUnder(baseDir));
        VBAModule mdAoMemberUtils = wbMember.getModule("AoMemberUtils");
        FullyQualifiedVBAModuleId referee =
                new FullyQualifiedVBAModuleId(wbMember, mdAoMemberUtils);
        //
        moduleReference = new VBAModuleReference(referrer, referee);
    }

    @Test
    public void test_getReferrer() {
        FullyQualifiedVBAModuleId referrer = moduleReference.getReferrer();
        assertThat(referrer).isNotNull();
        VBAModule module = referrer.getModule();
        assertThat(module.getName()).isEqualTo("年会費納入状況チェック");
    }

    @Test
    public void test_getReferee() {
        FullyQualifiedVBAModuleId referee = moduleReference.getReferee();
        assertThat(referee).isNotNull();
        VBAModule module = referee.getModule();
        assertThat(module.getName()).isEqualTo("AoMemberUtils");
    }

    @Test
    public void test_toJson() throws JsonProcessingException {
        String json = moduleReference.toJson();
        assertThat(json).isNotNull();
        assertThat(json).contains("referrer");
        assertThat(json).contains("referee");
        logger.debug("[test_toJson] " + json);
    }

    @Test
    public void test_toString() throws IOException {
        String json = moduleReference.toString();
        assertThat(json).isNotNull();
        logger.debug("[test_toString] " + json);
        Path out = classOutputDir.resolve("test_toString.json");
        Files.writeString(out, json);
        assertThat(out).exists();
    }
}

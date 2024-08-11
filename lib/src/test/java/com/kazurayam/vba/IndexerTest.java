package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

public class IndexerTest {
    private static final Logger logger =
            LoggerFactory.getLogger(IndexerTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(IndexerTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(IndexerTest.class)
                    .build();

    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");
    private Path classOutputDir;
    private Indexer indexer;
    private FullyQualifiedProcedureId referee;
    private ProcedureReference expectedReference;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        indexer = new Indexer();
        // FeePymentCheck workbook
        SensibleWorkbook wbFeePaymentCheck =
                new SensibleWorkbook(
                        MyWorkbook.FeePaymentCheck.getId(),
                        MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                        MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir));
        indexer.add(wbFeePaymentCheck);
        VBAModule md年会費納入状況チェック = wbFeePaymentCheck.getModule("年会費納入状況チェック");
        VBAProcedure procMain = md年会費納入状況チェック.getProcedure("Main");
        FullyQualifiedProcedureId referrer =
                new FullyQualifiedProcedureId(wbFeePaymentCheck,
                        md年会費納入状況チェック, procMain);

        // Member workbook
        SensibleWorkbook wbMember =
                new SensibleWorkbook(
                        MyWorkbook.Member.getId(),
                        MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                        MyWorkbook.Member.resolveSourceDirUnder(baseDir)
                );
        indexer.add(wbMember);
        VBAModule mdAoMemberUtil =
                wbMember.getModule("AoMemberUtils");
        VBAProcedure prFetchMemberTable =
                mdAoMemberUtil.getProcedure("FetchMemberTable");
        referee = new FullyQualifiedProcedureId(wbMember,
                        mdAoMemberUtil, prFetchMemberTable);
        //
        expectedReference = new ProcedureReference(referrer, referee);
    }

    @Test
    public void test_getWorkbooks() {
        List<SensibleWorkbook> workbookList = indexer.getWorkbooks();
        assertThat(workbookList).hasSize(2);
    }

    /**
     * This is the most interesting part of this project!
     */
    @Test
    public void test_findReferenceTo() {
        Set<ProcedureReference> foundReferences =
                indexer.findReferenceTo(referee);
        assertThat(foundReferences).isNotNull();
        assertThat(foundReferences).hasSize(17);
        assertThat(foundReferences).contains(expectedReference);
    }

    @Test
    public void test_xref() {
        List<SensibleWorkbook> workbookList = indexer.getWorkbooks();
        Set<ProcedureReference> foundReferences =
                indexer.xref(workbookList, referee);
        assertThat(foundReferences).isNotNull();
        assertThat(foundReferences).hasSize(17);
        assertThat(foundReferences).contains(expectedReference);
    }

    @Test
    public void test_toJson() throws JsonProcessingException {
        String json = indexer.toJson();
        assertThat(json).isNotNull();
        logger.debug("[test_toJson] " + json);
    }

    @Test
    public void test_toString() throws IOException {
        String json = indexer.toString();
        assertThat(json).isNotNull();
        Path out = classOutputDir.resolve("test_toString.json");
        Files.writeString(out, json);
    }
}

package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Set;
import java.util.SortedSet;

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
    private FullyQualifiedVBAProcedureId referee;
    private VBAProcedureReference expectedReference;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        indexer = new Indexer();
        indexer.setOptions(Options.KAZURAYAM);

        // FeePaymentCheck workbook
        SensibleWorkbook wbFeePaymentCheck =
                new SensibleWorkbook(
                        MyWorkbook.FeePaymentCheck.getId(),
                        MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                        MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir));
        indexer.add(wbFeePaymentCheck);
        VBAModule md年会費納入状況チェック = wbFeePaymentCheck.getModule("年会費納入状況チェック");
        FullyQualifiedVBAModuleId referrer =
                new FullyQualifiedVBAModuleId(wbFeePaymentCheck, md年会費納入状況チェック);
        VBASource referrerModuleSource = md年会費納入状況チェック.getVBASource();
        VBASourceLine referrerSourceLine =
                new VBASourceLine(51,
                        "    Set memberTable = AoMemberUtils.FetchMemberTable(memberFile, \"R6年度\", ThisWorkbook)");

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
        referee = new FullyQualifiedVBAProcedureId(wbMember,
                        mdAoMemberUtil, prFetchMemberTable);

        //
        expectedReference =
                new VBAProcedureReference(referrer,
                        referrerModuleSource, referrerSourceLine, referee);
    }

    @Test
    public void test_getWorkbooks() {
        List<SensibleWorkbook> workbookList = indexer.getWorkbooks();
        assertThat(workbookList).hasSize(2);
    }

    @Test
    public void test_findAllProcedureReferences() throws IOException {
        SortedSet<VBAProcedureReference> memo =
                indexer.findAllProcedureReferences();
        assertThat(memo).isNotNull();
        assertThat(memo.size()).isEqualTo(1);
        Path out = classOutputDir.resolve("test_findAllProcedureReferences.txt");
        PrintWriter pw = new PrintWriter(Files.newBufferedWriter(out));
        for (VBAProcedureReference ref : memo) {
            pw.println(ref.toString());
        }
        pw.flush();
        pw.close();
    }

    /**
     * This is the most interesting part of this project!
     */
    @Test
    public void test_findProcedureReferenceTo() throws IOException {
        Set<VBAProcedureReference> references =
                indexer.findProcedureReferenceTo(referee);
        assertThat(references).isNotNull();
        assertThat(references).hasSize(1);
        assertThat(references).contains(expectedReference);
        Path out = classOutputDir.resolve("test_findProcedureReferenceTo.txt");
        PrintWriter pw = new PrintWriter(Files.newBufferedWriter(out));
        for (VBAProcedureReference reference : references) {
            pw.println(reference);
        }
        pw.flush();
        pw.close();
    }

    @Test
    public void test_shouldIgnore_Initialize() throws IOException {
        SensibleWorkbook wbBackbone =
                new SensibleWorkbook(
                        MyWorkbook.Backbone.getId(),
                        MyWorkbook.Backbone.resolveWorkbookUnder(baseDir),
                        MyWorkbook.Backbone.resolveSourceDirUnder(baseDir));
        VBAModule mdDocTransformer = wbBackbone.getModule("DocTransformer");
        VBAProcedure prInitialize =
                mdDocTransformer.getProcedure("Initialize");
        FullyQualifiedVBAProcedureId referee =
                new FullyQualifiedVBAProcedureId(wbBackbone, mdDocTransformer,
                        prInitialize);
        //
        assertThat(indexer.shouldIgnore(referee)).isTrue();
    }

    @Test
    public void test_shouldIgnore_Class_Initialize() throws IOException {
        SensibleWorkbook wbCashbook =
                new SensibleWorkbook(
                        MyWorkbook.Cashbook.getId(),
                        MyWorkbook.Cashbook.resolveWorkbookUnder(baseDir),
                        MyWorkbook.Cashbook.resolveSourceDirUnder(baseDir));
        VBAModule mdCash = wbCashbook.getModule("Cash");
        VBAProcedure prClassInitialize =
                mdCash.getProcedure("Class_Initialize");
        FullyQualifiedVBAProcedureId referee =
                new FullyQualifiedVBAProcedureId(wbCashbook, mdCash,
                        prClassInitialize);
        //
        assertThat(indexer.shouldIgnore(referee)).isTrue();
    }

    @Test
    public void test_xref() {
        List<SensibleWorkbook> workbookList = indexer.getWorkbooks();
        Set<VBAProcedureReference> foundReferences =
                indexer.xref(workbookList, referee);
        assertThat(foundReferences).isNotNull();
        assertThat(foundReferences).hasSize(1);
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

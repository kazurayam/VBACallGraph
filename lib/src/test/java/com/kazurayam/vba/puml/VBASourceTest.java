package com.kazurayam.vba.puml;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.example.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.regex.Pattern;

import static org.assertj.core.api.Assertions.assertThat;

public class VBASourceTest {

    private static final Logger logger =
            LoggerFactory.getLogger(VBASourceTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(VBASourceTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(VBASourceTest.class)
                    .build();


    private Path classOutputDir;

    private VBASource source会費納入状況チェック;
    private VBASource sourceAccountsFinder;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        source会費納入状況チェック = createFixture(MyWorkbook.FeePaymentControl,
                        "会費納入状況チェック.bas");
        sourceAccountsFinder = createFixture(MyWorkbook.Cashbook,
                "CashbookTransformer.cls");
    }

    private VBASource createFixture(MyWorkbook myWorkbook, String fileName) {
        Path exportedSourceDir =
                myWorkbook.resolveSourceDirUnder();
        Path moduleSource = exportedSourceDir.resolve(fileName);
        String moduleName = getModuleNameFromSourceFileName(moduleSource);
        VBASource vbaSource = new VBASource(moduleName, moduleSource);
        assertThat(vbaSource).isNotNull();
        return vbaSource;
    }

    private String getModuleNameFromSourceFileName(Path moduleSource) {
        String fileName = moduleSource.getFileName().toString();
        return fileName.substring(0, fileName.lastIndexOf("."));
    }

    @Test
    public void test_getModuleName() {
        String moduleName = source会費納入状況チェック.getModuleName();
        assertThat(moduleName).isEqualTo("会費納入状況チェック");
    }

    @Test
    public void test_getSourcePath() {
        Path sourcePath = source会費納入状況チェック.getSourcePath();
        assertThat(sourcePath.getFileName().toString())
                .isEqualTo("会費納入状況チェック.bas");
        assertThat(sourcePath).exists();
    }

    @Test
    public void test_loadCode() throws IOException {
        List<String> code = source会費納入状況チェック.loadCode();
        assertThat(code).hasSizeGreaterThan(0);
        //logger.debug("[test_loadCode]\n" + String.join("\n", code));
    }


    @Test
    public void test_find_FetchMemberTable() throws IOException {
        List<Pattern> patterns = ProcedureNamePatternManager
                .createPatterns("FetchMemberTable");
        assertThat(patterns).isNotNull();
        List<VBASourceLine> linesFound = source会費納入状況チェック.find(patterns);
        assertThat(linesFound).hasSize(1);
        VBASourceLine sourceLine = linesFound.get(0);
        assertThat(sourceLine.getFound()).isTrue();
        //
        Path out = classOutputDir.resolve("test_find_FetchMemberTable.json");
        Files.writeString(out, linesFound.get(0).toString());
        assertThat(out).exists();
    }

    /**
     * see
     * https://github.com/kazurayam/VBACallGraph/issues/39
     */
    @Test
    public void test_find_AccountName() throws IOException {
        List<Pattern> patterns = ProcedureNamePatternManager
                .createPatterns("AccountName");
        List<VBASourceLine> linesFound = sourceAccountsFinder.find(patterns);
        logger.info("[test_find_AccountName] linesFound=" + linesFound.toString());
        /*
11:24:36.977 [Test worker] INFO com.kazurayam.vba.puml.VBASourceTest -- linesFound=[{
  "lineNo" : 75,
  "line" : "        Set dic2(k) = cSel_.SelectCashList(acc.accType, acc.AccountName, acc.SubAccountName, _"
}]
         */
        assertThat(linesFound).hasSize(1);
    }

    @Test
    public void test_toString() throws IOException {
        List<Pattern> patterns = ProcedureNamePatternManager.createPatterns("OpenMemberTable");
        assertThat(patterns).isNotNull();
        List<VBASourceLine> linesFound = source会費納入状況チェック.find(patterns);
        String json = source会費納入状況チェック.toString();
        Path out = classOutputDir.resolve("test_toString.json");
        Files.writeString(out, json);
        assertThat(json).contains("moduleName");
        assertThat(json).contains("sourcePath");
    }


}

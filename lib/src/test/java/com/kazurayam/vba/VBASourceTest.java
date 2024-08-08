package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Ignore;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

public class VBASourceTest {

    private static final Logger logger =
            LoggerFactory.getLogger(VBASourceTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(VBASourceTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(VBASourceTest.class)
                    .build();

    private static final Path baseDir = too.getProjectDirectory().resolve("src/test/fixture/hub");

    private Path classOutputDir;

    private VBASource vbaSource;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        Path exportedSourceDir =
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir);
        Path moduleSource = exportedSourceDir.resolve("年会費納入状況チェック.bas");
        vbaSource = new VBASource("年会費納入状況チェック", moduleSource);
        assertThat(vbaSource).isNotNull();
    }

    @Test
    public void test_getModuleName() {
        String moduleName = vbaSource.getModuleName();
        assertThat(moduleName).isEqualTo("年会費納入状況チェック");
    }

    @Test
    public void test_getSourcePath() {
        Path sourcePath = vbaSource.getSourcePath();
        assertThat(sourcePath.getFileName().toString())
                .isEqualTo("年会費納入状況チェック.bas");
        assertThat(sourcePath).exists();
    }

    @Test
    public void test_loadCode() throws IOException {
        List<String> code = vbaSource.loadCode();
        assertThat(code).hasSizeGreaterThan(0);
        //logger.debug("[test_loadCode]\n" + String.join("\n", code));
    }


    @Test
    public void test_find() throws IOException {
        List<VBASourceLine> linesFound =
                vbaSource.find(".FetchMemberTable(");
        assertThat(linesFound).hasSize(1);
        VBASourceLine sourceLine = linesFound.get(0);
        assertThat(sourceLine.getFound()).isTrue();
        //
        Path out = classOutputDir.resolve("test_find.json");
        Files.writeString(out, linesFound.get(0).toString());
        assertThat(out).exists();
    }

    @Test
    public void test_toString() throws IOException {
        List<VBASourceLine> linesFound =
            vbaSource.find("OpenMemberTable");
        String json = vbaSource.toString();
        Path out = classOutputDir.resolve("test_toString.json");
        Files.writeString(out, json);
        assertThat(json).contains("moduleName");
        assertThat(json).contains("sourcePath");
    }

    @Test
    public void test_escapeAsRegex() {
        String escaped = VBASource.escapeAsRegex(".FetchMemberTable(");
        assertThat(escaped).isEqualTo("\\.FetchMemberTable\\(");
    }
}

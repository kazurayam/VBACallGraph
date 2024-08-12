package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.nio.file.Files;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class VBASourceLineTest {

    private static final Logger logger =
            LoggerFactory.getLogger(VBASourceLineTest.class);

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(VBASourceLineTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(VBASourceLineTest.class)
                    .build();

    private static final Path baseDir = too.getProjectDirectory().resolve("src/test/fixture/hub");

    private Path classOutputDir;

    private VBASourceLine vbaSourceLine;

    @BeforeTest
    public void beforeTest() throws IOException {
        classOutputDir = too.cleanClassOutputDirectory();
        /*
        Path exportedSourceDir =
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir);
        Path moduleSource = exportedSourceDir.resolve("年会費納入状況チェック.bas");
        VBASource vbaSource = new VBASource("年会費納入状況チェック", moduleSource);
        assertThat(vbaSource).isNotNull();
        */
    }

    @BeforeMethod
    public void before() {
        vbaSourceLine = new VBASourceLine(52,
                "Set memberTable = AoMemberUtils.FetchMemberTable(memberFile, \"R6年度\", ThisWorkbook)");
    }

    @Test
    public void test_constructor() {
        assertThat(vbaSourceLine).isNotNull();
    }

    @Test
    public void test_getLineNo() {
        assertThat(vbaSourceLine.getLineNo()).isEqualTo(52);
    }

    @Test
    public void test_getLine() {
        assertThat(vbaSourceLine.getLine()).contains(".FetchMemberTable(");
    }

    @Test
    public void test_getFound_getMatcher() {
        assertThat(vbaSourceLine.getMatcher()).isNull();
        VBASourceLine augmented = augment(vbaSourceLine);
        assertThat(augmented).isNotNull();
        assertThat(augmented.getFound()).isTrue();
        assertThat(augmented.getMatcher().pattern().pattern())
                .isEqualTo("\\.FetchMemberTable\\(");
    }

    private VBASourceLine augment(VBASourceLine vbaSourceLine) {
        String patternString = ".FetchMemberTable(";
        Pattern ptn = Pattern.compile(PatternManager.escapeAsRegex(patternString));
        Matcher m = ptn.matcher(vbaSourceLine.getLine());
        Boolean found = m.find();
        vbaSourceLine.setFound(found);
        vbaSourceLine.setMatcher(m);
        return vbaSourceLine;
    }

    @Test
    public void test_toString() throws IOException {
        VBASourceLine augmented = augment(vbaSourceLine);
        String json = augmented.toString();
        Path output = classOutputDir.resolve("test_toString.json");
        Files.writeString(output, json);
        assertThat(output).exists();
    }
}
package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class ProcedureUsageAnalyzerTest {

    private static Logger logger =
            LoggerFactory.getLogger(ProcedureUsageAnalyzerTest.class);
    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(ProcedureUsageAnalyzerTest.class)
                    .subOutputDirectory(ProcedureUsageAnalyzerTest.class)
                    .build();
    private static final Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");
    private ProcedureUsageAnalyzer analyzer;

    @BeforeTest
    public void beforeTest() throws IOException {
        analyzer = new ProcedureUsageAnalyzer();

        analyzer.add(new SensibleWorkbook(
                MyWorkbook.FeePaymentCheck.getId(),
                MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir)
        ));

        analyzer.add(new SensibleWorkbook(
                MyWorkbook.Cashbook.getId(),
                MyWorkbook.Cashbook.resolveWorkbookUnder(baseDir),
                MyWorkbook.Cashbook.resolveSourceDirUnder(baseDir)
        ));

        analyzer.add(new SensibleWorkbook(
                MyWorkbook.Member.getId(),
                MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                MyWorkbook.Member.resolveSourceDirUnder(baseDir)
        ));

        analyzer.add(new SensibleWorkbook(
                MyWorkbook.Backbone.getId(),
                MyWorkbook.Backbone.resolveWorkbookUnder(baseDir),
                MyWorkbook.Backbone.resolveSourceDirUnder(baseDir)
        ));
    }

    @Test
    public void test_toString() {
        String str = analyzer.toString();
        assertThat(str).isNotNull();
        logger.info("[test_toString] " + str);
    }
}

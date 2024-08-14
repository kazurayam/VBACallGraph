package com.kazurayam.vbastudy;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.SensibleWorkbook;
import com.kazurayam.vbaexample.MyWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.poifs.macros.VBAMacroReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Map;

/**
 * https://github.com/kazurayam/VBAProcedureUsageAnalyzer/issues/28
 */
public class GettingVBProjectNameTest {
    Logger logger = LoggerFactory.getLogger(GettingVBProjectNameTest.class);

    private TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(GettingVBProjectNameTest.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(GettingVBProjectNameTest.class)
                    .build();

    private Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");

    private SensibleWorkbook wb;

    @BeforeTest
    public void beforeTest() {
    }

    @Test
    public void test_read_macros() throws IOException {
        File workbookFile = MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir).toFile();
        try (POIFSFileSystem fs = new POIFSFileSystem(workbookFile)) {
            VBAMacroReader mr = new VBAMacroReader(fs);
            Map<String,String> macros = mr.readMacros();
            logger.debug("[test_read_macros]" + macros);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

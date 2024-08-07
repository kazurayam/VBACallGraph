package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vbaexample.MyWorkbook;
import org.testng.annotations.Test;
import org.testng.log4testng.Logger;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

public class SourceDirVisitorTest {

    private Logger logger = Logger.getLogger(SourceDirVisitorTest.class);

    private final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(SourceDirVisitorTest.class)
                    .subOutputDirectory(SourceDirVisitorTest.class).build();
    private final Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");
    @Test
    public void test_visit_Backbone() throws IOException {
        Path vbaSourceDir = MyWorkbook.Backbone.resolveSourceDirUnder(baseDir);
        SourceDirVisitor visitor = new SourceDirVisitor();
        Files.walkFileTree(vbaSourceDir, visitor);
        List<Path> list = visitor.getList();
        logger.info("[test_visit_Backbone] : " + list.toString());
        assertThat(list.size()).isGreaterThan(0);
    }

}

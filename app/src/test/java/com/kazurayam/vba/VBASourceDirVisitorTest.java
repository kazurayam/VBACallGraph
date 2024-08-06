package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;
import org.testng.annotations.Test;
import org.testng.log4testng.Logger;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

public class VBASourceDirVisitorTest {

    private Logger logger = Logger.getLogger(VBASourceDirVisitorTest.class);

    private final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(VBASourceDirVisitorTest.class)
                    .subOutputDirectory(VBASourceDirVisitorTest.class).build();
    private final Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");
    @Test
    public void test_visit_Backbone() throws IOException {
        Path vbaSourceDir = WorkbookInstanceLocation.Backbone.resolveVBASourceDirBasedOn(baseDir);
        VBASourceDirVisitor visitor = new VBASourceDirVisitor();
        Files.walkFileTree(vbaSourceDir, visitor);
        List<Path> list = visitor.getList();
        logger.info("[test_visit_Backbone] : " + list.toString());
        assertThat(list.size()).isGreaterThan(0);
    }

}

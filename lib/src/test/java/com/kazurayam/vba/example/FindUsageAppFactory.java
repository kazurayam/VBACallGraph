package com.kazurayam.vba.example;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.puml.FindUsageApp;
import com.kazurayam.vba.puml.Options;
import com.kazurayam.vba.puml.SensibleWorkbook;

import java.io.IOException;
import java.nio.file.Path;

public class FindUsageAppFactory {

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(FindUsageAppFactory.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(FindUsageAppFactory.class)
                    .build();

    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");

    private FindUsageAppFactory() {}

    public static FindUsageApp createKazurayamSeven() throws IOException {

        FindUsageApp app = new FindUsageApp();

        app.add(new SensibleWorkbook(
                MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.FeePaymentCheck.getId()));

        app.add(new SensibleWorkbook(
                MyWorkbook.PleasePayFeeLetter.resolveWorkbookUnder(baseDir),
                MyWorkbook.PleasePayFeeLetter.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.PleasePayFeeLetter.getId()));

        app.add(new SensibleWorkbook(
                MyWorkbook.WebCredentials.resolveWorkbookUnder(baseDir),
                MyWorkbook.WebCredentials.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.WebCredentials.getId()));

        app.add(new SensibleWorkbook(
                MyWorkbook.Settlement.resolveWorkbookUnder(baseDir),
                MyWorkbook.Settlement.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Settlement.getId()
                ));


        app.add(new SensibleWorkbook(
                MyWorkbook.Cashbook.resolveWorkbookUnder(baseDir),
                MyWorkbook.Cashbook.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Cashbook.getId()));

        app.add(new SensibleWorkbook(
                MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                MyWorkbook.Member.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Member.getId()));

        app.add(new SensibleWorkbook(
                MyWorkbook.Backbone.resolveWorkbookUnder(baseDir),
                MyWorkbook.Backbone.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Backbone.getId()));

        app.setOptions(Options.KAZURAYAM);

        return app;
    }
}

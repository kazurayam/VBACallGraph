package com.kazurayam.vba.example;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.puml.CallGraphApp;
import com.kazurayam.vba.puml.Options;
import com.kazurayam.vba.puml.ModelWorkbook;

import java.io.IOException;
import java.nio.file.Path;

public class CallGraphAppFactory {

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(CallGraphAppFactory.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(CallGraphAppFactory.class)
                    .build();

    private static final Path baseDir =
            too.getProjectDirectory().resolve("src/test/fixture/hub");

    private CallGraphAppFactory() {}

    public static CallGraphApp createKazurayamSeven() throws IOException {

        CallGraphApp app = new CallGraphApp();

        app.add(new ModelWorkbook(
                MyWorkbook.FeePaymentCheck.resolveWorkbookUnder(baseDir),
                MyWorkbook.FeePaymentCheck.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.FeePaymentCheck.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.PleasePayFeeLetter.resolveWorkbookUnder(baseDir),
                MyWorkbook.PleasePayFeeLetter.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.PleasePayFeeLetter.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.WebCredentials.resolveWorkbookUnder(baseDir),
                MyWorkbook.WebCredentials.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.WebCredentials.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.Settlement.resolveWorkbookUnder(baseDir),
                MyWorkbook.Settlement.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Settlement.getId()
                ));


        app.add(new ModelWorkbook(
                MyWorkbook.Cashbook.resolveWorkbookUnder(baseDir),
                MyWorkbook.Cashbook.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Cashbook.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.Member.resolveWorkbookUnder(baseDir),
                MyWorkbook.Member.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Member.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.Backbone.resolveWorkbookUnder(baseDir),
                MyWorkbook.Backbone.resolveSourceDirUnder(baseDir))
                .id(MyWorkbook.Backbone.getId()));

        app.setOptions(Options.KAZURAYAM);

        return app;
    }
}

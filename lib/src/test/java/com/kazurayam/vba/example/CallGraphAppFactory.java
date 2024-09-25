package com.kazurayam.vba.example;

import com.kazurayam.unittest.TestOutputOrganizer;
import com.kazurayam.vba.puml.CallGraphApp;
import com.kazurayam.vba.puml.Options;
import com.kazurayam.vba.puml.ModelWorkbook;

import java.io.IOException;

public class CallGraphAppFactory {

    private static final TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(CallGraphAppFactory.class)
                    .outputDirectoryRelativeToProject("build/tmp/testOutput")
                    .subOutputDirectory(CallGraphAppFactory.class)
                    .build();

    private CallGraphAppFactory() {}

    public static CallGraphApp createKazurayamSeven() throws IOException {

        CallGraphApp app = new CallGraphApp();

        app.add(new ModelWorkbook(
                MyWorkbook.FeePaymentControl.resolveWorkbookUnder(),
                MyWorkbook.FeePaymentControl.resolveSourceDirUnder())
                .id(MyWorkbook.FeePaymentControl.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.WebCredentials.resolveWorkbookUnder(),
                MyWorkbook.WebCredentials.resolveSourceDirUnder())
                .id(MyWorkbook.WebCredentials.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.Settlement.resolveWorkbookUnder(),
                MyWorkbook.Settlement.resolveSourceDirUnder())
                .id(MyWorkbook.Settlement.getId()
                ));

        app.add(new ModelWorkbook(
                MyWorkbook.Cashbook.resolveWorkbookUnder(),
                MyWorkbook.Cashbook.resolveSourceDirUnder())
                .id(MyWorkbook.Cashbook.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.Member.resolveWorkbookUnder(),
                MyWorkbook.Member.resolveSourceDirUnder())
                .id(MyWorkbook.Member.getId()));

        app.add(new ModelWorkbook(
                MyWorkbook.Backbone.resolveWorkbookUnder(),
                MyWorkbook.Backbone.resolveSourceDirUnder())
                .id(MyWorkbook.Backbone.getId()));

        app.setOptions(Options.KAZURAYAM);

        return app;
    }

    public static CallGraphApp createKazurayamSevenPlus() throws IOException {

        CallGraphApp app = createKazurayamSeven();

        app.add(new ModelWorkbook(
                MyWorkbook.VBACallGraphSetup.resolveWorkbookUnder(),
                MyWorkbook.VBACallGraphSetup.resolveSourceDirUnder())
                .id(MyWorkbook.VBACallGraphSetup.getId()));

        app.setOptions(Options.KAZURAYAM);

        return app;
    }

    public static CallGraphApp createPerfectBook() throws IOException {
        CallGraphApp app = new CallGraphApp();
        app.add(new ModelWorkbook(
                MyWorkbook.PerfectExcelVBA.resolveWorkbookUnder(),
                MyWorkbook.PerfectExcelVBA.resolveSourceDirUnder())
                .id(MyWorkbook.PerfectExcelVBA.getId()));
        app.add(new ModelWorkbook(
                MyWorkbook.Backbone.resolveWorkbookUnder(),
                MyWorkbook.Backbone.resolveSourceDirUnder())
                .id(MyWorkbook.Backbone.getId()));
        app.add(new ModelWorkbook(
                MyWorkbook.VBACallGraphSetup.resolveWorkbookUnder(),
                MyWorkbook.VBACallGraphSetup.resolveSourceDirUnder())
                .id(MyWorkbook.VBACallGraphSetup.getId()));
        app.setOptions(Options.KAZURAYAM);

        return app;
    }

}

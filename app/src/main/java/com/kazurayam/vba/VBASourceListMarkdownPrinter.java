package com.kazurayam.vba;

import com.kazurayam.unittest.TestOutputOrganizer;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public class VBASourceListMarkdownPrinter {

    private final List<WorkbookInstance> vbaSourceDirList;

    public VBASourceListMarkdownPrinter() {
         vbaSourceDirList = new ArrayList<>();
    }

    public void add(WorkbookInstance vbaSource) {
        vbaSourceDirList.add(vbaSource);
    }

    public void printAllVBASourceDirs(Writer writer) throws IOException {
        BufferedWriter bw = new BufferedWriter(writer);
        for (WorkbookInstance resolved : vbaSourceDirList) {
            Path baseDir = resolved.getBaseDir();
            Path targetDir = resolved.getWorkbookInstanceLocation().resolveVBASourceDirBasedOn(baseDir);
            VBASourceDirVisitor visitor =
                    new VBASourceDirVisitor();
            Files.walkFileTree(targetDir, visitor);
            List<Path> sources = visitor.getList();
            this.printVBASourceDir(resolved, sources, bw);
            bw.write("\n\n");
        }
        bw.flush();
        bw.close();
    }

    public void printVBASourceDir(WorkbookInstance resolved,
                                  List<Path> sources,
                                  Writer writer) {
        PrintWriter pw = new PrintWriter(new BufferedWriter(writer));
        pw.println("### " + resolved.getWorkbookInstanceLocation().getId());
        pw.println("|No.|file name|");
        pw.println("|--:|:--------|");
        List<String> sortedFileNames =
                sources.stream()
                        .map(p -> p.getFileName().toString())
                        .sorted()
                        .toList();
        for (int i = 0; i < sortedFileNames.size(); i++) {
            pw.println(String.format("|%d|%s|", i+1, sortedFileNames.get(i)));
        }
        pw.flush();
    }

    public static void main(String[] args) throws IOException {
        TestOutputOrganizer too =
                new TestOutputOrganizer.Builder(VBASourceListMarkdownPrinter.class)
                        .subOutputDirectory(VBASourceListMarkdownPrinter.class)
                        .build();
        Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");
        VBASourceListMarkdownPrinter printer = new VBASourceListMarkdownPrinter();
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.Backbone));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.Member));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.Cashbook));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.Settlement));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.FeePaymentCheck));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.PleasePayFeeLetter));
        printer.add(new WorkbookInstance(baseDir, WorkbookInstanceLocation.WebCredentials));
        //
        Path report = too.getProjectDirectory().resolve("../../docs/MyVBASourceDirs.md");
        assert Files.exists(report.getParent());
        Writer writer = Files.newBufferedWriter(report);
        printer.printAllVBASourceDirs(writer);
    }
}

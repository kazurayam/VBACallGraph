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

public class VBASourceMarkdownPrinter {

    private final List<ResolvedVBASourceDir> vbaSourceDirList;

    public VBASourceMarkdownPrinter() {
         vbaSourceDirList = new ArrayList<>();
    }

    public void add(ResolvedVBASourceDir vbaSource) {
        vbaSourceDirList.add(vbaSource);
    }

    public void printAllVBASourceDirs(Writer writer) throws IOException {
        BufferedWriter bw = new BufferedWriter(writer);
        for (ResolvedVBASourceDir resolved : vbaSourceDirList) {
            Path baseDir = resolved.getBaseDir();
            Path targetDir = resolved.getVBASourceDir().resolveBasedOn(baseDir);
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

    public void printVBASourceDir(ResolvedVBASourceDir resolved,
                                  List<Path> sources,
                                  Writer writer) {
        PrintWriter pw = new PrintWriter(new BufferedWriter(writer));
        pw.println("### " + resolved.getVBASourceDir().getId());
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
                new TestOutputOrganizer.Builder(VBASourceMarkdownPrinter.class)
                        .subOutputDirectory(VBASourceMarkdownPrinter.class)
                        .build();
        Path baseDir = too.getProjectDirectory().resolve("../../../github-aogan");
        VBASourceMarkdownPrinter printer = new VBASourceMarkdownPrinter();
        printer.add(new ResolvedVBASourceDir(baseDir, VBASourceDir.Backbone));
        printer.add(new ResolvedVBASourceDir(baseDir, VBASourceDir.Member));
        printer.add(new ResolvedVBASourceDir(baseDir, VBASourceDir.Cashbook));
        printer.add(new ResolvedVBASourceDir(baseDir, VBASourceDir.Settlement));
        printer.add(new ResolvedVBASourceDir(baseDir, VBASourceDir.FeePaymentCheck));
        printer.add(new ResolvedVBASourceDir(baseDir, VBASourceDir.PleasePayFeeLetter));
        printer.add(new ResolvedVBASourceDir(baseDir, VBASourceDir.WebCredentials));
        //
        Path report = too.getProjectDirectory().resolve("../../docs/MyVBASourceDirs.md");
        assert Files.exists(report.getParent());
        Writer writer = Files.newBufferedWriter(report);
        printer.printAllVBASourceDirs(writer);
    }
}

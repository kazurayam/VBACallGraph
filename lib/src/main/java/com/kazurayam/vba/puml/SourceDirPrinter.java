package com.kazurayam.vba.puml;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public class SourceDirPrinter {

    private final List<ModelWorkbook> workbookList;

    public SourceDirPrinter() {
         workbookList = new ArrayList<>();
    }

    public void add(ModelWorkbook workbook) {
        workbookList.add(workbook);
    }

    public void printAllSourceDirs(Writer writer) throws IOException {
        BufferedWriter bw = new BufferedWriter(writer);
        for (ModelWorkbook wb : workbookList) {
            SourceDirVisitor visitor =
                    new SourceDirVisitor();
            Files.walkFileTree(wb.getSourceDirPath(), visitor);
            List<Path> sources = visitor.getList();
            this.printSourceDir(wb, sources, bw);
            bw.write("\n\n");
        }
        bw.flush();
        bw.close();
    }

    void printSourceDir(ModelWorkbook wb,
                        List<Path> sources,
                        Writer writer) {
        PrintWriter pw = new PrintWriter(new BufferedWriter(writer));
        pw.println("### " + wb.getId());
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
}

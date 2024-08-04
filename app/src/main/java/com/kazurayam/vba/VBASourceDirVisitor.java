package com.kazurayam.vba;

import java.io.IOException;
import java.nio.file.FileVisitResult;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.List;

import static java.nio.file.FileVisitResult.CONTINUE;

public class VBASourceDirVisitor extends SimpleFileVisitor<Path> {

    private final List<Path> vbaSources;

    public VBASourceDirVisitor() {
        vbaSources = new ArrayList<Path>();
    }

    @Override
    public FileVisitResult visitFile(Path file, BasicFileAttributes attr) {
        if (file.getFileName().toString().endsWith(".bas") ||
                file.getFileName().toString().endsWith(".cls"))
        vbaSources.add(file);
        return CONTINUE;
    }

    @Override
    public FileVisitResult postVisitDirectory(Path dir,
                                              IOException exc) {
        //System.out.format("Directory: %s%n", dir);
        return CONTINUE;
    }

    @Override
    public FileVisitResult visitFileFailed(Path file,
                                           IOException exc) {
        //System.err.println(exc);
        return CONTINUE;
    }

    public List<Path> getList() {
        return vbaSources;
    }
}

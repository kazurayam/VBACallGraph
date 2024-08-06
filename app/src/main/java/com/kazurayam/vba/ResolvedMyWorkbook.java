package com.kazurayam.vba;

import java.nio.file.Path;

public class ResolvedMyWorkbook implements Comparable<ResolvedMyWorkbook> {
    private final Path baseDir;
    private final MyWorkbook vbaSourceDir;
    public ResolvedMyWorkbook(Path baseDir, MyWorkbook myExcelFile) {
        this.baseDir = baseDir;
        this.vbaSourceDir = myExcelFile;
    }
    public Path getBaseDir() {
        return baseDir;
    }
    public MyWorkbook getVBASourceDir() {
        return vbaSourceDir;
    }
    @Override
    public boolean equals(Object obj) {
        if (!(obj instanceof ResolvedMyWorkbook)) {
            return false;
        }
        ResolvedMyWorkbook other = (ResolvedMyWorkbook)obj;
        if (this.baseDir == other.baseDir) {
            return (this.vbaSourceDir == other.vbaSourceDir);
        } else {
            return false;
        }
    }

    @Override
    public int hashCode() {
        int hash = 7;
        hash = 31 * hash + baseDir.hashCode();
        hash = 31 * hash + vbaSourceDir.hashCode();
        return hash;
    }

    @Override
    public String toString() {
        return String.format("{\"baseDir\":\"%s\", \"VBASource\":\"%s\"",
                baseDir.toString(), vbaSourceDir.getId());
    }

    @Override
    public int compareTo(ResolvedMyWorkbook other) {
        int baseDirComparison = this.baseDir.compareTo(other.baseDir);
        if (baseDirComparison == 0) {
            return this.vbaSourceDir.compareTo(other.vbaSourceDir);
        } else {
            return baseDirComparison;
        }
    }

}

package com.kazurayam.vba;

import java.nio.file.Path;

public class ResolvedVBASourceDir implements Comparable<ResolvedVBASourceDir> {
    private final Path baseDir;
    private final VBASourceDir vbaSourceDir;
    public ResolvedVBASourceDir(Path baseDir, VBASourceDir myExcelFile) {
        this.baseDir = baseDir;
        this.vbaSourceDir = myExcelFile;
    }
    public Path getBaseDir() {
        return baseDir;
    }
    public VBASourceDir getVBASourceDir() {
        return vbaSourceDir;
    }
    @Override
    public boolean equals(Object obj) {
        if (!(obj instanceof ResolvedVBASourceDir)) {
            return false;
        }
        ResolvedVBASourceDir other = (ResolvedVBASourceDir)obj;
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
    public int compareTo(ResolvedVBASourceDir other) {
        int baseDirComparison = this.baseDir.compareTo(other.baseDir);
        if (baseDirComparison == 0) {
            return this.vbaSourceDir.compareTo(other.vbaSourceDir);
        } else {
            return baseDirComparison;
        }
    }

}

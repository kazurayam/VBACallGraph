package com.kazurayam.vba;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

public class ProcedureUsageDiagramGenerator {

    private StringBuilder sb;

    public ProcedureUsageDiagramGenerator() {
        sb = new StringBuilder();
    }

    public void writeStartUml() {
        sb.append("@startuml\n");
    }

    public void writeStartWorkbook(String id) {
        sb.append(String.format("package %s {\n", id));
    }

    public void writeStartModule(String name) {
        sb.append(String.format("  map %s {\n", name));
    }

    public void writeProcedure(String name) {
        sb.append(String.format("    %s =>\n", name));
    }

    public void writeEndModule() {
        sb.append("  }\n");
    }

    public void writeEndWorkbook() {
        sb.append("}\n");
    }

    public void writeEndUml() {
        sb.append("@enduml\n");
    }

    @Override
    public String toString() {
        return sb.toString();
    }

    public void save(File file) throws IOException {
        this.save(file.toPath());
    }
    public void save(Path file) throws IOException {
        Files.writeString(file, sb.toString());
    }

}

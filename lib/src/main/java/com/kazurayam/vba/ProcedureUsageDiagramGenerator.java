package com.kazurayam.vba;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

public class ProcedureUsageDiagramGenerator {

    private final StringBuilder sb;

    public ProcedureUsageDiagramGenerator() {
        sb = new StringBuilder();
    }

    public void writeStartUml() {
        sb.append("@startuml\n");
        sb.append("left to right direction\n");
    }

    public void writeStartWorkbook(SensibleWorkbook wb) {
        sb.append(String.format("package \"workbook %s\" {\n", wb.getId()));
    }

    public void writeStartModule(VBAModule module) {
        sb.append(String.format("  package \"module %s\" {\n", module.getName()));
    }

    public void writeProcedure(VBAModule module, VBAProcedure procedure) {
        sb.append(String.format("    object %s.%s\n", module.getName(), procedure.getName()));
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

    public void generateTextDiagram(File file) throws IOException {
        this.generateTextDiagram(file.toPath());
    }
    public void generateTextDiagram(Path file) throws IOException {
        Files.writeString(file, sb.toString());
    }

}

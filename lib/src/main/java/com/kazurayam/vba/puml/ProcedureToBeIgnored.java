package com.kazurayam.vba.puml;

/**
 *
 */
public class ProcedureToBeIgnored {

    public static ProcedureToBeIgnored Class_Initialize =
            new ProcedureToBeIgnored(VBAModule.ModuleType.Class,
                    "Initialize");

    public static ProcedureToBeIgnored Class_Class_Initialize
            = new ProcedureToBeIgnored(VBAModule.ModuleType.Class,
            "Class_Initialize");

    public static ProcedureToBeIgnored Standard_ExportThisWorkbook
            = new ProcedureToBeIgnored(VBAModule.ModuleType.Standard,
            "ExportThisWorkbook");

    private VBAModule.ModuleType moduleType;
    private String procedureName;

    public ProcedureToBeIgnored(VBAModule.ModuleType moduleType,
                                String procedureName) {
        this.moduleType = moduleType;
        this.procedureName = procedureName;
    }

    public VBAModule.ModuleType getModuleType() {
        return moduleType;
    }

    public String getProcedureName() {
        return procedureName;
    }

    public Boolean matches(FullyQualifiedVBAProcedureId procedureId) {
        return this.getModuleType().equals(procedureId.getModule().getType()) &&
                this.getProcedureName().equals(procedureId.getProcedureName());
    }
}

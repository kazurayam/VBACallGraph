package com.kazurayam.vba;

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

    public static ProcedureToBeIgnored Standard_プロシージャー一覧を作る
            = new ProcedureToBeIgnored(VBAModule.ModuleType.Standard,
            "プロシージャー一覧を作る");

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

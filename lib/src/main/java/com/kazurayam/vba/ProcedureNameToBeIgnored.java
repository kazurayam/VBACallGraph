package com.kazurayam.vba;

/**
 *
 */
public class ProcedureNameToBeIgnored {

    public static ProcedureNameToBeIgnored Class_Initialize =
            new ProcedureNameToBeIgnored(VBAModule.ModuleType.Class,
                    "Initialize");

    public static ProcedureNameToBeIgnored Class_Class_Initialize
            = new ProcedureNameToBeIgnored(VBAModule.ModuleType.Class,
            "Class_Initialize");

    public static ProcedureNameToBeIgnored Standard_プロシージャー一覧を作る
            = new ProcedureNameToBeIgnored(VBAModule.ModuleType.Standard,
            "プロシージャー一覧を作る");

    private VBAModule.ModuleType moduleType;
    private String procedureName;

    public ProcedureNameToBeIgnored(VBAModule.ModuleType moduleType,
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

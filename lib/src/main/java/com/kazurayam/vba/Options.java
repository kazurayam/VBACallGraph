package com.kazurayam.vba;

import java.util.ArrayList;
import java.util.List;

public class Options {

    public static Options DEFAULT = new Options.Builder().build();

    public static Options KAZURAYAM =
            new Options.Builder().ignoreRefereeProcedure(
                    VBAModule.ModuleType.Standard, "プロシージャー一覧を作る")
                    .build();

    private final List<ProcedureNameToBeIgnored> procedureNamesToBeIgnored;
    private final Boolean excludeUnittestModules;

    private Options(Builder builder) {
        this.procedureNamesToBeIgnored = builder.procedureNamesToBeIgnored;
        this.excludeUnittestModules = builder.excludeUnitTestModules;
    }

    public Boolean shouldIgnoreRefereeProcedure(
            FullyQualifiedVBAProcedureId procedureId) {
        for (ProcedureNameToBeIgnored pnbi : procedureNamesToBeIgnored) {
            if (pnbi.matches(procedureId)) {
                return true;
            }
        }
        return false;
    }

    public Boolean shouldExclude(VBAModule module) {
        String moduleNameLowerCase = module.getName().toLowerCase();
        Boolean isUnitTestModule=
                (moduleNameLowerCase.startsWith("test") ||
                        moduleNameLowerCase.endsWith("test"));
        return excludeUnittestModules && isUnitTestModule;
    }

    public Boolean shouldExclude(FullyQualifiedVBAModuleId moduleId) {
        return shouldExclude(moduleId.getModule());
    }

    public Boolean shouldExclude(VBAProcedureReference procedureReference) {
        return shouldExclude(procedureReference.getReferrer().getModule());
    }

    /**
     *
     */
    public static class Builder {
        private final List<ProcedureNameToBeIgnored> procedureNamesToBeIgnored;
        private Boolean excludeUnitTestModules;

        public Builder() {
            this.procedureNamesToBeIgnored = new ArrayList<>();
            this.procedureNamesToBeIgnored.add(ProcedureNameToBeIgnored.Class_Initialize);
            this.procedureNamesToBeIgnored.add(ProcedureNameToBeIgnored.Class_Class_Initialize);
            this.excludeUnitTestModules = true;
        }
        public Builder ignoreRefereeProcedure(VBAModule.ModuleType type,
                                              String procedureName) {
            this.procedureNamesToBeIgnored.add(
                    new ProcedureNameToBeIgnored(type, procedureName));
            return this;
        }
        public Builder excludeUnittestModules(Boolean excludeUnittestModules) {
            this.excludeUnitTestModules = excludeUnittestModules;
            return this;
        }
        public Options build() {
            return new Options(this);
        }
    }
}

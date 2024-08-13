package com.kazurayam.vba;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

public class Options {

    public static Options DEFAULT =
            new Options.Builder().build();

    public static Options RELAXED =
            new Options.Builder()
                    .noExcludeModule()
                    .noIgnoreRefereeProcedure()
                    .build();

    public static Options KAZURAYAM =
            new Options.Builder()
                    .excludeModule(ModuleToBeExcluded.プロシージャー一覧を作る.getPattern())
                    .excludeModule(ModuleToBeExcluded.プロシージャ一覧を作る.getPattern())
                    .excludeModule(ModuleToBeExcluded.プロシジャ一覧を作る.getPattern())
                    .ignoreRefereeProcedure(VBAModule.ModuleType.Standard, "プロシージャー一覧を作る")
                    .ignoreRefereeProcedure(VBAModule.ModuleType.Standard, "プロシージャ一覧を作る")
                    .ignoreRefereeProcedure(VBAModule.ModuleType.Standard, "プロシジャ一覧を作る")
                    .build();

    private final List<ProcedureToBeIgnored> procedureNamesToBeIgnored;
    private final List<ModuleToBeExcluded> moduleNamesToBeExcluded;

    private Options(Builder builder) {
        this.procedureNamesToBeIgnored = builder.procedureToBeIgnoredList;
        this.moduleNamesToBeExcluded = builder.moduleToBeExcludedList;
    }

    public Boolean shouldIgnoreRefereeProcedure(
            FullyQualifiedVBAProcedureId procedureId) {
        for (ProcedureToBeIgnored pnbi : procedureNamesToBeIgnored) {
            if (pnbi.matches(procedureId)) {
                return true;
            }
        }
        return false;
    }

    public Boolean shouldExcludeModule(VBAModule module) {
        String moduleNameLowerCase = module.getName().toLowerCase();
        for (ModuleToBeExcluded pattern : moduleNamesToBeExcluded) {
            if (pattern.find(moduleNameLowerCase)) {
                return true;
            }
        }
        return false;
    }

    public Boolean shouldExcludeModule(VBAProcedureReference procedureReference) {
        return shouldExcludeModule(procedureReference.getReferrer().getModule());
    }

    /**
     *
     */
    public static class Builder {
        private final List<ProcedureToBeIgnored> procedureToBeIgnoredList;
        private final List<ModuleToBeExcluded> moduleToBeExcludedList;

        public Builder() {
            this.moduleToBeExcludedList = new ArrayList<>();
            this.moduleToBeExcludedList.add(ModuleToBeExcluded.STARTS_WITH_TEST);
            this.moduleToBeExcludedList.add(ModuleToBeExcluded.ENDS_WITH_TEST);
            //
            this.procedureToBeIgnoredList = new ArrayList<>();
            this.procedureToBeIgnoredList.add(ProcedureToBeIgnored.Class_Initialize);
            this.procedureToBeIgnoredList.add(ProcedureToBeIgnored.Class_Class_Initialize);
        }

        public Builder excludeModule(Pattern pattern) {
            moduleToBeExcludedList.add(
                    new ModuleToBeExcluded(pattern));
            return this;
        }

        public Builder noExcludeModule() {
            moduleToBeExcludedList.clear();
            return this;
        }

        public Builder ignoreRefereeProcedure(VBAModule.ModuleType type,
                                              String procedureName) {
            this.procedureToBeIgnoredList.add(
                    new ProcedureToBeIgnored(type, procedureName));
            return this;
        }

        public Builder noIgnoreRefereeProcedure() {
            this.procedureToBeIgnoredList.clear();
            return this;
        }

        public Options build() {
            return new Options(this);
        }
    }
}

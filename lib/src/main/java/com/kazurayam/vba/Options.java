package com.kazurayam.vba;

public class Options {

    public static Options DEFAULT = new Options.Builder().build();

    private final Boolean excludeUnittestModules;

    private Options(Builder builder) {
        this.excludeUnittestModules = builder.excludeUnitTestModules;
    }

    public Boolean getExcludeUnittestModules() {
        return excludeUnittestModules;
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
        private Boolean excludeUnitTestModules;
        public Builder() {
            this.excludeUnitTestModules = true;
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

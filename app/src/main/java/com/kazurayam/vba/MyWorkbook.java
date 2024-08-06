package com.kazurayam.vba;

import java.nio.file.Path;

public enum MyWorkbook {
    Backbone("Backboneライブラリ",
            "kazurayam-vba-lib",
            "office/kazurayam-vba-lib.xlsm",
            "office/exported-vba-source/kazurayam-vba-lib"),
    Member("Member会員名簿のためのVBAライブラリ",
            "aogan-vba-lib",
            "office/Member会員名簿のためのVBAライブラリ.xlsm",
            "office/exported-vba-source/Member会員名簿のためのVBAライブラリ"),
    Cashbook("Cashbook現金出納帳のためのVBAライブラリ",
            "aogan-vba-lib",
            "office/Cashbook現金出納帳のためのVBAライブラリ.xlsm",
            "office/exported-vba-source/Cashbook現金出納帳のためのVBAライブラリ"),
    Settlement("決算算出ワークブック",
            "aomori-gankaikai-jimukyoku",
            "office/決算算出ワークブック_令和5年度.xlsm",
            "office/exported-vba-source/決算算出ワークブック_令和5年度"),
    FeePaymentCheck("会費納入状況チェック",
            "aomori-gankaikai-jimukyoku",
            "office/会費納入状況チェック_青森県眼科医会_R6年度.xlsm",
            "office/exported-vba-source/会費納入状況チェック_青森県眼科医会_R6年度"),
    PleasePayFeeLetter("会費納入をお願いするletterを作成する",
            "aomori-gankaikai-jimukyoku",
            "office/会費納入のお願いletterを作成する.xlsm",
            "office/exported-vba-source/会費納入のお願いletterを作成する"),
    WebCredentials("会員名簿からIDパスワード管理情報を生成する",
            "aomori-gankaikai-jimukyoku",
            "office/会員名簿からIDパスワード管理情報を生成する.xlsm",
            "office/exported-vba-source/会員名簿からIDパスワード管理情報を生成する");

    private final String id;
    private final String repositoryName;
    private final String workbookSubPath;
    private final String vbaSourceDirSubPath;

    private MyWorkbook(String id, String repositoryName,
                       String workbookSubPath,
                       String vbaSourceDirSubPath) {
        this.id = id;
        this.repositoryName = repositoryName;
        this.workbookSubPath = workbookSubPath;
        this.vbaSourceDirSubPath = vbaSourceDirSubPath;
    }
    public String getId() {
        return id;
    }
    public String getRepositoryName() {
        return repositoryName;
    }
    public String getWorkbookSubPath() { return workbookSubPath; }
    public String getVbaSourceDirSubPath() {
        return vbaSourceDirSubPath;
    }
    public Path resolveWorkbookBasedOn(Path baseDir) {
        return baseDir.resolve(getRepositoryName()).resolve(getWorkbookSubPath());
    }
    public Path resolveVBASourceDirBasedOn(Path baseDir) {
        return baseDir.resolve(getRepositoryName()).resolve(getVbaSourceDirSubPath());
    }
}

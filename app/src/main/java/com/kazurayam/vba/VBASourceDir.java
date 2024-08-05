package com.kazurayam.vba;

import java.nio.file.Path;

public enum VBASourceDir {
    Backbone("Backboneライブラリ",
            "kazurayam-vba-lib",
            "office/exported-vba-source/kazurayam-vba-lib"),
    Member("Member会員名簿のためのVBAライブラリ",
            "aogan-vba-lib",
            "office/exported-vba-source/Member会員名簿のためのVBAライブラリ"),
    Cashbook("Cashbook現金出納帳のためのVBAライブラリ",
            "aogan-vba-lib",
            "office/exported-vba-source/Cashbook現金出納帳のためのVBAライブラリ"),
    Settlement("決算算出ワークブック",
            "aomori-gankaikai-jimukyoku",
            "office/exported-vba-source/決算算出ワークブック_令和5年度"),
    FeePaymentCheck("会費納入状況チェック",
            "aomori-gankaikai-jimukyoku",
            "office/exported-vba-source/会費納入状況チェック_青森県眼科医会_R5年度"),
    PleasePayFeeLetter("会費納入をお願いするletterを作成する",
            "aomori-gankaikai-jimukyoku",
            "office/exported-vba-source/会費納入のお願いletterを作成する"),
    WebCredentials("会員名簿からIDパスワード管理情報を生成する",
            "aomori-gankaikai-jimukyoku",
            "office/exported-vba-source/会員名簿からIDパスワード管理情報を生成する");

    private String id;
    private String repositoryName;
    private String subPath;

    private VBASourceDir(String id, String repositoryName, String subPath) {
        this.id = id;
        this.repositoryName = repositoryName;
        this.subPath = subPath;
    }
    public String getId() {
        return id;
    }
    public String getRepositoryName() {
        return repositoryName;
    }
    public String getSubPath() {
        return subPath;
    }
    public Path resolveBasedOn(Path baseDir) {
        return baseDir.resolve(repositoryName).resolve(subPath);
    }
}

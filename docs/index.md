# VBAProceduresIndexer

## このプロジェクトの背景

わたくし kazurayam は仕事としてMicrosoft Excelでアプリケーションをいくつか作った。ワークブック・ファイルに Visual Basic for Application 言語（以下、VBAと略記する）でコードを書いた。その仕事の中でkazurayamはいくつかVBAの難問にぶち当たった。この VBAProceduresIndexer プロジェクトはkazurayamが経験したVBAの難問のひとつを取り上げて解決することを目標としている。その難問とは何か？そこから説明を始めよう。

## 解決すべき問題

ひとつのアプリケーションを実現するのに、ひとつのxlsmファイルを作り、その中にすべてのVBAコード(SubやFunc)を格納するのが単純で実現しやすい。しかしアプリケーションを２つ、３つと開発してみるとすぐわかる。３つのアプリすべてで似たようなSubやFuncを使いたくなるのだ。ひとつのxlsmファイルにおいて書いたVBAコードをコピーして別のxlsmファイルに挿入するというやり方が可能ではある。しかしDon’t Repeat Yourselfの原則を破ることになる。プロフェッショナルなプログラマとして気持ち悪くてたまらない。だから共通モジュールを別個のxlsmファイルに括り出してライブラリ化した。ライブラリとしてのxlsmファイルからxlaファイルを生成した。これをMicrosoft用語ではAddinという。アプリケーションとしてのxlsmファイルがAddinの中のSubやFunctionを参照する構成にした。このやり方は珍しくもない。VBAプログラミングの教科書にちゃんと書いてる。そうやってkazurayamはライブラリとしてのxlsmファイルを３つとアプリケーションとしてのxlsmファイルを４つ作った。イメージを明確にするために具体的なファイル名を下記にあげておこう。

\![関連図\](<https://kazurayam.github.io/VBAProceduresIndexer/diagrams/out/ExcelFiles/事務局Excelブック関連図.png>)

`kazurayam_vba_lib` というのがGitHubレポジトリの名前で、その中に `Backboneライブラリ.xlsm` というExcelワークブックファイルがある、と了解してほしい。

さて、全部で7つのxlsmファイルがあって、その中にたくさんのモジュールやクラスがある。そしてそのモジュールがたくさんのSubとFunctionを宣言している。いったいいくつ、kazurayamはSubとFunctionを作ったろうか？ここに全部列挙してみよう。2024年7月末の時点でこうだ。

(TODO)

けっこうな数のSubとFuncをkazurayamが作ってしまったことを了解してほしい。 さて、ここからが本当の問題です。

kazurayamが7つのxlsmファイルの中に作り込んだたくさんのSubとFuncの大半は、アプリケーションの機能を実現するのに必要なものです。しかし全部が必要なものかといえばそうではない。ゴミが混じっているのだ。VBA初心者のkazurayamがVBAを習うためだけに書いたSubが残っている、とか。Subをひとつ作っている途中で名前を変更したくなったが、ちょっと訳があって、新しい名前のSubを別に作った、そして古いSubを削除するのを怠ったから残っている、とか。理由が何かを論じてもしょうがない。 **ゴミのSubを探し出して削除したい。しかしどれがゴミでどれがゴミでないか、どうやって見分ければいいのか?**

VBA言語の開発環境は貧弱で一つ一つのSubやFunctionがゴミであるかゴミでないかを見分ける手段を提供してくれない。困った。

## 解決方法

ツールを自作してゴミとゴミでないSubを見分ける手掛かりとなる情報を生成しようと思う。

## docs# VBAProceduresIndexer

# このプロジェクトの背景

わたくし kazurayam は仕事としてMicrosoft Excelでアプリケーションをいくつか作った。ワークブック・ファイルに Visual Basic for Application 言語（以下、VBAと略記する）でコードを書いた。その仕事の中でkazurayamはいくつかVBAの難問にぶち当たった。この VBAProceduresIndexer プロジェクトはkazurayamが経験したVBAの難問のひとつを取り上げて解決することを目標としている。その難問とは何か？そこから説明を始めよう。

# 解決すべき問題

ひとつのアプリケーションを実現するのに、ひとつのxlsmファイルを作り、その中にすべてのVBAコード(SubやFunc)を格納するのが単純で実現しやすい。しかしアプリケーションを２つ、３つと開発してみるとすぐわかる。３つのアプリすべてで似たようなSubやFuncを使いたくなるのだ。ひとつのxlsmファイルにおいて書いたVBAコードをコピーして別のxlsmファイルに挿入するというやり方が可能ではある。しかしDon’t Repeat Yourselfの原則を破ることになる。プロフェッショナルなプログラマとして気持ち悪くてたまらない。だから共通モジュールを別個のxlsmファイルに括り出してライブラリ化した。ライブラリとしてのxlsmファイルからxlaファイルを生成した。これをMicrosoft用語ではAddinという。アプリケーションとしてのxlsmファイルがAddinの中のSubやFunctionを参照する構成にした。このやり方は珍しくもない。VBAプログラミングの教科書にちゃんと書いてる。そうやってkazurayamはライブラリとしてのxlsmファイルを３つとアプリケーションとしてのxlsmファイルを４つ作った。イメージを明確にするために具体的なファイル名を下記にあげておこう。

<figure>
<img src="https://kazurayam.github.io/VBAProceduresIndexer/diagrams/out/ExcelFiles/事務局Excelブック関連図.png" alt="事務局Excelブック関連図" />
</figure>

`kazurayam_vba_lib` というのがGitHubレポジトリの名前で、その中に `Backboneライブラリ.xlsm` というExcelワークブックファイルがある、と了解してほしい。

さて、全部で7つのxlsmファイルがあって、その中にたくさんのモジュールやクラスがある。そしてそのモジュールがたくさんのSubとFunctionを宣言している。いったいいくつ、kazurayamはSubとFunctionを作ったろうか？ここに全部列挙してみよう。2024年7月末の時点でこうだ。

-   <https://kazurayam.github.io/VBAProcesuresIndexer/allVBASourceDirs.md>\]

けっこうな数のSubとFuncをkazurayamが作ってしまったことを了解してほしい。 さて、ここからが本当の問題です。

kazurayamが7つのxlsmファイルの中に作り込んだたくさんのSubとFuncの大半は、アプリケーションの機能を実現するのに必要なものです。しかし全部が必要なものかといえばそうではない。ゴミが混じっているのだ。VBA初心者のkazurayamがVBAを習うためだけに書いたSubが残っている、とか。Subをひとつ作っている途中で名前を変更したくなったが、ちょっと訳があって、新しい名前のSubを別に作った、そして古いSubを削除するのを怠ったから残っている、とか。理由が何かを論じてもしょうがない。 **ゴミのSubを探し出して削除したい。しかしどれがゴミでどれがゴミでないか、どうやって見分ければいいのか?**

VBA言語の開発環境は貧弱で一つ一つのSubやFunctionがゴミであるかゴミでないかを見分ける手段を提供してくれない。困った。

# 解決方法

ツールを自作してゴミとゴミでないSubを見分ける手掛かりとなる情報を生成しようと思う。

# docs

<https://kazurayam.github.io/VBAProceduresIndexer/>

<https://kazurayam.github.io/VBAProceduresIndexer/>

# Lorem ipsum

Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.

## Ut enim ad minim veriam

Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.

### Duis aute irure dolor

Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.

#### Excepteur sint occaecat cupidatat

Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.

# Quo usque tandem abutere

Quo usque tandem abutere, Catilina, patientia nostra?

# Header with (& - \_ : ; ) non Latin Characters

Quam diu etiam furor iste tuus nos eludet?

# Source code inclusion example

    #!/bin/bash

    # Under the current working directory, convert all the files
    # with name ending with `_.adoc` into `.md`.
    # A file with name `*_.adoc` is an Asciidoc document file.
    # A file with name `*.md` is a Markdown document file.
    # E.g, `index_.adoc` will be converted into `index.md`
    #
    # However, ffiles with name which starts with with `_` will be ignored.
    # E.g, `_index_.adoc` will be ignored.
    #
    # How to run this: in the command line, just type
    #
    #     `> ./adoc2md.sh`
    #
    # By specifying `-t` option in the command line,
    #
    #     `> ./adoc2md.sh -t`
    #
    # you can prepend a text segment into the output .md file:
    #
    # ```
    # - Table of Contents
    # {:toc}
    #
    # ```
    # This text segment will affect the [Jekyll](https://docs.github.com/en/pages/setting-up-a-github-pages-site-with-jekyll/about-github-pages-and-jekyll)
    # used by GitHub Pages to generate a table of contents.
    #

    requireTOC=false

    optstring="t"
    while getopts ${optstring} arg; do
      case ${arg} in
        t)
            requireTOC=true
            ;;
        ?)
            ;;
      esac
    done

    function processFile() {
      fname=$1
      #echo "fname=${fname}"
      #  using Asciidoctor, convert a *.adoc file into a docbook in XML
      md=${fname//adoc/md}
      xml=${fname//adoc/xml}
      echo "converting $fname into $md"
      asciidoctor -b docbook -a leveloffset=+1 -o - "$fname" > "$xml"
      # using Pandoc, generate a Markdown file without TOC
      cat "$xml" | pandoc --markdown-headings=atx --wrap=preserve -t markdown_strict -f docbook - > "$md"
      #echo deleting $xml
      rm -f "$xml"

      # We named `index_.adoc` rather than `index.adoc` because GitHub puts precedence to `index.adoc` over `index.md`. We want `index.md` to be presented first, not `*.adoc`. Therefore we named our adoc file with `*_.adoc` postfix.
      # This trick required further treatment.
      # `index_.adoc` will result `index_.adoc`. But we
      # want the final result to be `index.md`.
      # So, we will rename `*_.md` into `*.md`.
      # in other words, chomp an underline character (_) before `.md``
      # e.g,
      #   ./index_.adoc    -> ./index.md
      #   ./index-ja_.md -> ./index-ja.md
      newmd=${md%_.md}.md
      echo renaming $md to $newmd
      mv $md $newmd

      # Slightly modify the generated *.md file.
      # Prepend a text segment:
      # ```
      # - Table of contents
      # {:toc}
      # ```
      if [ $requireTOC = true ]; then
        echo "- Table of contents" > temp.md
        echo "{:toc}" >> temp.md
        echo "" >> temp.md
        cat $newmd >> temp.md
        cat temp.md > $newmd
        rm temp.md
        echo prepended the TOC segement in $newmd
      fi
      # just a blank line to separate the *.adoc files processed
      echo ""
    }



    # iterate over all *.adoc files
    find . -iname "*_.adoc" -type f -maxdepth 1 -not -name "_*.adoc" | while read fname; do
      processFile $fname
    done

# Diagram example

Generated using [PlantUML](https://plantuml.com/ja/)

<figure>
<img src="https://kazurayam.github.io/adoc2md/diagrams/out/sequence/sequence.png" alt="sequence" />
</figure>

# 宮沢賢治の詩その1

心象スケッチ　春と修羅

    わたくしといふ現象は
    仮定された有機交流電燈の
    ひとつの青い照明です
    （あらゆる透明な幽霊の複合体）
    風景やみんなといつしよに
    せはしくせはしく明滅しながら
    いかにもたしかにともりつづける
    因果交流電燈の
    ひとつの青い照明です
    （ひかりはたもち　その電燈は失はれ）

    これらは二十二箇月の
    過去とかんずる方角から
    紙と鉱質インクをつらね
    （すべてわたくしと明滅し
    　みんなが同時に感ずるもの）
    ここまでたもちつゞけられた
    かげとひかりのひとくさりづつ
    そのとほりの心象スケッチです

# 宮沢賢治の詩その2

    雨ニモマケズ
    風ニモマケズ
    雪ニモ夏ノ暑サニモマケヌ
    丈夫ナカラダヲモチ
    欲ハナク
    決シテ瞋ラズ
    イツモシヅカニワラッテヰル
    一日ニ玄米四合ト
    味噌ト少シノ野菜ヲタベ
    アラユルコトヲ
    ジブンノカンジョウニ入レズニ
    ヨクミキキシワカリ
    ソシテワスレズ
    野原ノ松ノ林ノ蔭ノ
    小サナ萱ブキノ小屋にヰテ
    東に病気ノコドモアレバ
    行ッテ看病シテヤリ
    西ニ疲レタ母アレバ
    行ッテソノ稲ノ束ヲ負ヒ
    南ニシニサウナ人アレバ
    行ッテコハガラナクテモイゝトイヒ
    北にケンクヮヤソショウガアレバ
    ツマラナイカラヤメロとイヒ
    ヒドリノトキハナミダヲナガし
    サムサノナツハオロオロアルキ
    ミンナニデクノボー トヨバレ
    ホメラレモセズ
    クニモサレズ
    サウイフモノニ
    ワタシハナリタイ

    南無無辺行菩薩
    南無上行菩薩
    南無多宝如来
    南無妙法蓮華経
    南無釈迦牟尼仏
    南無浄行菩薩
    南無安立行菩薩

# Dolorem ipsum

Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Sed tempus urna et pharetra pharetra massa. Praesent tristique magna sit amet purus. Fermentum et sollicitudin ac orci phasellus egestas tellus rutrum tellus. In cursus turpis massa tincidunt dui ut. Massa tincidunt nunc pulvinar sapien et ligula ullamcorper. Facilisis mauris sit amet massa vitae tortor condimentum. Sapien faucibus et molestie ac feugiat sed lectus vestibulum. Eu turpis egestas pretium aenean pharetra magna ac placerat. Nulla aliquet porttitor lacus luctus accumsan tortor posuere. Sit amet purus gravida quis blandit turpis cursus. Suspendisse sed nisi lacus sed. Laoreet non curabitur gravida arcu ac. Dolor sit amet consectetur adipiscing elit pellentesque. Egestas quis ipsum suspendisse ultrices gravida dictum fusce ut placerat.

Praesent tristique magna sit amet purus gravida quis blandit turpis. Vitae congue mauris rhoncus aenean vel elit scelerisque mauris. Suspendisse faucibus interdum posuere lorem ipsum dolor sit. Sem integer vitae justo eget magna fermentum iaculis eu. Felis bibendum ut tristique et egestas. Vel orci porta non pulvinar neque laoreet. Ac tortor vitae purus faucibus ornare suspendisse sed. In mollis nunc sed id semper. Pellentesque id nibh tortor id aliquet lectus proin. Id faucibus nisl tincidunt eget nullam non nisi est. Senectus et netus et malesuada fames ac turpis egestas. Viverra vitae congue eu consequat ac. Vel pharetra vel turpis nunc eget lorem dolor sed viverra. Sodales ut eu sem integer vitae. Pharetra vel turpis nunc eget lorem. Nam at lectus urna duis convallis convallis tellus. Massa tempor nec feugiat nisl pretium fusce id velit ut. Et ultrices neque ornare aenean euismod elementum nisi. Ut morbi tincidunt augue interdum velit. Fermentum iaculis eu non diam phasellus vestibulum lorem sed risus.

Et sollicitudin ac orci phasellus egestas tellus. Sit amet purus gravida quis blandit turpis cursus. Interdum varius sit amet mattis vulputate enim nulla aliquet. Quam id leo in vitae turpis massa sed elementum. Justo donec enim diam vulputate ut pharetra sit amet aliquam. Mauris commodo quis imperdiet massa tincidunt nunc pulvinar sapien. Leo integer malesuada nunc vel risus commodo viverra maecenas. Neque viverra justo nec ultrices dui sapien eget mi. Vel risus commodo viverra maecenas. Purus in massa tempor nec feugiat nisl pretium fusce. In nisl nisi scelerisque eu ultrices vitae. Nulla aliquet porttitor lacus luctus accumsan tortor posuere ac ut. Elit ut aliquam purus sit amet luctus venenatis lectus. Ultrices sagittis orci a scelerisque purus semper. Aliquam sem et tortor consequat id. Sed vulputate odio ut enim blandit volutpat maecenas volutpat. Sed euismod nisi porta lorem mollis aliquam.

Faucibus ornare suspendisse sed nisi lacus sed viverra. Duis at tellus at urna. Sem fringilla ut morbi tincidunt augue interdum. Proin libero nunc consequat interdum varius sit amet. Elementum facilisis leo vel fringilla est. In fermentum posuere urna nec tincidunt. Urna id volutpat lacus laoreet non curabitur gravida arcu ac. Arcu ac tortor dignissim convallis aenean et tortor. Odio facilisis mauris sit amet massa vitae tortor condimentum. At tellus at urna condimentum mattis pellentesque. Sed tempus urna et pharetra pharetra massa. Gravida quis blandit turpis cursus in hac. Enim neque volutpat ac tincidunt vitae semper quis lectus nulla. Egestas sed tempus urna et pharetra pharetra massa massa ultricies.

Aliquam purus sit amet luctus venenatis lectus magna. Vitae ultricies leo integer malesuada. Sit amet facilisis magna etiam tempor orci eu. Nec nam aliquam sem et tortor consequat id porta. Est ante in nibh mauris. Eu feugiat pretium nibh ipsum consequat nisl vel. Tristique risus nec feugiat in fermentum posuere urna nec tincidunt. Mauris vitae ultricies leo integer malesuada nunc vel risus commodo. Ultrices neque ornare aenean euismod elementum nisi quis eleifend quam. Lacus vel facilisis volutpat est velit egestas. Non curabitur gravida arcu ac tortor dignissim convallis aenean et. Volutpat est velit egestas dui id ornare arcu odio ut. Venenatis a condimentum vitae sapien pellentesque. Amet nisl suscipit adipiscing bibendum est ultricies integer. Lectus quam id leo in. Non curabitur gravida arcu ac tortor dignissim convallis aenean.

# VBAソースのサンプルを作るためのサブプロジェクト --- office 


VBAProceduresIndexerプロジェクトを開発するにあたってVisual Basic for Application言語で書かれたプログラムコードが必要だった。このプロジェクトはVBAのソースコードを入力として読み、構文解析してSubやFunctionの名前を抽出しようとする。最終的にはそれらProcedure名をノードとして矢印で結んだ有向グラフを出力しようとする。このプログラムをテストするためにはサンプルとしてのVBAコードが必要だった。そこでインターネットで見つけた教材からVBAつきExcelブックsを写経して利用することにした。

- https://tonari-it.com/excel-vba-class-addin/
- https://tonari-it.com/excel-vba-class-addin-reference/

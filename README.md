### 温度グラフの生成
---
#### ファイルの準備
テキストファイルをダウンロードしtemp_graph直下に格納する。

1.Marge.batを実行しファイルをマージする

（マージ前のファイルはdatフォルダに格納される）

<img src="https://github.com/TA1851/temp_graph/blob/main/img/marge.PNG">

---
#### 処理の流れ 
<img src="https://github.com/TA1851/temp_graph/blob/main/img/%E3%83%95%E3%83%AD%E3%83%BC.PNG">

---
csv_create.xlsm -> Marge.logをデータ整形し、csv fileに変換する

TempGraph_Single.xlsm -> グラフの作成

single_macroを実行する。

* 2種類の処理を追加

ダイアログから拡張子がcsvのファイルを読み込む
ファイルパスを指定しグラフを生成

* 2種類の処理を追加

Graph1の実行 -> 折れ線グラフを生成する
Graph2の実行 -> 散布図を生成する
  
---
Requirement

* Library list
* Visual Basic For Applicasions
* Microsoft Excel 16.0 Object Library
* OLE Automation
* Microsoft Office 16.0 Object Library

---
Note

Source code backup 取得は、VBCAを採用
[参考記事](https://tonari-it.com/vba-vbac-git/)

GitHubにPushした際に、bin folderが表示されなかった為、Pushする時に、foler nameをbin > mainに変更してPushした。

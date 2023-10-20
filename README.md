### 温度グラフの生成
---
Flow

<img src="https://github.com/TA1851/temp_graph/blob/main/img/marge.PNG">

File 構成

<img src="https://github.com/TA1851/temp_graph/blob/main/img/%E3%83%95%E3%83%AD%E3%83%BC.PNG">

* Marge.bat -> File Margeとdat フォルダの作成
  * server からDownloadしてきたファイルを新規作成したフォルダに格納し、Marge.batを実行する。
  * Marge.batを実行すると、温度log(.dat)をmargeしてMarge.logを生成してdat フォルダを生成する。
  * dat フォルダには、Downloadしてきたファイルを格納する。

* csv_create.xlsm -> Marge.logをデータ整形し、csv fileに変換する

---
#### Graph Create

---
Flow

<img src="https://github.com/TA1851/ACR_Single_TempGraph/blob/main/img/flow2.PNG">

* ACR_TempGraph_Single.xlsm
  * single_macro -> ダイアログから拡張子がcsvのファイルを読み込む
  * writing1 -> 区切り文字を置換して、ファイルに書き出す
  * writing2 -> 不要なデータの削除と温度の計算を行い、グラフ作成用データにする
  * Graph1 -> 折れ線グラフを生成する
  
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


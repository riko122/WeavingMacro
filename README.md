# WeavingMacro
このツールは、織物用の組織図と配色図を作るものです。Microsoft Excel VBAで稼働します。

綜絖とタイアップと踏み木を黒くして実行すると組織図が作成されます。  
更に、経糸と緯糸に使用する色を指定して実行すると配色図が作成されます。

## Requirement
- Microsoft Windows
- Microsoft Excel

### 動作確認済の環境
- Windows 8.1 + Excel 2013
- Windows 10 + Excel 2000
- Windows 10 + Excel 2010
- Windows 10 + Excel 2013
- Windows 10 + Excel 2016
- MacOS Mojave 10.14.6 + Excel for Mac 2019

他の環境でも動いたら、ここに追記しますので、commitへのコメントなどでご一報いただけると有難いです。
2020/3/8にMacOS Mojaveでの動作報告をいただいたので追加しました。


なお、LibreOfficeで動くものを、odsフォルダ以下に置きました。 
LibreOffice版なら、Mac OSでも動くのではないかと期待しています。
但し、LibreOffice版は、フォントサイズや配色図での罫線において不具合があります。

## Thanks

Gitでソースコード管理をするため、igetaさんのAriawase(vbac.wsf)を使用して、VBAのエクスポート・インポートをしています。
Ariawaseのライセンスは以下の通りです。  
https://github.com/vbaidiot/Ariawase#license


## Setup
「Code」から「Download ZIP」をクリックしてZIPファイルをDL後解凍してください。  
binフォルダ内にxlsmファイルが入っています。
古いExcel用にxlsファイル(V3.8.2)も入っていますが、基本的にはxlsmファイル(V3.9)をご利用ください。
今後、保守はxlsmファイルの方のみを行う予定です。

## Usage

※　画像はV3.8より前のもののため、設定項目のセル位置が最新版と異なります。

#### 1.初期化
踏み木本数と、綜絖枚数と、タイアップ部分をどの位置にするかと、図の幅・図の高さを入力して[初期化]ボタンをクリックして、マス目を作ります。

#### 2.準備
踏むと綜絖が上がるか下がるかを「↑」と「↓」から選びます。
以下の図のように、綜絖とタイアップと踏み木を黒くします。  
![ドラフト3-1](https://blog-imgs-95.fc2.com/r/i/k/riko122/img652_draft3-1.png)

#### 3.組織図
「組織図」ボタンをクリックすると   
![ドラフト3-2](https://blog-imgs-95.fc2.com/r/i/k/riko122/img653_draft3-2.png)  
こんな風に、どんな組織になるかが表示されます。

#### 4.配色図
以下の図のように、色を指定して、  
![ドラフト3-3](https://blog-imgs-95.fc2.com/r/i/k/riko122/img654_draft3-3.png)  

「配色図」ボタンをクリックすると   
![ドラフト3-4](https://blog-imgs-95.fc2.com/r/i/k/riko122/img655_draft3-4.png)  
こんな風に、どんな配色になるのかが表示されます。

#### お願い
簡単なテストはしていますが、実際に使ってみると不具合が生じるかもしれません。その際はIssuesへの書き込みでご連絡下さい。
可能な範囲で修正したいと思います。  
要望などもIssuesでお知らせください。検討します。

## Licence
This software is released under the Mozilla Public License 2.0, see LICENSE.

Mozilla Public License 2.0の下、使用にあたっては制限はありません。  
改変や第三者への配布を行う場合に、制限が生じます。詳細はLICENCEをご確認ください。  
また、出版物に載せるなどの場合はcommitにコメントをつけるなどでご一報ください。

なお、個人が生成した図を画像にしてサイトに載せる場合はご連絡くださらなくてけっこうです。
ただし、本に載っていたものから完全意匠図を作成した場合、マクロではなく、その本の著作権に抵触しないか、十分お気を付けください。

## Authors

[Riko](https://github.com/riko122)

## History

- 2021/ 1/ 7 V3.8.2 印刷時のフッターをコピーライト表記からURLに変更
- 2021/ 1/ 6 V3.9.1 コピーライトを追記
- 2020/ 7/12 V3.9   ファイル形式をxlsmで保存
- 2020/ 7/12 V3.8.2 「緯糸の色↓」の矢印が下を向くように修正
- 2019/ 5/ 3 LibreOffice版のV3.8.2oもGitHubでの管理下に追加
- 2019/ 5/ 1 V3.8.1 マクロの内部的な処理変更
- 2019/ 4/20 V3.8   「↑又は↓」を選ぶセルの位置の変更
- 2019/ 2/ 5 V3.7.1 よりGitHubでの管理開始

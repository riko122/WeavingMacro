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

他の環境でも動いたら、ここに追記しますので、commitへのコメントなどでご一報いただけると有難いです。

なお、LibreOfficeで動くものは別途用意します。  
LibreOffice版なら、Mac OSでも動くのではないかと期待しています。

## Thanks

Gitでソースコード管理をするため、igetaさんのAriawase(vbac.wsf)を使用して、VBAのエクスポート・インポートをしています。
Ariawaseのライセンスは以下の通りです。  
https://github.com/vbaidiot/Ariawase#license


## Setup
「Clone or download」から「Download ZIP」をクリックしてZIPファイルをDL後解凍してください。  
binフォルダ内にxlsファイルが入っています。

## Usage

#### 1.初期化
踏み木本数と、踏むと綜絖が上がるか下がるかと、綜絖枚数と、タイアップ部分をどの位置にするかと、図の幅・図の高さを入力して[初期化]ボタンをクリックして、マス目を作ります。

#### 2.準備
以下の図のように、綜絖とタイアップと踏み木を黒くします。  
![ドラフト3-1](http://sky.geocities.jp/riko_21_riko/misc_img/img652_draft3-1.PNG)

#### 3.組織図
「組織図」ボタンをクリックすると   
![ドラフト3-2](http://sky.geocities.jp/riko_21_riko/misc_img/img653_draft3-2.PNG)  
こんな風に、どんな組織になるかが表示されます。

#### 4.配色図
以下の図のように、色を指定して、  
![ドラフト3-3](http://sky.geocities.jp/riko_21_riko/misc_img/img654_draft3-3.PNG)  

「配色図」ボタンをクリックすると   
![ドラフト3-4](http://sky.geocities.jp/riko_21_riko/misc_img/img655_draft3-4.PNG)  
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

- 2019/2/5 V3.7.1 よりGitHubでの管理開始

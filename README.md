K# GetThisWorkbookLocalPath
## OneDriveでThisWorkbook.PathがURLを返す問題を解決する    
  
### 解決したい問題 ### 
  
OneDrive上のExcel VBAを動かすとThisWorkbook.PathがURLを返す問題が起きます。自分自身のローカルパスを取得できず、FileSystemObjectまで使えなくなるという不便な状態になります。  
  
この問題の解決にはいくつかの方法が提案されていますが、URLパスを文字列処理してローカルパスに変換する方法はうまくいかない場合があります。特に OneDrive for Business においてはURLに含まれるテナントコードをテナント名に変換するなどの処理が必要で、文字列処理による方法では解決できません。  
  
TeamsやSharePointファイルを同期するには「同期クライアント」と「OneDriveへのショートカットの追加」の二つの方法がありますが、それぞれローカルドライブ上のパスが異なり、さらに、どちらの方法で同期されているかをURLパスから知ることはできません。  
  
このような理由からThisWorkbook.Pathが返すURLを文字列処理によってローカルパスに変換する方法には事実上無理があります。

### 提案する解決策 ###

ここでは異なる三つの方法を提案します。
一つ目は「最近開いた項目の表示」を利用するもの、二つ目は開いているエクスプローラーを利用するもの、三つ目はSendKeysを利用するものです。
それぞれ前提条件がありますので、単独もしくは必要に応じて組み合わせて利用するとよいでしょう。

ソースコードは標準モジュールをエクスポートしたファイルをそのまま掲載していますので、インポートするか、必要な部分をコピペしてお使いください。  
  
異なる三つの方法を以下の３つのファイルで掲載しています。  
Module1.bas　「最近開いた項目の表示」を利用する方法  
Module2.bas　開いているエクスプローラーを利用する方法  
Module3.bas　SendKeysを利用する方法  

### 提案する解決策 （その１） ###  
  
ソースコードはModule1.basです。ローカルパスを取得する関数は GetThisWorkbookLocalPath1() です。

このコードは「最近開いた項目の表示」を利用するもので、最近開いたファイルやフォルダーが
  
    C:\Users\<user-name>\AppData\Roaming\Microsoft\Windows\Recent  
  
のフォルダーにリンクファイル（LNKファイル）として自動的に記録される機能を利用しています。このリンクファイルのリンク先を取得することでローカルドライブ上のパスを得ることができます。 
  
「最近開いた項目の表示」を利用するためには、Windowsの設定で「個人用設定」→「スタート」で「最近開いた項目の表示」をオンにします。

Windows 11 の場合は「最近開いた項目をスタート、ジャンプ リスト、ファイル エクスプローラーに表示する」

Windows 10 の場合は、「スタート メニューまたはタスク バーの、ジャンプ リストとエクスプローラーのクイック アクセスに最近開いた項目を表示する」

となっています。この設定がオフの場合は、上述のリンクファイル（LNKファイル）が記録されないため、GetThisWorkbookLocalPath1() は空文字（長さゼロの文字列）を返します。
  
### レジストリキーの読み出し ###  
  
GetThisWorkbookLocalPath1() を呼び出す前に「最近開いた項目の表示」をオンになっているかどうかを知るにはレジストリキーを読んで調べます。そのための関数がIs_Start_TrackDocs() です。  
  
この関数はレジストリキー  

    HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\  

にある Start_TrackDocs の値を読んで、その値を返します。

### 提案する解決策 （その２） ###  
  
ソースコードはModule2.basです。ローカルパスを取得する関数は GetThisWorkbookLocalPath2() です。

このコードは現在開いているExcelファイル（つまりThisWorkbook）が置かれているフォルダーを表示しているエクスプローラーからローカルパスを取得します。

具体的には、エクスプローラーのウインドウからのWindowオブジェクトを取得しLocationURLプロパティで"file:///C:/Users/～/～/OneDrive～/～"という絶対パス（URI）を取得します。

この絶対パス（URI）はエンコードされているのでデコードします。DecodeURL() 関数はそのためのものです。エンコードされるのは特定のASCII文字だけなので、簡易版のDecodeURL_ASCII() 関数も参考に記述しました。

このようにGetThisWorkbookLocalPath2() はエクスプローラーから情報を得ていますので、該当するエクスプローラーを閉じてしまうと情報が得られなくなります。この場合、GetThisWorkbookLocalPath2() は空文字（長さゼロの文字列）を返します。

なお、OneDrive、OneDrive for Businessの直下（ルートフォルダー）に置かれている場合、ThisWorkbook.Pathはそれぞれ特定のURLパターンを返すので、エクスプローラーから情報を得ずとも、OneDriveにはEnviron("OneDrive")、OneDrive for BusinessにはEnviron("OneDriveCommercial")のローカルパスを対応させています。

### 提案する解決策 （その３） ###  
  
ソースコードはModule3.basです。ローカルパスを取得する関数は GetThisWorkbookLocalPath3() です。

このコードは現在開いているExcelファイル（つまりThisWorkbook）自身にSendKeysによってキーストロークを送ってローカルパスを取得します。

OneDrive上に置かれたExcelファイルは「ファイル(F)」タブ→「情報(I)」→「ローカルパスのコピー(L)」でローカルパスをクリップボードに取得できます。

キーストロークは「Alt」→「F」→「I」→「L」となります。このあと元のホームタブに戻るため「Alt」→「H」→「↑」→「Enter」を送信します。

したがって実際のキーストロークは「Alt」→「F」→「I」→「L」→「Alt」→「H」→「↑」→「Enter」となります。

SendKeysはVBAのApplication.SendKeysメソッドは使えません。自分自身のリボンタブの操作にApplication.SendKeysメソッドはうまく機能しないためです。

この問題はPowerShellによって外部からExcelにキーストロークを送ることで解決できます。PowerShellで実行するキーストロークを送るスクリプトはコードの中に埋め込んでいます。

実は元のホームタブに戻るために「Esc」キーを送信したいところですが、タイミングによってはVBAが中断される場合があり、「Esc」キーの送信は避けています。

キーストロークの送信タイミングはスクリプトの中のStart-Sleepコマンドレットで指定しています。余裕のあるタイミングにしていますが、WindowsやOffice環境によってはStart-Sleepのタイミングを調整する必要があるかもしません。

キーストロークの送信によってウインドウが切り替わりますが正常な動作ですのでご承知おきください。キーストロークの送信に失敗した場合、GetThisWorkbookLocalPath3() は空文字（長さゼロの文字列）を返します。

もともとのソースコードでThisWorkbook.Pathを多用している場合、単純にThisWorkbook.PathをGetThisWorkbookLocalPath3()に置き換えると頻繁に画面が動くので、グローバル変数に取り込むなどしてGetThisWorkbookLocalPath3()の呼び出し回数をできるだけ少なくするとよいでしょう。

### 最後に ###

OneDrive、OneDrive for Business、またはTeamsやSharePointを「OneDriveへのショートカットの追加」によってローカルドライブとして利用できます。これはWebアクセスを意識せずに利用できるメリットがあります。  
一方でこれらの新しい仕組みに対してVBAは非力です。今回の提案はそれを補う一つの方法ですが、そもそもVBAは2012年を最後に大きなアップデートはなくMicrosoftが提案する新しいソリューションに対して置き去りにされた感は否めません。  
  
ThisWorkbook.PathがURLを返す問題は解決されたとしても、SharePointを「OneDriveへのショートカットの追加」でファイルを利用するケースでは組織の関係者が共有しているという関係から、CheckOut／CheckInといった排他制御が必要になる場合があります。  
  
もちろんVBAにはCheckOut／CheckInのメッソドがありますが、リトライ処理を含むフロー制御が必要になり単純ではありません。
その意味から今回の提案は、他に解決方法がない場合の暫定的な手段と捉えるべきでしょう。

### ライセンス ###

このコードはMITライセンスに基づき利用できます。 

■    

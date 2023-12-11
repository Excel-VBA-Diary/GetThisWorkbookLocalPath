# GetThisWorkbookLocalPath
## OneDriveでThisWorkbook.PathがURLを返す問題を解決する    
  
### 解決したい問題 ### 
  
OneDrive上のExcel VBAを動かすとThisWorkbook.PathがURLを返す問題が起きます。自分自身のローカルパスを取得できず、FileSystemObjectまで使えなくなるという不便な状態になります。    
この問題の解決にはいくつかの方法が提案されていますが、URLパスを文字列処理してローカルパスに変換する方法はうまく処理できない場合があります。特に OneDrive for Business においてはURLに含まれるテナント名などの解決はほぼ不可能です。
SharePointファイルを同期するには「同期クライアント」と「OneDriveへのショートカットの追加」の二つの方法があります。それぞれローカルドライブ上のパスが異なります。どちらの方法で同期されているかをURLパスから知ることはできません。
このような理由からThisWorkbook.Pathが返すURLを文字列変換によってローカルパスに変換する方法には事実上無理があります。  

### 提案する解決策 （その１）###  
  
ここで紹介する方法は「最近開いた項目の表示」を利用するもので、最近開いたファイルやフォルダーが
  
    C:\Users\<user-name>\AppData\Roaming\Microsoft\Windows\Recent  
  
のフォルダーにリンクファイル（LNKファイル）として自動的に記録される機能を利用してます。このリンクファイルのリンク先を取得することでローカルドライブ上のパスを得ることができます。    
  
ローカルパスを取得する関数は GetThisWorkbookLocalPath1() です。
  
### レジストリーキーの読み出し ###  
  
ThisWorkbookLocalPathは2つのレジストリキーを読んで処理を実行しています。そのための関数がStart_TrackDocs()とHideFileExt()です。  
  
### Start_TrackDocs() ###
  
この関数はレジストリキー  

    HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\  

にある Start_TrackDocs の値を読んで、その値を返します。  
「最近使った項目の表示」を利用するには、「Windowsの設定」⇒「個人用設定」⇒「スタート」の「～に最近開いた項目の表示する」または「最近開いた項目を～に表示する」が有効である必要があります。  
「～最近使った項目の表示」が有効か無効かは Start_TrackDocs の値を読んで判別できます。  
  
なお、Windowsの環境によってはStart_TrackDocsキーが設定されていない場合があります。その場合は「～に最近開いた項目の表示する」または「最近開いた項目を～に表示する」を一旦無効にしてから有効にしてください。  
  
  
■    

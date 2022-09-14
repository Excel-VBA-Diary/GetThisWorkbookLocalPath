# ThisWorkbookLocalPath
## OneDriveでThisWorkbook.PathがURLを返す問題を解決する    
### Resolve the problem of ThisWorkbook.Path returning a URL in OneDrive ###   
  
  
### 解決したい問題 (Problem to be solved) ### 
  
OneDrive上のExcel VBAを動かすとThisWorkbook.PathがURLを返す問題が起きます。自分自身のローカルパスを取得できず、FileSystemObjectまで使えなくなるという不便な状態になります。    
この問題の解決にはいくつかの方法が提案されていますが、URLパスを文字列処理してローカルパスに変換する方法は実際には使えません。  
SharePointファイルを同期するには「同期クライアント」と「OneDriveへのショートカットの追加」の二つの方法があります。それぞれローカルドライブ上のパスが異なります。どちらの方法で同期されているかをURLパスから知ることはできないことや、OneDriveのフォルダー名を変更できることから、URLパスからローカルパスに変換することは事実上無理があります。  

### 提案する解決策 (Proposed Solution) ###  
  
ここで紹介する方法は「最近使った項目の表示」を利用するもので、最近使ったファイルやフォルダーが
  
C:\Users\\<user-name>\AppData\Roaming\Microsoft\Windows\Recent  
  
のフォルダーにリンクファイル（LNKファイル）として自動的に記録される機能を利用してます。このリンクファイルのリンク先を取得することでローカルドライブ上のパスを得ることができます。    
  
ローカルパスを取得する関数のは ThisWorkbookLocalPath() で、OneDriveに同期したSharePointファイルのローカルドライブ上のパスを返します。つまり、ThisWorkbookLocalPath()はThisWorkbook.Pathに置き換えて使うことができます。 このマクロを任意のOneDrive上のフォルダーに置き、そのまま起動することができます。
この関数をテストするコードは Test_ThisWorkbookLocalPath() です。  
  
### Start_TrackDocs() ###
  
「最近使った項目の表示」を利用するには、「Windowsの設定」⇒「個人用設定」⇒「スタート」の「～最近使った項目の表示」が有効である必要があります。  
この「～最近使った項目の表示」が有効か無効かはレジストリを読んで判断できます。その関数が Start_TrackDocs() です。  
  
### HideFileExt() ###
  
フォルダーオプションの「登録されている拡張子は表示しない」が有効か無効かによってリンクファイルの名前が変わります。
有効である場合、マクロ付きExcelファイルならリンクファイルは「～.xlsm.LNK」となりますが、無効である場合は「～.LNK」となります。  
この判別もレジストリを読んで判断しています。その関数が HideFileExt() です。    
  
■    

# ThisWorkbookLocalPath
## OneDriveに同期したSharePointファイルのローカルドライブ上のパスを取得する    
**(Returns the path on the local drive of SharePoint files synced to OneDrive)** 
  
  
**解決したい問題**  
  
OneDrive上のExcel VBAを動かすとThisWorkbook.PathがURLを返す問題に遭遇する。  
この問題の解決にはいくつかの方法が提案されているが、URLパスを文字列処理してローカルパスに変換する方法は使えない。  
SharePointファイルを同期するには「同期クライアント」と「OneDriveへのショートカットの追加」の二つの方法がある。  
それぞれローカルドライブ上のパスは異なる。どちらの方法で同期されているかをURLパスから知ることはできない。

**提案する解決策**  
  
ここで紹介する方法は「最近使った項目の表示」を利用するもので、最近使ったファイルやフォルダーが
  
C:\Users\\<ユーザ名>\AppData\Roaming\Microsoft\Windows\Recent  
  
のフォルダーにリンクファイル（LNKファイル）として自動的に記録される機能を利用している。 このリンクファイルのリンク先を取得することでローカルドライブ上のパスを得る。  
  
関数の本体は ThisWorkbookLocalPath() でOneDriveに同期したSharePointファイルのローカルドライブ上のパスを返す。 
この関数をテストするコードは Test_ThisWorkbookLocalPath() 。  
このマクロを任意のSharePointフォルダーに置き、そのフォルダーと同時しているOneDrive上のフォルダーでマクロを起動すればよい。
  
**遭遇した課題と解決(1)**
  
「最近使った項目の表示」を利用するには、Windows設定⇒個人用設定⇒スタートの「～最近使った項目の表示」が有効である必要がある。  
この「～最近使った項目の表示」が有効か無効かはレジストリを読んで判断できる。（Start_TrackDocs()）  
  
**遭遇した課題と解決(2)**
  
フォルダーオプションの「登録されている拡張子は表示しない」が有効か無効かによってリンクファイルの名前も変わる。
有効である場合、マクロ付きExcelファイルならリンクファイルは「～.xlsm.LNK」となるが、無効である場合は「～.LNK」となる。  
この判別もレジストリを読んで判断している。（HideFileExt()）    
  
**利用にあたって**
  
ライセンスは自由度の高いMITライセンスとした。  
  

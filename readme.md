*これはVBScriptからExcelの複数の引数を取るマクロを呼び出すサンプルです。*  
*vbsファイルの呼び出し元に戻り値を返します。*  
*VC++からのvbsファイル実行のサンプルも載せてあります。*  

# 目的  
Excelの世界の事はExcelで閉じる事。  
Excelの帳票を出力するようなプログラムを作成する場合に、C++やC#からCOMでExcelを直接いじることを  
避ける事ができるので、側の都合（バージョン、帳票フォーマット変更など）によって  
プログラムのコードを直さなくて済むようになるので、メンテナンス性が向上します。  


/template:引数で指定したExcelファイルを開き、  
/macro:引数で指定したマクロを起動し、  
/filepath:引数で渡したビットマップをシートに貼って、  
/outfile:引数で渡したファイル名で保存します。  

vbspastesave.vbsを実行すると、実行時に渡した引数を表示する確認メッセージを出します。  

Windows 7 Ultimate SP1, Excel 2007の環境で動作を確認しています。  

## コマンドプロンプトから実行するコマンドラインの例  
vbspastesave.vbs /template:"C:\work\test\vbs\Book2.xls" /macro:"func_paste_bmp_save" /filepath:"C:\work\test\vbs\test.bmp" /outfile:"C:\work\test\vbs\test_save.xls"  

## VC++からVBSを実行したい場合は、下記のどちらかで。  
### 1.コマンドラインをShellExecuteに渡す  
### 2.CreateProcessにしたい場合はwscript/cscriptをコマンドラインの頭につけて呼び出す。  
  test.cppに呼び出し方の例を載せました。  

## batファイルからvbsを実行したい場合の例は、calltest.batです。  
その場合、vbspastesave.vbs内の49行目の WScript.Echo eRet を活かして下さい。  

## 参考サイト：  
* VBScript実行時エラー  
  http://msdn.microsoft.com/ja-jp/library/cc392383.aspx  
* VBScriptにおけるエラー処理  
  http://www.atmarkit.co.jp/fwin2k/tutor/cformwsh09/cformwsh09_02.html  
* バッチで使うVBScript  
  http://ys21.org/html/vbs4bat.html  
* コマンドラインにエラーを返すには  
  http://wsh.style-mods.net/topic7.htm  
* エクセルのシートへ画像ファイルを挿入し、セルのサイズ（セル範囲）に合わせて拡大・縮小して貼り付けてくれるＶＢＡプログラム。  
  http://plaza.rakuten.co.jp/plaplanet2007/diary/200705100000/  
* Excel でブックを閉じるときに表示される "変更を保存しますか?" というメッセージを非表示にする方法  
  http://support.microsoft.com/kb/213428/ja  


# HTML内でVBSファイルに引数を渡して起動（後日追加）  
Webシステム上で、帳票などをExcelで出力したいという場合に、サーバ側がLinuxだとそっちでは作れないので、  
クライアント側がWindowsでIEとExcelが入っているならそっちで処理してしまえ  
という発想のサンプルです。  

## 試し方
Webサーバを立てて  
クライアント側でIEからvbsstarttest.html  
にアクセスすると、calc(vbscript)ボタンのみが表示されます。  
IE_security.pngにある通り、IEでActiveXコントロールのスクリプト実行のところのセキュリティを弱めないと動きません。  
（常にそれだとまずいと思うけど・・・）  
（動作はIE 9で確認）
ボタンを押下すると、クライアント側のHKCUの"TEST_PATH2"環境変数のパスを見に行って、  
その中のvbspastebmp.vbsを実行します。  
テストのために引数を色々渡しています。  
Book1.xls：このサンプルだとHKCUの"TEST_PATH"（TEST_PATH2とは別にしてある）に存在する前提の帳票テンプレートファイル  
test.bmp：テンプレートに貼り付ける画像。このサンプルだとHKCUの"TEST_PATH"に存在する前提。  
test_save3.xls：このサンプルだと、この名前でHKCUの"TEST_PATH"に保存する。  
paste_bmp_test：画像貼り付けのマクロ名  
save_test：名前をつけて保存のマクロ名  

Book1.xlsを手元で作って、Book1_VBA_code.txtの内容を貼り付けてお使い下さい。  

## vbsファイルを暗号化したい場合
* WSHスクリプト・コードを暗号化する  
  http://www.atmarkit.co.jp/fwin2k/win2ktips/443wshenc/wshenc.html  

vbeファイルも、vbsファイルと同様に実行、動作します。


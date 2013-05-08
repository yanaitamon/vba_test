これはVBScriptからExcelの複数の引数を取るマクロを呼び出すサンプルです。  
vbsファイルの呼び出し元に戻り値を返します。  
VC++からのvbsファイル実行のサンプルも載せてあります。  

*目的*  
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

*コマンドプロンプトから実行するコマンドラインの例*  
vbspastesave.vbs /template:"C:\work\test\vbs\Book2.xls" /macro:"func_paste_bmp_save" /filepath:"C:\work\test\vbs\test.bmp" /outfile:"C:\work\test\vbs\test_save.xls"  

*VC++からVBSを実行したい場合は、下記のどちらかで。*  
1.コマンドラインをShellExecuteに渡す  
2.CreateProcessにしたい場合はwscript/cscriptをコマンドラインの頭につけて呼び出す。  
  test.cppに呼び出し方の例を載せました。  

*batファイルからvbsを実行したい場合の例は、calltest.batです。*  
その場合、vbspastesave.vbs内の49行目の WScript.Echo eRet を活かして下さい。  

*参考サイト：*  
VBScript実行時エラー  
http://msdn.microsoft.com/ja-jp/library/cc392383.aspx  
VBScriptにおけるエラー処理  
http://www.atmarkit.co.jp/fwin2k/tutor/cformwsh09/cformwsh09_02.html  
バッチで使うVBScript  
http://ys21.org/html/vbs4bat.html  
コマンドラインにエラーを返すには  
http://wsh.style-mods.net/topic7.htm  

これはVBScriptからExcelの複数の引数を取るマクロを呼び出すサンプルです。

/template:引数で指定したExcelファイルを開き、
/macro:引数で指定したマクロを起動し、
/filepath:引数で渡したビットマップをシートに貼って、
/outfile:引数で渡したファイル名で保存します。

vbspastesave.vbsを実行すると、実行時に渡した引数を表示する確認メッセージを出します。

Windows 7 Ultimate SP1, Excel 2007の環境で動作を確認しています。

コマンドラインの例
vbspastesave.vbs /template:"C:\work\test\vbs\Book2.xls" /macro:"func_paste_bmp_save" /filepath:"C:\work\test\vbs\test.bmp" /outfile:"C:\work\test\vbs\test_save.xls"


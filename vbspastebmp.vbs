Option Explicit

' vbspastebmp.vbs /template:"C:\work\test\vbs\Book1.xls" /macro_p:"paste_bmp_test" /macro_s:"save_test" /filepath:"C:\work\test\vbs\test.bmp" /outfile:"C:\work\test\vbs\test_save3.xls"
' CreateObject("Wscript.Shell") のオブジェクトのRunメソッドに引数を渡してvbsファイルを起動する時は、
' WshShell.Run "aaa.vbs /template:Book1.xls /macro_p:paste_bmp_test /macro_s:save_test /filepath:test.bmp /outfile:test_save3.xls"
' の形で良い。
' このサンプルは、
' クライアントのHKCU環境変数%TEST_PATH%を取得して、その中に保存されている、
' テンプレートエクセル Book1.xlsに
' 画像ファイル test.bmpを貼って
' %TEST_PATH%内にtest_save3.xlsという名前で保存する
' IEでしか動作しない。
' このまま使うと確認用のメッセージが出まくるので注意の事。

Dim oApp
' コマンドライン引数のチェック
Dim strArgument
If WScript.Arguments.Count = 0 Then
    WScript.Echo "it is called with no argument."
Else
    For Each strArgument In WScript.Arguments
        WScript.Echo strArgument
    Next
End If

Dim WshShell
Set WshShell = CreateObject("Wscript.Shell")

Dim WshSysEnv
Set WshSysEnv = WshShell.Environment("USER")

Dim strTestPath
strTestPath = WshSysEnv("TEST_PATH")
msgbox strTestPath

Set oApp = CreateObject("Excel.Application")

' 帳票Excelは非表示にする
oApp.Visible = False

'引数のチェック、ファイルを開く
Dim WshArguments
Dim WshNamed
Set WshArguments = WScript.Arguments
Set WshNamed = WshArguments.Named

' 本当は書き方次第なのだろうけど、
' Dim strMacroP As String と書くとコンパイルエラーと言われる
Dim strMacroP
strMacroP = "'" & strMacroP & strTestPath & "\" & WshNamed("template") & "'!" & WshNamed("macro_p")
Dim strMacroS
strMacroS = "'" & strMacroS & strTestPath & "\" & WshNamed("template") & "'!" & WshNamed("macro_s")

Dim strTemplate
strTemplate = strTestPath & "\" & WshNamed("template")

Dim strBmp
strBmp = strTestPath & "\" & WshNamed("filepath")

Dim strXls
strXls = strTestPath & "\" & WshNamed("outfile")

If WshNamed.Exists("template") Then
	oApp.Workbooks.Open strTemplate 'ファイルを開く
	' oApp.ActiveWorkbook.Worksheets("Sheet1").select
	' CStrか何かで明示的に文字列にしないと型が合わないと言われる
	oApp.Run strMacroP, CStr(strBmp)
	oApp.Run strMacroS, CStr(strXls)
	
	oApp.Workbooks.Close
End If

' ちゃんと解放
Set oApp = Nothing
Set WshShell = Nothing

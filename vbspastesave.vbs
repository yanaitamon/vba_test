Option Explicit

' コマンドラインの例
' vbspastesave.vbs /template:"C:\work\test\vbs\Book2.xls" /macro:"func_paste_bmp_save" /filepath:"C:\work\test\vbs\test.bmp" /outfile:"C:\work\test\vbs\test_save.xls"
' CreateProcessする場合のコマンドラインの例
' wscript vbspastesave.vbs /template:"C:\work\test\vbs\Book2.xls" /macro:"func_paste_bmp_save" /filepath:"C:\work\test\vbs\test.bmp" /outfile:"C:\work\test\vbs\test_save.xls"

Dim eRet
eRet = -1

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

Set oApp = CreateObject("Excel.Application")

' Excel操作中は非表示にする
oApp.Visible = False

' 引数のチェック、ファイルを開く、マクロ呼び出し
Dim WshArguments
Dim WshNamed
Set WshArguments = WScript.Arguments
Set WshNamed = WshArguments.Named

Dim strMacro
strMacro = ""
strMacro = "'" & strMacro & WshNamed("template") & "'!" & WshNamed("macro")

If WshNamed.Exists("template") Then
	oApp.Workbooks.Open WshNamed("template") 'ファイルを開く
	eRet = oApp.Run( strMacro, CStr(WshNamed("filepath")), CStr(WshNamed("outfile")) ) ' マクロ呼び出し
	oApp.Workbooks.Close
End If

If Err.Number <> 0 Then
	' WScript.Echo Err.Number
	eRet = Err.Number
End If

Set oApp = Nothing
' WScript.Echo eRet ' batファイルで呼び出した場合はQuitのリターンコードを見ない。Echoしたものが返る。
WScript.Quit (eRet)

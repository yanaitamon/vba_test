Option Explicit

' �R�}���h���C���̗�
' vbspastesave.vbs /template:"C:\work\test\vbs\Book2.xls" /macro:"func_paste_bmp_save" /filepath:"C:\work\test\vbs\test.bmp" /outfile:"C:\work\test\vbs\test_save.xls"

Dim oApp
' �R�}���h���C�������̃`�F�b�N
Dim strArgument
If WScript.Arguments.Count = 0 Then
    WScript.Echo "it is called with no argument."
Else
    For Each strArgument In WScript.Arguments
        WScript.Echo strArgument
    Next
End If

Set oApp = CreateObject("Excel.Application")

' Excel���쒆�͔�\���ɂ���
oApp.Visible = False

' �����̃`�F�b�N�A�t�@�C�����J���A�}�N���Ăяo��
Dim WshArguments
Dim WshNamed
Set WshArguments = WScript.Arguments
Set WshNamed = WshArguments.Named

Dim strMacro
strMacro = ""
strMacro = "'" & strMacro & WshNamed("template") & "'!" & WshNamed("macro")

If WshNamed.Exists("template") Then
	oApp.Workbooks.Open WshNamed("template") '�t�@�C�����J��
	WScript.Echo oApp.Run( strMacro, CStr(WshNamed("filepath")), CStr(WshNamed("outfile")) ) ' �}�N���Ăяo��

	oApp.Workbooks.Close
End If

Set oApp = Nothing

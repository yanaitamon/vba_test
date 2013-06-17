Option Explicit

' vbspastebmp.vbs /template:"C:\work\test\vbs\Book1.xls" /macro_p:"paste_bmp_test" /macro_s:"save_test" /filepath:"C:\work\test\vbs\test.bmp" /outfile:"C:\work\test\vbs\test_save3.xls"
' CreateObject("Wscript.Shell") �̃I�u�W�F�N�g��Run���\�b�h�Ɉ�����n����vbs�t�@�C�����N�����鎞�́A
' WshShell.Run "aaa.vbs /template:Book1.xls /macro_p:paste_bmp_test /macro_s:save_test /filepath:test.bmp /outfile:test_save3.xls"
' �̌`�ŗǂ��B
' ���̃T���v���́A
' �N���C�A���g��HKCU���ϐ�%TEST_PATH%���擾���āA���̒��ɕۑ�����Ă���A
' �e���v���[�g�G�N�Z�� Book1.xls��
' �摜�t�@�C�� test.bmp��\����
' %TEST_PATH%����test_save3.xls�Ƃ������O�ŕۑ�����
' IE�ł������삵�Ȃ��B
' ���̂܂܎g���Ɗm�F�p�̃��b�Z�[�W���o�܂���̂Œ��ӂ̎��B

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

Dim WshShell
Set WshShell = CreateObject("Wscript.Shell")

Dim WshSysEnv
Set WshSysEnv = WshShell.Environment("USER")

Dim strTestPath
strTestPath = WshSysEnv("TEST_PATH")
msgbox strTestPath

Set oApp = CreateObject("Excel.Application")

' ���[Excel�͔�\���ɂ���
oApp.Visible = False

'�����̃`�F�b�N�A�t�@�C�����J��
Dim WshArguments
Dim WshNamed
Set WshArguments = WScript.Arguments
Set WshNamed = WshArguments.Named

' �{���͏���������Ȃ̂��낤���ǁA
' Dim strMacroP As String �Ə����ƃR���p�C���G���[�ƌ�����
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
	oApp.Workbooks.Open strTemplate '�t�@�C�����J��
	' oApp.ActiveWorkbook.Worksheets("Sheet1").select
	' CStr�������Ŗ����I�ɕ�����ɂ��Ȃ��ƌ^������Ȃ��ƌ�����
	oApp.Run strMacroP, CStr(strBmp)
	oApp.Run strMacroS, CStr(strXls)
	
	oApp.Workbooks.Close
End If

' �����Ɖ��
Set oApp = Nothing
Set WshShell = Nothing

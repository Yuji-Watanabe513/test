' �y�g�����z'
' C:\>cscript runMacro.vbs "Excel�t�@�C����_�t���p�X" "���s����}�N����"'

Dim excelApp,file,macro,fpath
Dim excelWB
Dim sheet,uri,jyu,ara
dim fso

set fso = createObject("Scripting.FileSystemObject")
fpath = fso.getParentFolderName(WScript.ScriptFullName)
file = fpath + "\" + WScript.Arguments(0)
macro = WScript.Arguments(1)

Set excelApp = CreateObject("Excel.Application")

excelApp.Visible = False            ' Excel���\���ɂ���'
excelApp.ScreenUpdating = False
excelApp.DisplayAlerts = False      ' �|�b�v�A�b�v���b�Z�[�W���\���ɂ���'
excelApp.AutomationSecurity = 1     ' �}�N����L���ɂ���'
' Excel�t�@�C����r/w�A�O���Q�ƍX�V�ŊJ��'
set excelWB = excelApp.Workbooks.Open(file,3)
set sheet = ExcelWB.WorkSheets.Item(1)
if strComp(macro, "�f�[�^�X�V") = 0 then
	jyu = sheet.Cells(17,5)
	uri = sheet.Cells(17,9)
	ara = sheet.Cells(17,13)
end if
WScript.Echo "   �t�@�C���F" & file, "�}�N���F" & macro
' �}�N�������s����'
excelApp.Run macro
'WScript.Echo "---�}�N���̎��s���������܂���---"

if strComp(macro, "�f�[�^�X�V") = 0 then
	WScript.Echo "��="&cstr(round(jyu,0)),"����="&cstr(round(uri,0)),"�e��="&cstr(round(ara, 0))
	jyu = sheet.Cells(17,5)
	uri = sheet.Cells(17,9)
	ara = sheet.Cells(17,13)
	WScript.Echo "��="&cstr(round(jyu,0)),"����="&cstr(round(uri,0)),"�e��="&cstr(round(ara, 0))
end if
excelApp.Wait(Now + TimeValue("0:00:02"))

'WScript.Echo "---�ۑ����ďI�����܂�---"
excelApp.ScreenUpdating = True
excelWB.Save
excelWB.Close
excelApp.Quit
set excelWB = Nothing
Set excelApp = Nothing

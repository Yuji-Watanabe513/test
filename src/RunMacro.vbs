' 【使い方】'
' C:\>cscript runMacro.vbs "Excelファイル名_フルパス" "実行するマクロ名"'

Dim excelApp,file,macro,fpath
Dim excelWB
Dim sheet,uri,jyu,ara
dim fso

set fso = createObject("Scripting.FileSystemObject")
fpath = fso.getParentFolderName(WScript.ScriptFullName)
file = fpath + "\" + WScript.Arguments(0)
macro = WScript.Arguments(1)

Set excelApp = CreateObject("Excel.Application")

excelApp.Visible = False            ' Excelを非表示にする'
excelApp.ScreenUpdating = False
excelApp.DisplayAlerts = False      ' ポップアップメッセージを非表示にする'
excelApp.AutomationSecurity = 1     ' マクロを有効にする'
' Excelファイルをr/w、外部参照更新で開く'
set excelWB = excelApp.Workbooks.Open(file,3)
set sheet = ExcelWB.WorkSheets.Item(1)
if strComp(macro, "データ更新") = 0 then
	jyu = sheet.Cells(17,5)
	uri = sheet.Cells(17,9)
	ara = sheet.Cells(17,13)
end if
WScript.Echo "   ファイル：" & file, "マクロ：" & macro
' マクロを実行する'
excelApp.Run macro
'WScript.Echo "---マクロの実行が完了しました---"

if strComp(macro, "データ更新") = 0 then
	WScript.Echo "受注="&cstr(round(jyu,0)),"売上="&cstr(round(uri,0)),"粗利="&cstr(round(ara, 0))
	jyu = sheet.Cells(17,5)
	uri = sheet.Cells(17,9)
	ara = sheet.Cells(17,13)
	WScript.Echo "受注="&cstr(round(jyu,0)),"売上="&cstr(round(uri,0)),"粗利="&cstr(round(ara, 0))
end if
excelApp.Wait(Now + TimeValue("0:00:02"))

'WScript.Echo "---保存して終了します---"
excelApp.ScreenUpdating = True
excelWB.Save
excelWB.Close
excelApp.Quit
set excelWB = Nothing
Set excelApp = Nothing

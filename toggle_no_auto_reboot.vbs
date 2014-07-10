Dim wShell, currentSetting
const key="HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers"
Set wShell = CreateObject("WScript.Shell")

Function isAutoReboot()
    ' 未構成の場合、RegReadが失敗するため
    On Error Resume Next
    isAutoReboot=wShell.RegRead(key)
End Function

currentSetting = isAutoReboot()

If currentSetting = "" Or currentSetting = 0 Then
    If msgbox("自動再起動を無効にしますか？" ,vbYesNo,"Windows Update後の自動再起動") = vbYes Then
        wShell.RegWrite key, 1, "REG_DWORD"
	msgbox ("自動再起動を無効にしました。")
    End If
ElseIf currentSetting = 1 Then
    If msgbox("自動再起動を有効にしますか？" ,vbYesNo,"Windows Update後の自動再起動") = vbYes Then
        wShell.RegWrite key, 0, "REG_DWORD"
	msgbox ("自動再起動を有効にしました。")
    End If
End If

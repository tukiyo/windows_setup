Dim wShell, currentSetting
const key="HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers"
Set wShell = CreateObject("WScript.Shell")

Function isAutoReboot()
    ' ���\���̏ꍇ�ARegRead�����s���邽��
    On Error Resume Next
    isAutoReboot=wShell.RegRead(key)
End Function

currentSetting = isAutoReboot()

If currentSetting = "" Or currentSetting = 0 Then
    If msgbox("�����ċN���𖳌��ɂ��܂����H" ,vbYesNo,"Windows Update��̎����ċN��") = vbYes Then
        wShell.RegWrite key, 1, "REG_DWORD"
	msgbox ("�����ċN���𖳌��ɂ��܂����B")
    End If
ElseIf currentSetting = 1 Then
    If msgbox("�����ċN����L���ɂ��܂����H" ,vbYesNo,"Windows Update��̎����ċN��") = vbYes Then
        wShell.RegWrite key, 0, "REG_DWORD"
	msgbox ("�����ċN����L���ɂ��܂����B")
    End If
End If

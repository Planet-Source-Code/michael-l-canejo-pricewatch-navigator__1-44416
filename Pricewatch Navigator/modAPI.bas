Attribute VB_Name = "modAPI"
'Original Versions 1.0.0 - 1.2.7 [04/1/03]
'Brought to you and written by Mike Canejo
'-----------------------------------------
'AOL/AIM: Mikey3dd
'Email: MikeCanejo@hotmail.com
'-----------------------------------------
'Comments:
'-----------------------------------------
'None
'-----------------------------------------
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Sub FormDrag(frmObj As Form)
    frmMain.ActiveWindow
    DoEvents 'Let title bar refresh graphic change
    ReleaseCapture
    SendMessage frmObj.hwnd, &HA1, 2, 0& 'drag
End Sub

Public Sub GotoURL(strURL As String)
    ShellExecute FindWindow("ieframe", vbNullString), _
    vbNullString, strURL, vbNullString, "C:\", 1
End Sub

Public Function GetOpt(strKey As String, strDefault As String) As String
    GetOpt = GetSetting("Pricewatch Navigator", "Settings", strKey, strDefault)
End Function

Public Function SaveOpt(strKey As String, strValue As String) As String
    SaveSetting "Pricewatch Navigator", "Settings", strKey, strValue
End Function

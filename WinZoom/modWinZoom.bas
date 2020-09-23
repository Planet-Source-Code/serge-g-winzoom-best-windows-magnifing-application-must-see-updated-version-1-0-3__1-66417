Attribute VB_Name = "Module1"
'''''''''''''''''''''''Declare API for Always On Top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
''''''''''''''''''''''''Active Window API
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
'''''''''''''''''''''''Decare Constants for Always On Top


Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

'''''''''''''''''''''''Layered Window API

Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Const GWL_STYLE = (-16)

Const LWA_COLORKEY = &H1
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const BM_SETSTATE = &HF3

Const LWA_ALPHA = &H2
Const LWA_BOTH = &H3


'''''''''''''''''''''''Stansard declarations

Public Const APP_NAME = "WinZoom_1"
Public Const SECTION_NAME = "StartUp_Position"

Public isOnTop As Boolean
Public intHeight As Integer
Public intWidth As Integer
Public intTop As Integer
Public intLeft As Integer
Public isTop As Boolean
Public isNoFreeze As Boolean
Public isAim As Boolean
Public zoomAt As Integer
Public isBorder As Boolean
Public isPixColor As Boolean
Public pixBackColor As String
Public copyHEX As String
Public copyVBHEX As String
Public copyRGB As String
Public pixForeColor As String
Public pixLeft As Integer
Public pixTop As Integer
Public isPixActive As Boolean
Public isAttach As Boolean
Public bCancel As Boolean
Public useHKeys As Boolean
Public infoLabel As Boolean
Private isReg As Boolean
Public activeWinImg As Long
Public activeWinPixel As Long
Public activeWinAbout As Long
Public activeWinHelp As Long
Public activeWinTip As Long
Public activeWinComposer As Long
Public invisibleMode As Boolean
Public isComposerActive As Boolean
Public R%, G%, b%


Public Sub setTopMost()

    If isOnTop Then
        SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
    
End Sub

Public Sub setTopMost2(frm As Form)

    If isOnTop Then
        SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
    
End Sub


Public Sub getRegSet()

    intHeight = GetSetting(APP_NAME, SECTION_NAME, "FormHeight", 3200)
    intWidth = GetSetting(APP_NAME, SECTION_NAME, "FormWidth", 8100)
    intTop = GetSetting(APP_NAME, SECTION_NAME, "FormTop", (Screen.Height / 2) - (frmMain.Height / 2))
    intLeft = GetSetting(APP_NAME, SECTION_NAME, "FormLeft", (Screen.Width / 2) - (frmMain.Width / 2))
    isTop = GetSetting(APP_NAME, SECTION_NAME, "FormOnTop", "True")
    isNoFreeze = GetSetting(APP_NAME, SECTION_NAME, "DisableFreeze", "True")
    zoomAt = GetSetting(APP_NAME, SECTION_NAME, "ZoomLevel", 5)
    isAim = GetSetting(APP_NAME, SECTION_NAME, "ShowAim", "False")
    isBorder = GetSetting(APP_NAME, SECTION_NAME, "ShowBorder", "True")
    pixBackColor = GetSetting(APP_NAME, SECTION_NAME, "PixelFormBackColor", "&H00404040")
    pixForeColor = GetSetting(APP_NAME, SECTION_NAME, "PixelFormForeColor", "&H8000000E")
    pixTop = GetSetting(APP_NAME, SECTION_NAME, "PixFormTop", 0)
    pixLeft = GetSetting(APP_NAME, SECTION_NAME, "PixFormLeft", 0)
    isAttach = GetSetting(APP_NAME, SECTION_NAME, "PixFormAttach", "False")
    infoLabel = GetSetting(APP_NAME, SECTION_NAME, "InfoLabel", "True")
    useHKeys = GetSetting(APP_NAME, SECTION_NAME, "UseHotKeys", "True")
    isReg = GetSetting(APP_NAME, SECTION_NAME, "IsRegistryData", "False")
    invisibleMode = GetSetting(APP_NAME, SECTION_NAME, "InvisibleMode", "True")
    frmMain.mnuCursoronTopItem.Checked = GetSetting(APP_NAME, SECTION_NAME, "XYOnScreen", "False")
    
    If frmMain.mnuCursoronTopItem.Checked = True Then
        frmMain.lblCurPos_Click
    End If
    
    If isReg Then
        frmMain.mnuClearSetting.Enabled = True
    Else
        frmMain.mnuClearSetting.Enabled = False
    End If
    
End Sub

Public Sub setRegSet()

    isReg = True

    SaveSetting APP_NAME, SECTION_NAME, "FormHeight", intHeight
    SaveSetting APP_NAME, SECTION_NAME, "FormWidth", intWidth
    SaveSetting APP_NAME, SECTION_NAME, "FormTop", intTop
    SaveSetting APP_NAME, SECTION_NAME, "FormLeft", intLeft
    SaveSetting APP_NAME, SECTION_NAME, "FormOnTop", isTop
    SaveSetting APP_NAME, SECTION_NAME, "DisableFreeze", isNoFreeze
    SaveSetting APP_NAME, SECTION_NAME, "ZoomLevel", zoomAt
    SaveSetting APP_NAME, SECTION_NAME, "ShowAim", isAim
    SaveSetting APP_NAME, SECTION_NAME, "ShowBorder", isBorder
    SaveSetting APP_NAME, SECTION_NAME, "PixelFormBackColor", pixBackColor
    SaveSetting APP_NAME, SECTION_NAME, "PixelFormForeColor", pixForeColor
    SaveSetting APP_NAME, SECTION_NAME, "PixFormTop", pixTop
    SaveSetting APP_NAME, SECTION_NAME, "PixFormLeft", pixLeft
    SaveSetting APP_NAME, SECTION_NAME, "PixFormAttach", isAttach
    SaveSetting APP_NAME, SECTION_NAME, "InfoLabel", infoLabel
    SaveSetting APP_NAME, SECTION_NAME, "UseHotKeys", useHKeys
    SaveSetting APP_NAME, SECTION_NAME, "IsRegistryData", isReg
    SaveSetting APP_NAME, SECTION_NAME, "InvisibleMode", invisibleMode
    SaveSetting APP_NAME, SECTION_NAME, "XYOnScreen", frmMain.mnuCursoronTopItem.Checked
    
End Sub

Public Sub delRegSet()

    On Error GoTo erHand
    DeleteSetting APP_NAME, SECTION_NAME
    'MsgBox ("All user settings are deleted."), vbExclamation, "Registry Cleared"
    isReg = False
    frmMain.mnuClearSetting.Enabled = False
    Exit Sub
    
erHand:
    Exit Sub

End Sub


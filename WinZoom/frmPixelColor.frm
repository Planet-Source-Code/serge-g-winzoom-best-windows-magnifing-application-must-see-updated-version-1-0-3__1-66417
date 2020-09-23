VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPixelColor 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   375
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   1125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   1125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmnDColor 
      Left            =   -180
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menu"
      Height          =   255
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   1740
      Left            =   690
      MousePointer    =   15  'Size All
      TabIndex        =   0
      ToolTipText     =   "Right - Click for Menu"
      Top             =   45
      Width           =   6480
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   765
      TabIndex        =   6
      Top             =   75
      Width           =   255
   End
   Begin VB.Label label2 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   735
      TabIndex        =   5
      Top             =   45
      Width           =   300
   End
   Begin VB.Label Frame3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VBHEX : "
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   1575
   End
   Begin VB.Label Frame2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HEX : "
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Frame1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RGB : "
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Menu hdnMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuBColorItem 
         Caption         =   "Change Background Color"
      End
      Begin VB.Menu mnuChangeForeItem 
         Caption         =   "Change Text Color"
      End
      Begin VB.Menu mnuResetBGItem 
         Caption         =   "Reset Colors"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowRGBItem 
         Caption         =   "Show RGB"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowHEXItem 
         Caption         =   "Show HEX"
      End
      Begin VB.Menu mnuShowVBItem 
         Caption         =   "Show VBHEX"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyRGBItem 
         Caption         =   "Copy RGB to Clipboard"
      End
      Begin VB.Menu mnuCopyHEXItem 
         Caption         =   "Copy HEX to Clipboard"
      End
      Begin VB.Menu mnuCopyVBItem 
         Caption         =   "Copy VB - HEX to Clipboard"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAttachItem 
         Caption         =   "Attach To WinZoom"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "Hide Pixel Color"
      End
   End
End
Attribute VB_Name = "frmPixelColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isMove As Boolean
Dim intVisible As Integer
Dim theX As Integer
Dim theY As Integer

Private Sub Command1_Click()

    PopupMenu hdnMenu

End Sub

Private Sub Form_Load()

    isPixActive = True
    isMove = False
    Me.BackColor = pixBackColor
    Frame1.ForeColor = pixForeColor
    Frame2.ForeColor = pixForeColor
    Frame3.ForeColor = pixForeColor
    Me.Top = pixTop
    Me.Left = pixLeft
    setSize
    doAttach
    
    Show
    activeWinPixel = GetActiveWindow
    
    If isComposerActive Then
        frmComposer.cmdFromPixel.Enabled = True
    End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

    pixTop = Me.Top
    pixLeft = Me.Left
    setRegSet
    frmMain.mnuShowPixelColorItem.Checked = False
    frmMain.Toolbar2.Buttons(10).Enabled = True
    frmMain.Toolbar2.Buttons(11).Enabled = False
    isPixColor = False
    pixBackColor = Me.BackColor
    isPixActive = False
    
    If isComposerActive = True Then
        frmComposer.cmdFromPixel.Enabled = True
    End If
    
    activeWinPixel = -1

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        isMove = True
        theX = X
        theY = Y
    ElseIf Button = 2 Then
        isMove = False
        PopupMenu hdnMenu
    End If

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If isMove Then
        Me.Top = Me.Top + Y - theY
        Me.Left = Me.Left + X - theX
    Else
    End If

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If isMove = True Then
        If isAttach Then
            Me.Top = frmMain.Top + frmMain.Height - Me.Height
            Me.Left = frmMain.Left
        End If
    Else
    End If
    isMove = False

End Sub

Sub setSize()

    Dim intSpace As Integer
    Dim cnt As Integer
    Dim setFrames(1 To 3) As String

    intVisible = 0
    intSpace = 1200
    cnt = 0
    
    If mnuShowRGBItem.Checked = True Then
        intVisible = 1
        intSpace = intSpace + 1500
        cnt = cnt + 1
        setFrames(cnt) = "RGB"
    End If
    
    If mnuShowHEXItem.Checked = True Then
        intVisible = intVisible + 1
        intSpace = intSpace + 1500
        cnt = cnt + 1
        setFrames(cnt) = "HEX"
    End If
    
    If mnuShowVBItem.Checked = True Then
        intVisible = intVisible + 1
        intSpace = intSpace + 1500
        cnt = cnt + 1
        setFrames(cnt) = "HEXVB"
    End If
    
    prepFrames
    
    For i = LBound(setFrames) To UBound(setFrames)
        If setFrames(i) = "" Then Exit For
        'MsgBox (setFrames(i))
        Select Case setFrames(i)
        Case "RGB"
            Frame1.Visible = True
            Frame1.Top = 100
            Frame1.Left = selectCaseI(i)
        Case "HEX"
            Frame2.Left = selectCaseI(i)
            Frame2.Visible = True
            Frame2.Top = 100
        Case "HEXVB"
            Frame3.Left = selectCaseI(i)
            Frame3.Visible = True
            Frame3.Top = 100
        End Select
    Next i
    
    If intSpace = 0 Then
        Unload Me
    Else
        Me.Width = intSpace
    End If
    
End Sub

Private Function selectCaseI(theI)

    Dim lft As Integer
        
    Select Case theI
    Case 1: lft = 980
    Case 2: lft = 2400
    Case 3: lft = 3900
    End Select

    selectCaseI = lft

End Function

Private Sub prepFrames()

    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False

End Sub

Private Sub mnuAttachItem_Click()

    mnuAttachItem.Checked = Not mnuAttachItem.Checked
    isAttach = mnuAttachItem.Checked
    
    If isAttach Then
        Me.Top = frmMain.Top + frmMain.Height - Me.Height
        Me.Left = frmMain.Left
    Else
    End If

End Sub

Private Sub mnuBColorItem_Click()

    cmnDColor.ShowColor
    On Error GoTo erHand
    Me.BackColor = cmnDColor.Color
    pixBackColor = Me.BackColor
    
erHand:
    Exit Sub

End Sub

Private Sub mnuChangeForeItem_Click()

    cmnDColor.ShowColor
    On Error GoTo erHand
    pixForeColor = cmnDColor.Color
    Frame1.ForeColor = pixForeColor
    Frame2.ForeColor = pixForeColor
    Frame3.ForeColor = pixForeColor
    
erHand:
    Exit Sub

End Sub

Private Sub mnuCopyHEXItem_Click()

    Clipboard.Clear
    Clipboard.SetText copyHEX

End Sub

Private Sub mnuCopyRGBItem_Click()

    Clipboard.Clear
    Clipboard.SetText copyRGB

End Sub

Private Sub mnuCopyVBItem_Click()

    Clipboard.Clear
    Clipboard.SetText copyVBHEX

End Sub

Private Sub mnuExitItem_Click()

    Unload Me

End Sub

Private Sub mnuResetBGItem_Click()

    Me.BackColor = &H404040
    Frame1.ForeColor = vbWhite
    Frame2.ForeColor = vbWhite
    Frame3.ForeColor = vbWhite
    pixForeColor = vbWhite
    pixBackColor = &H404040

End Sub

Private Sub mnuShowHEXItem_Click()

    mnuShowHEXItem.Checked = Not mnuShowHEXItem.Checked
    setSize

End Sub

Private Sub mnuShowRGBItem_Click()

    mnuShowRGBItem.Checked = Not mnuShowRGBItem.Checked
    setSize

End Sub

Private Sub mnuShowVBItem_Click()

    mnuShowVBItem.Checked = Not mnuShowVBItem.Checked
    setSize

End Sub

Sub doAttach()

    If isAttach = True Then
        mnuAttachItem_Click
    End If

End Sub

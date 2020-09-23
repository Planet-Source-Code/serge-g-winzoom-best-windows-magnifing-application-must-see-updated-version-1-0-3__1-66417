VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLoadImage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Preview"
   ClientHeight    =   4950
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7455
   Icon            =   "frmLoadImage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cmnDlg2 
      Left            =   3330
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   4890
      Left            =   0
      Picture         =   "frmLoadImage.frx":0442
      Stretch         =   -1  'True
      Top             =   30
      Width           =   7455
   End
   Begin VB.Menu mnuFileItm 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenItem 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnloadItem 
         Caption         =   "&Go Back"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit Program"
      End
   End
End
Attribute VB_Name = "frmLoadImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isOnTopTemp As Boolean

Private Sub Form_Load()

    setTopMost2 Me
    frmMain.Toolbar3.Buttons(3).Enabled = True
    frmMain.mnuSampleItem.Checked = True
    Me.Show
    activeWinImg = GetActiveWindow
    
End Sub

Private Sub Form_Resize()

'
'

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.Toolbar3.Buttons(3).Enabled = False
    frmMain.mnuSampleItem.Checked = False
    activeWinImg = -1

End Sub

Private Sub mnuExitItem_Click()

    Unload frmMain
    Unload Me
    End

End Sub

Private Sub mnuOpenItem_Click()

    cmnDlg2.Filter = "Bitmap Images( *.bmp)|*.bmp|JPEG Images ( *.jpg)|*.jpg|All Images (*.*)|*.*"
    cmnDlg2.CancelError = True
    On Error GoTo erHand
    cmnDlg2.ShowOpen
    If cmnDlg2.FileName <> "" Then
        Image1.Picture = LoadPicture(cmnDlg2.FileName)
        Image1.Stretch = False
        
        Me.Height = Image1.Height + 1000
        Me.Width = Image1.Width + 250
        
        Image1.Top = 100
        Image1.Left = 60
        
    End If
erHand:
    If Err.Number = 32755 Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox ("Invalid File"), , "Error"
    End If

End Sub

Private Sub mnuUnloadItem_Click()
    
    Unload Me

End Sub

VERSION 5.00
Begin VB.Form frmTip 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4485
   ClientLeft      =   2325
   ClientTop       =   2070
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   3990
      TabIndex        =   4
      Top             =   0
      Width           =   3990
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3690
         TabIndex        =   5
         Top             =   15
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start-Up Tip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   15
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   120
      Picture         =   "frmTip.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   240
      Width           =   3735
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   720
         TabIndex        =   10
         Top             =   3165
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   2775
         TabIndex        =   9
         Top             =   3165
         Width           =   180
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTip.frx":030A
         Height          =   855
         Left            =   195
         TabIndex        =   8
         Top             =   510
         Width           =   3270
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   1710
         Left            =   180
         TabIndex        =   7
         Top             =   1425
         Width           =   3315
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   3
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "This will not appear again."
         Height          =   195
         Left            =   900
         TabIndex        =   2
         Top             =   3150
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000A&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1425
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const LB_ITEMFROMPOINT = &H1A9

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg _
        As Long, ByVal wParam As Long, lParam As Any) As Long
    
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type
Const GRADIENT_FILL_RECT_H As Long = &H0
Const GRADIENT_FILL_RECT_V  As Long = &H1
Const GRADIENT_FILL_TRIANGLE As Long = &H2
Const GRADIENT_FILL_OP_FLAG As Long = &HFF
Private Declare Function GdiGradientFillRect Lib "gdi32" Alias "GdiGradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Private Function LongToUShort(Unsigned As Long) As Integer

    LongToUShort = CInt(Unsigned - &H10000)
    
End Function



Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub Command1_Click()

    cmdOK_Click

End Sub

Private Sub Form_Load()

    Dim R As String
    
    R = vbCrLf
    
    Label3.Caption = "Control + (+) : Zoom + 1" & R & "Control 6 (-) : Zoom - 1" & R & "Control + A : Toggle Aim" & R & "Control + C : Toggle Composer" & R & "Control + F : Freeze the image" & R & "Control + T : Toggle Toolbar" & R & "Control + P : Toggle Pixel Color Window" & R & "Control + Escape : Quit"
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Dim SR&, SG&, SB&
    SR = 33
    SG = 60
    SB = 126
    
    pic.ScaleMode = 3
    For i = 0 To pic.ScaleHeight
        For j = 0 To pic.ScaleWidth
            pic.PSet (j, i), RGB(SR + i, SG + i, SB + j)
        Next j
    Next i
    
    Show
    activeWinTip = GetActiveWindow
    setTopMost2 Me
    'setGradient   - Decomment to make form grdiently fade from light-blue to white...

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim l As Long

    If Button = 1 Then
        ReleaseCapture
        l = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    activeWinTip = -1
    
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim l As Long

    If Button = 1 Then
        ReleaseCapture
        l = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
    
End Sub

Sub setGradient()
    
    ScaleMode = 3

    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    'from
    With vert(0)
        .x = 0
        .y = 0
        .Red = LongToUShort(&HFF00&) '0&
        .Green = LongToUShort(&HFF00&) '0& '&HFF&   '0&
        .Blue = LongToUShort(&HFF00&) '0&
        .Alpha = 0&
    End With
    'to
    With vert(1)
        .x = Me.ScaleWidth
        .y = Me.ScaleHeight
        .Red = 0&
        .Green = LongToUShort(&HAA00&) ' 0&
        .Blue = LongToUShort(&HFF00&)   '0&
        .Alpha = 0&
    End With
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    GdiGradientFillRect Me.hdc, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V
    
    ScaleMode = 1

End Sub


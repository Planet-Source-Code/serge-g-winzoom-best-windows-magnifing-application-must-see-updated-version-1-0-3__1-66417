VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4485
      Top             =   3300
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H80000009&
      Height          =   1620
      Left            =   270
      ScaleHeight     =   1560
      ScaleWidth      =   4440
      TabIndex        =   1
      Top             =   270
      Width           =   4500
      Begin VB.PictureBox picAbout 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   2820
         Left            =   105
         ScaleHeight     =   2820
         ScaleWidth      =   4380
         TabIndex        =   2
         Top             =   -1335
         Width           =   4380
         Begin VB.PictureBox picZ 
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   225
            ScaleHeight     =   435
            ScaleWidth      =   3690
            TabIndex        =   4
            Top             =   2280
            Width           =   3690
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "huleek. Thank you for your interest in my app."
               Height          =   195
               Left            =   225
               TabIndex        =   5
               Top             =   75
               Width           =   3270
            End
            Begin VB.Image Image1 
               Height          =   240
               Left            =   -15
               Picture         =   "frmAbout.frx":0CCA
               Top             =   15
               Width           =   240
            End
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":1B0C
            Height          =   1935
            Left            =   240
            TabIndex        =   3
            Top             =   90
            Width           =   4260
         End
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H002EB9B9&
      Default         =   -1  'True
      DownPicture     =   "frmAbout.frx":1CFE
      Height          =   420
      Left            =   3480
      MaskColor       =   &H8000000A&
      Picture         =   "frmAbout.frx":2E80
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2445
      UseMaskColor    =   -1  'True
      Width           =   1395
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00404040&
      X1              =   105
      X2              =   105
      Y1              =   2160
      Y2              =   3030
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      X1              =   975
      X2              =   975
      Y1              =   2160
      Y2              =   3045
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      X1              =   105
      X2              =   990
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00404040&
      X1              =   105
      X2              =   990
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By"
      Height          =   195
      Left            =   1305
      TabIndex        =   7
      Top             =   2475
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "huleek"
      Height          =   195
      Left            =   1350
      TabIndex        =   6
      Top             =   2865
      Width           =   480
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   1080
      Picture         =   "frmAbout.frx":4002
      Stretch         =   -1  'True
      Top             =   2805
      Width           =   240
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   885
      Left            =   105
      Picture         =   "frmAbout.frx":4E44
      Top             =   2160
      Width           =   885
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   4920
      X2              =   4920
      Y1              =   90
      Y2              =   2055
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4920
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4920
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   2040
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      X1              =   4800
      X2              =   4800
      Y1              =   240
      Y2              =   1920
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   4800
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   4800
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   1920
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Click()
    
    onClick

End Sub

Private Sub Form_Load()

    picAbout.Container = picFrame
    picAbout.Move 50, 0
    Label2.Container = picAbout
    Label2.Move 0, 0
    
    setTopMost2 Me
    
    Show
    activeWinAbout = GetActiveWindow
    frmMain.mnuAboutItem.Checked = True
    
    gradient

End Sub

Private Sub Form_Unload(Cancel As Integer)

    activeWinAbout = -1
    frmMain.mnuAboutItem.Checked = False

End Sub

Private Sub Image1_Click()

    onClick

End Sub

Private Sub Label1_Click()

    onClick

End Sub

Private Sub Label2_Click()

    onClick

End Sub

Private Sub picAbout_Click()

    onClick

End Sub

Private Sub picFrame_Click()

    onClick

End Sub

Private Sub Timer1_Timer()

    picAbout.Top = picAbout.Top - 10
    If picAbout.Top <= (picAbout.Height * -1) + 400 Then picAbout.Top = picFrame.Height

End Sub

Private Sub onClick()

    Timer1.Enabled = Not Timer1.Enabled

End Sub

Sub gradient()
    
    ScaleMode = 3

    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    'from black
    With vert(0)
        .x = 0
        .y = 0
        .Red = LongToUShort(&HFF00&) '0&
        .Green = LongToUShort(&HFF00&) '0& '&HFF&   '0&
        .Blue = LongToUShort(&HFF00&) '0&
        .Alpha = 0&
    End With
    'to red
    With vert(1)
        .x = Me.ScaleWidth
        .y = Me.ScaleHeight
        .Red = LongToUShort(&HAA00&) ' 0&
        .Green = LongToUShort(&HAA00&)  '0&
        .Blue = 0&
        .Alpha = 0&
    End With
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    GdiGradientFillRect Me.hdc, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V
    
    ScaleMode = 1

End Sub

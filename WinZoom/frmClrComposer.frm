VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmComposer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Composer"
   ClientHeight    =   5880
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5790
   FillColor       =   &H80000000&
   Icon            =   "frmClrComposer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRR 
      Height          =   375
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   4710
      Width           =   375
   End
   Begin VB.CommandButton cmdRG 
      Height          =   375
      Left            =   3615
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   4710
      Width           =   375
   End
   Begin VB.CommandButton cmdRB 
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   4710
      Width           =   375
   End
   Begin VB.CommandButton cmdRnd 
      Height          =   375
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   4710
      Width           =   405
   End
   Begin VB.CommandButton cmdOKHEX 
      Caption         =   "OK"
      Height          =   255
      Left            =   4605
      TabIndex        =   44
      Top             =   3045
      Width           =   375
   End
   Begin VB.CommandButton cmdToggleH 
      Caption         =   "Basic Colors"
      Height          =   255
      Left            =   330
      TabIndex        =   95
      Tag             =   "b"
      Top             =   2760
      Width           =   2015
   End
   Begin VB.PictureBox picHide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   45
      ScaleHeight     =   2175
      ScaleWidth      =   2625
      TabIndex        =   86
      Top             =   3000
      Width           =   2625
      Begin VB.PictureBox picMinPalette 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   1740
         Left            =   285
         MouseIcon       =   "frmClrComposer.frx":014A
         MousePointer    =   99  'Custom
         ScaleHeight     =   130
         ScaleMode       =   0  'User
         ScaleWidth      =   127.09
         TabIndex        =   87
         Top             =   340
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4830
      TabIndex        =   108
      Tag             =   "more"
      Top             =   5520
      Width           =   855
   End
   Begin VB.PictureBox picRnd 
      Height          =   375
      Left            =   4080
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   107
      Top             =   8640
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picRB 
      Height          =   375
      Left            =   3720
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   105
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picRG 
      Height          =   375
      Left            =   3360
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   104
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picRR 
      Height          =   375
      Left            =   3000
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   103
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picConvert 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   5970
      MouseIcon       =   "frmClrComposer.frx":029C
      MousePointer    =   99  'Custom
      ScaleHeight     =   315
      ScaleWidth      =   1995
      TabIndex        =   80
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Convert To"
      Height          =   1935
      Left            =   5970
      TabIndex        =   53
      Top             =   3360
      Width           =   2055
      Begin VB.CommandButton cmdConvert 
         Caption         =   "Convert"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdCopyConverted 
         Caption         =   "Copy Result"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optToLong 
         Caption         =   "   Long"
         Height          =   255
         Left            =   960
         TabIndex        =   58
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optToVB 
         Caption         =   "VBHEX"
         Height          =   255
         Left            =   960
         TabIndex        =   57
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optToHEX 
         Caption         =   "HEX"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optToRGB 
         Caption         =   "RGB"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtConverted 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Convert From"
      Height          =   1215
      Left            =   5970
      TabIndex        =   47
      Top             =   2040
      Width           =   2055
      Begin VB.OptionButton optFromLong 
         Caption         =   "   Long"
         Height          =   255
         Left            =   960
         TabIndex        =   52
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optFromVB 
         Caption         =   "VBHEX"
         Height          =   255
         Left            =   960
         TabIndex        =   51
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optFromHEX 
         Caption         =   "HEX"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton optFromRGB 
         Caption         =   "RGB"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtConvertTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   2070
      TabIndex        =   41
      Text            =   "Long"
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLong 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "0"
      Top             =   9570
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComCtl2.FlatScrollBar VSB 
      Height          =   1900
      Left            =   4205
      TabIndex        =   36
      Top             =   120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   3360
      _Version        =   393216
      Appearance      =   2
      LargeChange     =   15
      Max             =   255
      Orientation     =   1179648
   End
   Begin MSComCtl2.FlatScrollBar VSG 
      Height          =   1900
      Left            =   3725
      TabIndex        =   35
      Top             =   120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   3360
      _Version        =   393216
      Appearance      =   2
      LargeChange     =   15
      Max             =   255
      Orientation     =   1179648
   End
   Begin MSComCtl2.FlatScrollBar VSR 
      Height          =   1900
      Left            =   3225
      TabIndex        =   34
      Top             =   120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   3360
      _Version        =   393216
      Appearance      =   2
      LargeChange     =   15
      Max             =   255
      Orientation     =   1179648
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   2280
      ScaleHeight     =   1875
      ScaleWidth      =   435
      TabIndex        =   30
      Top             =   120
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cmnDlg1 
      Left            =   1440
      Top             =   9120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCstmReset 
      Caption         =   "Reset Custom Colors"
      Height          =   255
      Left            =   6120
      TabIndex        =   23
      Top             =   945
      Width           =   1815
   End
   Begin VB.CommandButton cmdPalette 
      Caption         =   "System Palette"
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   1665
      Width           =   1815
   End
   Begin VB.CommandButton cmdFromPixel 
      Caption         =   "Get From PixelColor"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6120
      TabIndex        =   21
      ToolTipText     =   "When PixelColor Toolbar is active, you can transfer the color to the Composer"
      Top             =   1305
      Width           =   1815
   End
   Begin VB.CommandButton cmdFlash 
      Caption         =   "cmdFlash"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   9480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add Custom Colors"
      Height          =   735
      Left            =   5955
      TabIndex        =   16
      Top             =   105
      Width           =   2055
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         FillColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         MouseIcon       =   "frmClrComposer.frx":03EE
         MousePointer    =   99  'Custom
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         FillColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         MouseIcon       =   "frmClrComposer.frx":0540
         MousePointer    =   99  'Custom
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         FillColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         MouseIcon       =   "frmClrComposer.frx":0692
         MousePointer    =   99  'Custom
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmClrComposer.frx":07E4
         MousePointer    =   99  'Custom
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   17
         Tag             =   "e"
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Exit"
      Height          =   255
      Left            =   285
      TabIndex        =   15
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox Text6 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   2085
      TabIndex        =   14
      Text            =   "VBHEX"
      Top             =   9210
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   4365
      TabIndex        =   13
      Text            =   "HEX"
      Top             =   9660
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   200
      Left            =   4380
      TabIndex        =   12
      Text            =   "RGB"
      Top             =   9285
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtVB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "&H"
      Top             =   9180
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtHEX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4305
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   9630
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0  0  0"
      Top             =   9240
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   840
      Top             =   9120
   End
   Begin VB.PictureBox picView 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000002&
      Height          =   300
      Left            =   4065
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   2595
      Width           =   375
   End
   Begin VB.VScrollBar oVSB 
      Height          =   2175
      LargeChange     =   15
      Left            =   600
      Max             =   255
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9105
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.VScrollBar oVSG 
      Height          =   2175
      LargeChange     =   15
      Left            =   360
      Max             =   255
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9105
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.VScrollBar oVSR 
      Height          =   2175
      LargeChange     =   15
      Left            =   120
      Max             =   255
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9105
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000002&
      Height          =   300
      Left            =   3120
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "f"
      Text            =   "0"
      Top             =   2595
      Width           =   375
   End
   Begin VB.TextBox txtG 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000002&
      Height          =   300
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   2595
      Width           =   375
   End
   Begin VB.PictureBox picOR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmClrComposer.frx":0936
      MousePointer    =   99  'Custom
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   31
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox picGradient 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmClrComposer.frx":0A88
      MousePointer    =   99  'Custom
      ScaleHeight     =   315
      ScaleWidth      =   2625
      TabIndex        =   29
      Top             =   2160
      Width           =   2685
   End
   Begin VB.PictureBox picOG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3600
      MouseIcon       =   "frmClrComposer.frx":0BDA
      MousePointer    =   99  'Custom
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   32
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox picOB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4080
      MouseIcon       =   "frmClrComposer.frx":0D2C
      MousePointer    =   99  'Custom
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   33
      Top             =   2160
      Width           =   375
   End
   Begin VB.VScrollBar VStxtR 
      Height          =   300
      Left            =   3495
      Max             =   255
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2595
      Width           =   135
   End
   Begin VB.VScrollBar VStxtG 
      Height          =   300
      Left            =   3960
      Max             =   255
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2595
      Width           =   135
   End
   Begin VB.VScrollBar VStxtB 
      Height          =   300
      Left            =   4425
      Max             =   255
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2595
      Width           =   135
   End
   Begin VB.Frame Frame4 
      Caption         =   "Color Randomizer"
      Height          =   870
      Left            =   2955
      TabIndex        =   99
      Top             =   4410
      Width           =   2145
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   2790
      TabIndex        =   97
      Text            =   "VB"
      Top             =   3420
      Width           =   240
   End
   Begin VB.TextBox txtEnterVB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2775
      TabIndex        =   39
      Top             =   3405
      Width           =   1800
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   2775
      TabIndex        =   98
      Text            =   "Long"
      Top             =   3810
      Width           =   390
   End
   Begin VB.TextBox txtEnterLong 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   43
      Top             =   3795
      Width           =   1800
   End
   Begin VB.CommandButton cmdCopyRGB 
      Caption         =   "Copy"
      Height          =   255
      Left            =   4590
      TabIndex        =   7
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdOKVB 
      Caption         =   "OK"
      Height          =   255
      Left            =   4605
      TabIndex        =   45
      Top             =   3435
      Width           =   375
   End
   Begin VB.CommandButton cmdCopyHEX 
      Caption         =   "Copy"
      Height          =   255
      Left            =   5010
      TabIndex        =   37
      Top             =   3045
      Width           =   615
   End
   Begin VB.CommandButton cmdCopyVB 
      Caption         =   "Copy"
      Height          =   255
      Left            =   5010
      TabIndex        =   38
      Top             =   3435
      Width           =   615
   End
   Begin VB.CommandButton cmdCopyLong 
      Caption         =   "Copy"
      Height          =   255
      Left            =   5010
      TabIndex        =   42
      Top             =   3825
      Width           =   615
   End
   Begin VB.CommandButton cmdOKLong 
      Caption         =   "OK"
      Height          =   255
      Left            =   4590
      TabIndex        =   46
      Top             =   3825
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   2805
      TabIndex        =   96
      Text            =   "HEX"
      Top             =   3030
      Width           =   345
   End
   Begin VB.TextBox txtEnterHEX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2790
      TabIndex        =   109
      Top             =   3015
      Width           =   1800
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   2655
      X2              =   75
      Y1              =   4755
      Y2              =   4755
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00404040&
      X1              =   75
      X2              =   75
      Y1              =   4755
      Y2              =   3240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   75
      X2              =   2655
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   2655
      X2              =   2655
      Y1              =   3240
      Y2              =   4755
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   2325
      MouseIcon       =   "frmClrComposer.frx":0E7E
      MousePointer    =   99  'Custom
      TabIndex        =   94
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   1965
      MouseIcon       =   "frmClrComposer.frx":0FD0
      MousePointer    =   99  'Custom
      TabIndex        =   93
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   1605
      MouseIcon       =   "frmClrComposer.frx":1122
      MousePointer    =   99  'Custom
      TabIndex        =   92
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   1245
      MouseIcon       =   "frmClrComposer.frx":1274
      MousePointer    =   99  'Custom
      TabIndex        =   91
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   885
      MouseIcon       =   "frmClrComposer.frx":13C6
      MousePointer    =   99  'Custom
      TabIndex        =   90
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   525
      MouseIcon       =   "frmClrComposer.frx":1518
      MousePointer    =   99  'Custom
      TabIndex        =   89
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   165
      MouseIcon       =   "frmClrComposer.frx":166A
      MousePointer    =   99  'Custom
      TabIndex        =   88
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblSelectTop 
      AutoSize        =   -1  'True
      Caption         =   "6"
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
      Left            =   120
      TabIndex        =   84
      Top             =   2010
      Width           =   180
   End
   Begin VB.Label lblSelectBottom 
      AutoSize        =   -1  'True
      Caption         =   "5"
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
      Left            =   120
      TabIndex        =   83
      Top             =   2520
      Width           =   180
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   2325
      MouseIcon       =   "frmClrComposer.frx":17BC
      MousePointer    =   99  'Custom
      TabIndex        =   79
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   1965
      MouseIcon       =   "frmClrComposer.frx":190E
      MousePointer    =   99  'Custom
      TabIndex        =   78
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   2325
      MouseIcon       =   "frmClrComposer.frx":1A60
      MousePointer    =   99  'Custom
      TabIndex        =   77
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   1965
      MouseIcon       =   "frmClrComposer.frx":1BB2
      MousePointer    =   99  'Custom
      TabIndex        =   76
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   1605
      MouseIcon       =   "frmClrComposer.frx":1D04
      MousePointer    =   99  'Custom
      TabIndex        =   75
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   1245
      MouseIcon       =   "frmClrComposer.frx":1E56
      MousePointer    =   99  'Custom
      TabIndex        =   74
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   885
      MouseIcon       =   "frmClrComposer.frx":1FA8
      MousePointer    =   99  'Custom
      TabIndex        =   73
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   525
      MouseIcon       =   "frmClrComposer.frx":20FA
      MousePointer    =   99  'Custom
      TabIndex        =   72
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   165
      MouseIcon       =   "frmClrComposer.frx":224C
      MousePointer    =   99  'Custom
      TabIndex        =   71
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   1605
      MouseIcon       =   "frmClrComposer.frx":239E
      MousePointer    =   99  'Custom
      TabIndex        =   70
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   1245
      MouseIcon       =   "frmClrComposer.frx":24F0
      MousePointer    =   99  'Custom
      TabIndex        =   69
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   885
      MouseIcon       =   "frmClrComposer.frx":2642
      MousePointer    =   99  'Custom
      TabIndex        =   68
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   525
      MouseIcon       =   "frmClrComposer.frx":2794
      MousePointer    =   99  'Custom
      TabIndex        =   67
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   165
      MouseIcon       =   "frmClrComposer.frx":28E6
      MousePointer    =   99  'Custom
      TabIndex        =   66
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   2325
      MouseIcon       =   "frmClrComposer.frx":2A38
      MousePointer    =   99  'Custom
      TabIndex        =   65
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1965
      MouseIcon       =   "frmClrComposer.frx":2B8A
      MousePointer    =   99  'Custom
      TabIndex        =   64
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1605
      MouseIcon       =   "frmClrComposer.frx":2CDC
      MousePointer    =   99  'Custom
      TabIndex        =   63
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1245
      MouseIcon       =   "frmClrComposer.frx":2E2E
      MousePointer    =   99  'Custom
      TabIndex        =   62
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   885
      MouseIcon       =   "frmClrComposer.frx":2F80
      MousePointer    =   99  'Custom
      TabIndex        =   61
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   525
      MouseIcon       =   "frmClrComposer.frx":30D2
      MousePointer    =   99  'Custom
      TabIndex        =   60
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblBasicColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   165
      MouseIcon       =   "frmClrComposer.frx":3224
      MousePointer    =   99  'Custom
      TabIndex        =   59
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblSymbol 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   5625
      TabIndex        =   28
      Top             =   4800
      Width           =   165
   End
   Begin VB.Label lblSymbol 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   5625
      TabIndex        =   27
      Top             =   945
      Width           =   180
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   105
      X2              =   2650
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   105
      X2              =   2650
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   5835
      X2              =   5835
      Y1              =   -90
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   5835
      X2              =   5835
      Y1              =   -90
      Y2              =   6480
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   75
      TabIndex        =   85
      Top             =   3240
      Width           =   2595
   End
   Begin VB.Menu mnuH 
      Caption         =   "mnuH"
      Visible         =   0   'False
      Begin VB.Menu mnuAddItem 
         Caption         =   "Add Color"
      End
      Begin VB.Menu mnuClearItem 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuH2 
      Caption         =   "mnuH2"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd2 
         Caption         =   "Add Color"
      End
      Begin VB.Menu mnuClear2 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuH3 
      Caption         =   "mnuH3"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd3 
         Caption         =   "Add Color"
      End
      Begin VB.Menu mnuClear3 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuH4 
      Caption         =   "mnuH4"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd4 
         Caption         =   "Add Color"
      End
      Begin VB.Menu mnuClear4 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuHEdit 
      Caption         =   "mnuHEdit"
      Visible         =   0   'False
      Begin VB.Menu mnuHCutItem 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuHCopyItem 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuHPasteItem 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuHDeleteItem 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmComposer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''Triangle Gradient
Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type
Private Type TRIVERTEX
    TRIX As Long
    TRIY As Long
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
Private Declare Function GradientFillTriangle Lib "msimg32" _
Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, _
ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, _
ByVal dwMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

''''''''''''''''''''Gradient
Const GRADIENT_FILL_OP_FLAG As Long = &HFF
Private Declare Function GdiGradientFillRect Lib "gdi32" Alias "GdiGradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Public Enum StartGradient
    FromRed = 1
    FromGreen = 2
    FromBlue = 3
    FromWhite = 4
    FromBlack = 5
End Enum
Public Enum EndGradient
    ToRed = 11
    ToGreen = 12
    ToBlue = 13
    ToWhite = 14
    ToBlack = 15
End Enum

Dim tempVal(1 To 3) As Integer

Private Sub cmdConvert_Click()

    picView.SetFocus
    
    Select Case True
    Case optFromRGB.Value
        On Error GoTo rgbError
        Dim tmpString
        tmpString = Trim(txtConvertTo.Text)
        If InStr(1, tmpString, ",", vbTextCompare) > 0 Then
            tmpString = Split(txtConvertTo.Text, ",", , vbTextCompare)
        ElseIf InStr(1, tmpString, " ", vbTextCompare) > 0 Then
            tmpString = Split(txtConvertTo.Text, " ", , vbTextCompare)
        End If
        Trim (tmpString(0))
        Trim (tmpString(1))
        Trim (tmpString(2))
        picConvert.BackColor = RGB(tmpString(0), tmpString(1), tmpString(2))
    Case optFromVB.Value
        On Error GoTo VBError
        txtConvertTo.Text = Trim(txtConvertTo.Text)
        If Right(txtConvertTo.Text, 1) = "&" Then txtConvertTo = Left(txtConvertTo.Text, Len(txtConvertTo.Text) - 1)
        If Left(txtConvertTo.Text, 1) <> "&" Then txtConvertTo.Text = "&" & txtConvertTo.Text
        If LCase(Mid(txtConvertTo.Text, 2, 1)) <> "h" Then txtConvertTo.Text = UCase("&H" & Right(txtConvertTo.Text, Len(txtConvertTo.Text) - 1))
        picConvert.BackColor = txtConvertTo.Text
    Case optFromHEX.Value
        On Error GoTo HEXError
        Dim tmpHEX, tR, tG, tB
        txtConvertTo.Text = Trim(txtConvertTo.Text)
        If Right(txtConvertTo.Text, 1) = "&" Then txtConvertTo = Left(txtConvertTo.Text, Len(txtConvertTo.Text) - 1)
        tmpHEX = Right("000000" & txtConvertTo.Text, 6)
        txtConvertTo.Text = UCase(tmpHEX)
        tR = Val("&h" & Left(tmpHEX, 2))
        tG = Val("&h" & Mid(tmpHEX, 3, 2))
        tB = Val("&h" & Right(tmpHEX, 2))
        picConvert.BackColor = RGB(tR, tG, tB)
    Case optFromLong.Value
        On Error GoTo LongError
        picConvert.BackColor = Trim(CLng(txtConvertTo.Text))
    End Select
    
    Select Case True
    Case optToRGB.Value
        Dim mR, mG, mB
        
        deLong picConvert.BackColor, mR, mG, mB
        txtConverted.Text = mR & "," & mG & "," & mB
    Case optToHEX.Value
        Dim aTemp
        
        aTemp = Right$("000000" & Hex(picConvert.BackColor), 6)
        txtConverted = UCase(Right$(aTemp, 2) & Mid$(aTemp, 3, 2) & Left$(aTemp, 2))
    Case optToVB.Value
        Dim RHex, GHex, BHex, fR, fG, fB
        Dim vbStr As String
        
        deLong picConvert.BackColor, fR, fG, fB
        RHex = Hex(fR)
        If Len(CStr(RHex)) < 2 Then RHex = "0" & RHex
        GHex = Hex(fG)
        If Len(CStr(GHex)) < 2 Then GHex = "0" & GHex
        BHex = Hex(fB)
        If Len(CStr(BHex)) < 2 Then BHex = "0" & BHex
        vbStr = "&H" & BHex & GHex & RHex & "&"
        txtConverted.Text = vbStr
    Case optToLong.Value
        txtConverted.Text = picConvert.BackColor
    End Select
    
    Exit Sub
    
VBError:
    txtConverted.Text = "HEX Error"
    Exit Sub
    
LongError:
    txtConverted.Text = "Long Error"
    Exit Sub
    
HEXError:
    txtConverted.Text = "VBHEX Error"
    Exit Sub
    
rgbError:
    txtConverted.Text = "RGB Error."
    Exit Sub

End Sub

Private Sub cmdCopyConverted_Click()

    picView.SetFocus

    cmdFlash.Top = (Me.Height / 2) - (cmdFlash.Height / 2)
    cmdFlash.Left = (Me.Width / 2) - (cmdFlash.Width / 2)
    Select Case True
    Case optToRGB.Value
        cmdFlash.Caption = "Copying RGB..."
    Case optToVB.Value
        cmdFlash.Caption = "Copying VBHEX..."
    Case optToHEX.Value
        cmdFlash.Caption = "Copying HEX..."
    Case optToLong.Value
        cmdFlash.Caption = "Copying Long..."
    End Select
    cmdFlash.Visible = True
    Clipboard.Clear: Clipboard.SetText txtConverted.Text
    Timer1.Enabled = True

End Sub

Private Sub cmdCopyLong_Click()

    picView.SetFocus

    cmdFlash.Top = (Me.Height / 2) - (cmdFlash.Height / 2)
    cmdFlash.Left = (Me.Width / 2) - (cmdFlash.Width / 2)
    cmdFlash.Caption = "Copying Long..."
    cmdFlash.Visible = True
    Clipboard.Clear: Clipboard.SetText Trim(txtLong.Text)
    Timer1.Enabled = True
    
End Sub

Private Sub cmdCopyRGB_Click()

    picView.SetFocus

    cmdFlash.Top = (Me.Height / 2) - (cmdFlash.Height / 2)
    cmdFlash.Left = (Me.Width / 2) - (cmdFlash.Width / 2)
    cmdFlash.Caption = "Copying RGB..."
    cmdFlash.Visible = True
    Clipboard.Clear: Clipboard.SetText Trim(txtRGB.Text)
    Timer1.Enabled = True
    
End Sub

Private Sub cmdCstmReset_Click()

    picView.SetFocus

    mnuClearItem_Click
    mnuClear2_Click
    mnuClear3_Click
    mnuClear4_Click
    
End Sub

Private Sub cmdEnd_Click()

    picView.SetFocus
    Unload Me

End Sub

Private Sub cmdFromPixel_Click()

   picView.SetFocus

   VSR = R
   VSG = G
   VSB = B
   
End Sub

Private Sub cmdMore_Click()

    picView.SetFocus

    If cmdMore.Tag = "more" Then
        cmdMore.Caption = "^"
        cmdMore.Tag = "less"
        lblSymbol(0).Caption = "3"
        lblSymbol(1).Caption = "3"
        Me.Width = Me.Width + 2400
    Else
        cmdMore.Caption = "_"
        cmdMore.Tag = "more"
        lblSymbol(0).Caption = "4"
        lblSymbol(1).Caption = "4"
        Me.Width = Me.Width - 2400
    End If

End Sub

Private Sub cmdCopyHEX_Click()

    picView.SetFocus

    cmdFlash.Top = (Me.Height / 2) - (cmdFlash.Height / 2)
    cmdFlash.Left = (Me.Width / 2) - (cmdFlash.Width / 2)
    cmdFlash.Caption = "Copying HEX..."
    cmdFlash.ZOrder
    cmdFlash.Visible = True
    Clipboard.Clear: Clipboard.SetText Trim(txtHEX.Text)
    Timer1.Enabled = True
    
End Sub

Private Sub cmdCopyVB_Click()

    picView.SetFocus
    
    cmdFlash.Top = (Me.Height / 2) - (cmdFlash.Height / 2)
    cmdFlash.Left = (Me.Width / 2) - (cmdFlash.Width / 2)
    cmdFlash.Caption = "Copying VBHEX..."
    cmdFlash.Visible = True
    Clipboard.Clear: Clipboard.SetText Trim(txtVB.Text)
    Timer1.Enabled = True
    
End Sub


Private Sub cmdOKHEX_Click()

    picView.SetFocus
    
    On Error GoTo erHand
    
    Dim tmpTxt
    
    txtEnterHEX.Text = Trim(txtEnterHEX.Text)
    If Right(txtEnterHEX.Text, 1) = "&" Then
        txtEnterHEX.Text = Left(txtEnterHEX.Text, Len(txtEnterHEX.Text) - 1)
    End If
    tmpTxt = Right("000000" & txtEnterHEX.Text, 6)
    
    VSR = Val("&h" & Left(tmpTxt, 2))
    VSG = Val("&h" & Mid(tmpTxt, 3, 2))
    VSB = Val("&h" & Right(tmpTxt, 2))
    
    Exit Sub
    
erHand:
    txtEnterHEX.Text = "Error"
    Exit Sub

End Sub

Private Sub cmdOKLong_Click()

    picView.SetFocus

    Dim tR, tG, tB
    
    On Error GoTo erHand
    
    picView.BackColor = txtEnterLong.Text
    deLong picView.BackColor, tR, tG, tB
    VSR = tR
    VSG = tG
    VSB = tB
    
    Exit Sub
    
erHand:
    txtEnterLong.Text = "Error"
    
    Exit Sub

End Sub

Private Sub cmdOKVB_Click()

    picView.SetFocus
    
    Dim tR, tG, tB
    
    On Error GoTo erHand
    
    txtEnterVB.Text = Trim(txtEnterVB.Text)
    If Right(txtEnterVB.Text, 1) = "&" Then
        txtEnterVB.Text = Left(txtEnterVB.Text, Len(txtEnterVB.Text) - 1)
    End If
    picView.BackColor = txtEnterVB.Text
    deLong picView.BackColor, tR, tG, tB
    VSR = tR
    VSG = tG
    VSB = tB
    
    Exit Sub
    
erHand:
    txtEnterVB.Text = "Error"
    
    Exit Sub

End Sub

Private Sub cmdPalette_Click()

    picView.SetFocus
    
    Dim bTmp As Boolean
    
    If frmMain.Timer1.Enabled = True Then
        bTmp = True
    Else
        bTmp = False
    End If

    cmnDlg1.Color = RGB(VSR, VSG, VSB)
    cmnDlg1.Flags = &H2 Or &H1
    cmnDlg1.ShowColor
    
    Dim tmpClr
    
    tmpClr = Right("000000" & Hex(cmnDlg1.Color), 6)
    VSR = Val("&h" & Right(tmpClr, 2))
    VSG = Val("&h" & Mid(tmpClr, 3, 2))
    VSB = Val("&h" & Left(tmpClr, 2))
    
    If bTmp = True Then
        frmMain.cmdRestart_Click
    End If
    
End Sub

Private Sub cmdRB_Click()

    picView.SetFocus
    VSB.Value = Int(255 * Rnd + 1)

End Sub

Private Sub cmdRG_Click()

    picView.SetFocus
    VSG.Value = Int(255 * Rnd + 1)

End Sub

Private Sub cmdRnd_Click()

    picView.SetFocus
    
    Randomize Timer
    
    VSR = Int(255 * Rnd + 1)
    VSG = Int(255 * Rnd + 1)
    VSB = Int(255 * Rnd + 1)

End Sub

Private Sub cmdRR_Click()

    picView.SetFocus
    VSR.Value = Int(255 * Rnd + 1)

End Sub

Private Sub cmdToggleH_Click()

    picView.SetFocus
    
    If cmdToggleH.Tag = "b" Then
        cmdToggleH.Tag = "p"
        picHide.Visible = False
        cmdToggleH.Caption = "Palette"
    Else
        cmdToggleH.Tag = "b"
        picHide.Visible = True
        cmdToggleH.Caption = "Basic Colors"
    End If
    

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        VSR.Value = CInt(txtR.Text)
        VSG.Value = CInt(txtG.Text)
        VSB.Value = CInt(txtB.Text)
    End If

End Sub

Private Sub Form_Load()

    Dim rH, rW
    Dim ranR, ranB, ranG
    
    Randomize Timer

    rW = Int((picGradient.Width - 100) * Rnd + 1)
    rH = Int((picGradient.Height - 75) * Rnd + 1)

    'picHide.Left = 0
    'picHide.Top = 4200

    isComposerActive = True
    setTopMost2 Me
    Show
    activeWinComposer = GetActiveWindow
    
    ranR = Int(255 * Rnd + 1)
    ranG = Int(255 * Rnd + 1)
    ranB = Int(255 * Rnd + 1)
    
    txtConvertTo.Text = ranR & " " & ranG & " " & ranB
    optToHEX_Click
    
    Me.Width = 5880
    Me.Height = 6255
    
    picView.BackColor = RGB(VSR, VSG, VSB)
    txtVB.Text = "&H" & Hex(picView.BackColor)
    
    getLocalReg
    
    picTemp.BackColor = picGradient.Point(rW, rH)
    
    If isPixActive Then
        frmComposer.cmdFromPixel.Enabled = True
    Else
        frmComposer.cmdFromPixel.Enabled = False
    End If
    
    cmdFlash.ZOrder
    
    setTriGradient
    setCustomCMD

End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Dim isMore As Boolean
    
    activeWinComposer = -1
    If cmdMore.Tag = "less" Then
        isMore = True
    Else
        isMore = False
    End If
    SaveSetting APP_NAME, SECTION_NAME, "ComposerShowMore", isMore
    SaveSetting APP_NAME, SECTION_NAME, "ComposerStartColor", picView.BackColor
    SaveSetting APP_NAME, SECTION_NAME, "ComposerStartPalette", picHide.Visible
    
    frmMain.mnuClearSetting.Enabled = True
    frmMain.mnuComposerItem.Checked = False
    isComposerActive = False

End Sub

Private Sub lblBasicColors_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        picView.BackColor = lblBasicColors(Index).BackColor
        Dim rV, gV, bV
        deLong picView.BackColor, rV, gV, bV
        VSR.Value = rV
        VSG.Value = gV
        VSB.Value = bV
    End If

End Sub

Private Sub lblBasicColors_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    picTemp.BackColor = lblBasicColors(Index).BackColor

End Sub

Private Sub lblSymbol_Click(Index As Integer)

    cmdMore_Click

End Sub

Private Sub mnuAddItem_Click()

    Picture1.BackColor = picView.BackColor
    setLocalReg

End Sub

Private Sub mnuClearItem_Click()

    Picture1.BackColor = &HC8D0D4
    setLocalReg

End Sub

Private Sub mnuHCopyItem_Click()

    Clipboard.Clear
    Clipboard.SetText ActiveControl.Text

End Sub

Private Sub mnuHCutItem_Click()

    Clipboard.Clear
    Clipboard.SetText ActiveControl.Text
    ActiveControl.Text = ""

End Sub

Private Sub mnuHDeleteItem_Click()

    ActiveControl.Text = ""

End Sub


Private Sub mnuHPasteItem_Click()

    ActiveControl.Text = Clipboard.GetText

End Sub

Private Sub optFromHEX_Click()

    'txtConvertTo.Text = "0"
    txtConvertTo.SetFocus

End Sub

Private Sub optFromLong_Click()

    'txtConvertTo.Text = "0"
    txtConvertTo.SetFocus

End Sub

Private Sub optFromRGB_Click()

    'txtConvertTo.Text = "0,0,0"
    txtConvertTo.SetFocus

End Sub

Private Sub optFromVB_Click()

    'txtConvertTo.Text = "0"
    txtConvertTo.SetFocus

End Sub

Private Sub optToHEX_Click()

    cmdConvert_Click

End Sub

Private Sub optToLong_Click()

    cmdConvert_Click

End Sub

Private Sub optToRGB_Click()

    cmdConvert_Click

End Sub

Private Sub optToVB_Click()

    cmdConvert_Click

End Sub

Private Sub picConvert_Click()

    picView.BackColor = picConvert.BackColor
    Dim rV, gV, bV
    deLong picView.BackColor, rV, gV, bV
    VSR.Value = rV
    VSG.Value = gV
    VSB.Value = bV

End Sub

Private Sub picConvert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picTemp.BackColor = picConvert.BackColor

End Sub

Private Sub picGradient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        picView.BackColor = picGradient.Point(X, Y)
        Dim rV, gV, bV
        deLong picView.BackColor, rV, gV, bV
        VSR.Value = rV
        VSG.Value = gV
        VSB.Value = bV
    End If

End Sub

Private Sub picGradient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picTemp.BackColor = picGradient.Point(X, Y)
    lblSelectBottom.Left = X + 50
    lblSelectTop.Left = X + 50

End Sub

Private Sub picMinPalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picView.BackColor = picMinPalette.Point(X, Y)

End Sub

Private Sub picMinPalette_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picTemp.BackColor = picMinPalette.Point(X, Y)
    Dim rV, gV, bV
    deLong picView.BackColor, rV, gV, bV
    VSR.Value = rV
    VSG.Value = gV
    VSB.Value = bV

End Sub

Private Sub picOB_Click()

    picView.BackColor = picOB.BackColor
    
    Dim rV, gV, bV
    deLong picView.BackColor, rV, gV, bV
    VSR.Value = rV
    VSG.Value = gV
    VSB.Value = bV


End Sub

Private Sub picOB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picTemp.BackColor = picOB.BackColor

End Sub

Private Sub picOG_Click()

    picView.BackColor = picOG.BackColor
    
    Dim rV, gV, bV
    deLong picView.BackColor, rV, gV, bV
    VSR.Value = rV
    VSG.Value = gV
    VSB.Value = bV


End Sub

Private Sub picOG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picTemp.BackColor = picOG.BackColor


End Sub

Private Sub picOR_Click()

    picView.BackColor = picOR.BackColor
    
    Dim rV, gV, bV
    deLong picView.BackColor, rV, gV, bV
    VSR.Value = rV
    VSG.Value = gV
    VSB.Value = bV
    
End Sub

Private Sub picOR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picTemp.BackColor = picOR.BackColor

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuH
    Else
        Dim sTmp1
        picView.BackColor = Picture1.BackColor
        sTmp1 = Right$("000000" & Hex(Picture1.BackColor), 6)
        VSR.Value = Val("&h" & Right(sTmp1, 2))
        VSG.Value = Val("&h" & Mid(sTmp1, 3, 2))
        VSB.Value = Val("&h" & Left(sTmp1, 2))
    End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picTemp.BackColor = Picture1.BackColor

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picTemp.BackColor = Picture2.BackColor

End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picTemp.BackColor = Picture3.BackColor

End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picTemp.BackColor = Picture4.BackColor

End Sub

Private Sub Timer1_Timer()

    Timer1.Enabled = False
    cmdFlash.Visible = False

End Sub


Private Sub txtB_Change()

    checkValue txtB

End Sub

Private Sub txtB_GotFocus()

    txtGotFocus txtB, tempVal(3)

End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)

    KeyAscii = blockChars(txtB, KeyAscii)

End Sub

Private Sub txtB_LostFocus()

    txtLostFocus txtB, tempVal(3)

End Sub


Private Sub txtConverted_GotFocus()

    txtConverted.BackColor = &HC0C0C0
    txtConverted.SelStart = 0
    txtConverted.SelLength = Len(txtConverted.Text)

End Sub

Private Sub txtConverted_LostFocus()

    txtConverted.BackColor = vbWhite

End Sub

Private Sub txtConverted_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then hiddenMenuPopup False

End Sub

Private Sub txtConvertTo_GotFocus()

    txtConvertTo.BackColor = RGB(255, 255, 164)
    txtConvertTo.SelStart = 0
    txtConvertTo.SelLength = Len(txtConvertTo.Text)

End Sub

Private Sub txtConvertTo_LostFocus()

    txtConvertTo.BackColor = vbWhite

End Sub

Private Sub txtConvertTo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then hiddenMenuPopup True

End Sub

Private Sub txtEnterHEX_GotFocus()

    txtEnterHEX.BackColor = RGB(255, 255, 164)
    Text2.BackColor = RGB(255, 255, 164)
    txtEnterHEX.SelStart = 0
    txtEnterHEX.SelLength = Len(txtHEX.Text)

End Sub

Private Sub txtEnterHEX_LostFocus()

    txtEnterHEX.BackColor = vbWhite
    Text2.BackColor = vbWhite

End Sub

Private Sub txtEnterHEX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then hiddenMenuPopup True

End Sub

Private Sub txtEnterLong_GotFocus()

    txtEnterLong.BackColor = RGB(255, 255, 164)
    Text7.BackColor = RGB(255, 255, 164)
    txtEnterLong.SelStart = 0
    txtEnterLong.SelLength = Len(txtEnterLong.Text)
    

End Sub

Private Sub txtEnterLong_LostFocus()

    txtEnterLong.BackColor = vbWhite
    Text7.BackColor = vbWhite

End Sub

Private Sub txtEnterLong_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then hiddenMenuPopup True

End Sub

Private Sub txtEnterVB_GotFocus()

    txtEnterVB.BackColor = RGB(255, 255, 164)
    Text3.BackColor = RGB(255, 255, 164)
    txtEnterVB.SelStart = 0
    txtEnterVB.SelLength = Len(txtEnterVB.Text)

End Sub

Private Sub txtEnterVB_LostFocus()

    txtEnterVB.BackColor = vbWhite
    Text3.BackColor = vbWhite

End Sub

Private Sub txtEnterVB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then hiddenMenuPopup True

End Sub

Private Sub txtG_Change()

    checkValue txtG

End Sub

Private Sub txtG_GotFocus()

    txtGotFocus txtG, tempVal(2)

End Sub

Private Sub txtG_KeyPress(KeyAscii As Integer)

    KeyAscii = blockChars(txtG, KeyAscii)

End Sub

Private Sub txtG_LostFocus()

    txtLostFocus txtG, tempVal(2)

End Sub

Private Sub txtHEX_GotFocus()
    
    txtHEX.SelStart = 0
    txtHEX.SelLength = Len(txtHEX.Text)

End Sub

Private Sub txtR_Change()

    checkValue txtR

End Sub

Private Sub txtR_GotFocus()

    txtGotFocus txtR, tempVal(1)
    
End Sub

Private Sub txtR_KeyPress(KeyAscii As Integer)

    KeyAscii = blockChars(txtR, KeyAscii)

End Sub

Private Sub txtR_LostFocus()

    txtLostFocus txtR, tempVal(1)

End Sub

Private Sub txtRGB_GotFocus()

    txtRGB.SelStart = 0
    txtRGB.SelLength = Len(txtRGB.Text)

End Sub

Private Sub txtVB_GotFocus()

    txtVB.SelStart = 0
    txtVB.SelLength = Len(txtVB.Text)

End Sub

Private Sub VSB_Change()

    changeColor

End Sub

Private Sub VSB_Scroll()

    changeColor

End Sub

Private Sub VSG_Change()

    changeColor

End Sub

Private Sub VSG_Scroll()

    changeColor

End Sub

Private Sub VSR_Change()

    changeColor

End Sub

Sub changeColor()

    Dim sTmp, theHex

    picView.BackColor = RGB(VSR, VSG, VSB)
    txtR.Text = VSR.Value: txtR.Tag = "t"
    txtG.Text = VSG.Value: txtG.Tag = "t"
    txtB.Text = VSB.Value: txtB.Tag = "t"
    txtRGB.Text = Trim(txtR.Text) & " " & Trim(txtG.Text) & " " & Trim(txtB.Text)
    txtVB.Text = "&H" & Hex(picView.BackColor)
    
    txtLong.Text = picView.BackColor
    txtEnterVB.Text = txtVB.Text
    txtEnterLong.Text = picView.BackColor
    
    VStxtR.Value = VSR.Value
    VStxtG.Value = VSG.Value
    VStxtB.Value = VSB.Value
    
    sTmp = Right$("000000" & Hex(picView.BackColor), 6)
    theHex = Right$(sTmp, 2) & Mid$(sTmp, 3, 2) & Left$(sTmp, 2)
    txtHEX.Text = theHex
    
    txtEnterHEX.Text = txtHEX.Text
    
    'SaveSetting APP_NAME, SECTION_NAME, "ComposerStartColor", picView.BackColor
    picOR.BackColor = RGB(VSR.Value, 0, 0)
    picOG.BackColor = RGB(0, VSG.Value, 0)
    picOB.BackColor = RGB(0, 0, VSB.Value)
    SetGradient

End Sub

Private Sub VSR_Scroll()

    changeColor

End Sub

Private Function blockChars(txt As TextBox, Key As Integer)

    If Key = 8 Then blockChars = Key: txt.Tag = "t": Exit Function
    If Key < 48 Or Key > 57 Then
        blockChars = 0
        txt.Tag = "f"
    Else
        blockChars = Key
        txt.Tag = "t"
    End If

End Function

Sub txtGotFocus(txt As TextBox, tmp As Integer)

    If txt.Text = "" Then txt.Text = "0"
    tmp = CInt(txt.Text)
    txt.Text = ""
    txt.Tag = "f"

End Sub

Sub txtLostFocus(txt As TextBox, tmp As Integer)

    If txt.Tag = "f" Then
        txt.Text = tmp
    End If
    
    Form_KeyDown 13, 0

End Sub

Sub checkValue(txt As TextBox)

    On Error GoTo erHand
    If CInt(txt.Text) > 255 Then
        txt.Text = "255"
        txt.SelStart = Len(txt.Text)
    End If
    
    Exit Sub
    
erHand:
    Exit Sub

End Sub

Private Sub mnuAdd2_Click()

    Picture2.BackColor = picView.BackColor
    setLocalReg

End Sub

Private Sub mnuClear2_Click()

    Picture2.BackColor = &HC8D0D4
    setLocalReg

End Sub


Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuH2
    Else
        Dim sTmp1
        picView.BackColor = Picture2.BackColor
        sTmp1 = Right$("000000" & Hex(Picture2.BackColor), 6)
        VSR.Value = Val("&h" & Right(sTmp1, 2))
        VSG.Value = Val("&h" & Mid(sTmp1, 3, 2))
        VSB.Value = Val("&h" & Left(sTmp1, 2))
    End If

End Sub

Private Sub mnuAdd3_Click()

    Picture3.BackColor = picView.BackColor
    setLocalReg

End Sub

Private Sub mnuClear3_Click()

    Picture3.BackColor = &HC8D0D4
    setLocalReg

End Sub


Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuH3
    Else
        Dim sTmp1
        picView.BackColor = Picture3.BackColor
        sTmp1 = Right$("000000" & Hex(Picture3.BackColor), 6)
        VSR.Value = Val("&h" & Right(sTmp1, 2))
        VSG.Value = Val("&h" & Mid(sTmp1, 3, 2))
        VSB.Value = Val("&h" & Left(sTmp1, 2))
    End If

End Sub

Private Sub mnuAdd4_Click()

    Picture4.BackColor = picView.BackColor
    setLocalReg

End Sub

Private Sub mnuClear4_Click()

    Picture4.BackColor = &HC8D0D4
    setLocalReg

End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuH4
    Else
        Dim sTmp1
        picView.BackColor = Picture4.BackColor
        sTmp1 = Right$("000000" & Hex(Picture4.BackColor), 6)
        VSR.Value = Val("&h" & Right(sTmp1, 2))
        VSG.Value = Val("&h" & Mid(sTmp1, 3, 2))
        VSB.Value = Val("&h" & Left(sTmp1, 2))
    End If

End Sub

Public Sub getLocalReg()

    Dim isMore As Boolean
    Dim isPalette As Boolean
    Dim tmpPicView

    Picture1.BackColor = GetSetting(APP_NAME, SECTION_NAME, "ComposerColor1", &HC8D0D4)
    Picture2.BackColor = GetSetting(APP_NAME, SECTION_NAME, "ComposerColor2", &HC8D0D4)
    Picture3.BackColor = GetSetting(APP_NAME, SECTION_NAME, "ComposerColor3", &HC8D0D4)
    Picture4.BackColor = GetSetting(APP_NAME, SECTION_NAME, "ComposerColor4", &HC8D0D4)
    isMore = GetSetting(APP_NAME, SECTION_NAME, "ComposerShowMore", "False")
    tmpPicView = GetSetting(APP_NAME, SECTION_NAME, "ComposerStartColor", "-1")
    isPalette = GetSetting(APP_NAME, SECTION_NAME, "ComposerStartPalette", "True")
    
    If tmpPicView = "-1" Then
        VSR = Int(255 * Rnd + 1)
        VSG = Int(255 * Rnd + 1)
        VSB = Int(255 * Rnd + 1)
    Else
        picView.BackColor = tmpPicView
    End If
    
    If isPalette = False Then
        cmdToggleH_Click
    End If

    Dim lngTmp
    lngTmp = Right$("000000" & Hex(picView.BackColor), 6)
    VSR.Value = Val("&h" & Right(lngTmp, 2))
    VSG.Value = Val("&h" & Mid(lngTmp, 3, 2))
    VSB.Value = Val("&h" & Left(lngTmp, 2))

    If isMore Then
        cmdMore_Click
    End If

End Sub

Public Sub setLocalReg()

    SaveSetting APP_NAME, SECTION_NAME, "ComposerColor1", Picture1.BackColor
    SaveSetting APP_NAME, SECTION_NAME, "ComposerColor2", Picture2.BackColor
    SaveSetting APP_NAME, SECTION_NAME, "ComposerColor3", Picture3.BackColor
    SaveSetting APP_NAME, SECTION_NAME, "ComposerColor4", Picture4.BackColor

    frmMain.mnuClearSetting.Enabled = True

End Sub

Private Sub VStxtB_Change()

    VSB.Value = VStxtB.Value

End Sub

Private Sub VStxtG_Change()

    VSG.Value = VStxtG.Value

End Sub

Private Sub VStxtR_Change()

    VSR.Value = VStxtR.Value

End Sub

Private Sub SetGradient()

'    picGradient.ScaleMode = vbPixels
'    picGradient.DrawWidth = 1
'
'    For i = 0 To picGradient.Width
'        picGradient.Line (i, 0)-(i, picGradient.ScaleWidth), picView.BackColor + i
'    Next i

'    On Error Resume Next
'
'    picGradient.ScaleMode = vbPixels
'
'    For i = 0 To picGradient.Width
'        deLong (picView.BackColor)
'        picGradient.Line (i, 0)-(i, picGradient.ScaleWidth), RGB(R - i, G - i, B)  'Picture2.BackColor + i
'    Next i

    Dim picH&, picW&
    Dim R2 As Double
    Dim G2 As Double
    Dim B2 As Double
    Dim rCol, gCol, bCol
    Dim i
    Dim R%, G%, B%


    picH = picGradient.Height
    picW = picGradient.Width
    
    deLong picView.BackColor, R, G, B
    R2 = (0 - R) / picW
    G2 = (0 - G) / picW
    B2 = (0 - B) / picW

    For i = 0 To picW
        rCol = Round(R + (R2 * (i + 1)))
        gCol = Round(G + (G2 * (i + 1)))
        bCol = Round(B + (B2 * (i + 1)))
        picGradient.Line (i, 0)-(i, picH), RGB(rCol, gCol, bCol)
    Next i

End Sub

Sub deLong(ByRef aCLR As Long, aR, aG, aB)

    Dim tmp
    
    tmp = Right("000000" & Hex(aCLR), 6)
    aR = Val("&H" & Right(tmp, 2))
    aG = Val("&H" & Mid(tmp, 3, 2))
    aB = Val("&H" & Left(tmp, 2))
    
End Sub

Sub setTriGradient()

    Dim vert(4) As TRIVERTEX
    Dim gTRi(1) As GRADIENT_TRIANGLE
    picMinPalette.ScaleMode = vbPixels
    vert(0).TRIX = 0
    vert(0).TRIY = 0
    vert(0).Red = -256
    vert(0).Green = 0&
    vert(0).Blue = 0&
    vert(0).Alpha = 0&
    
    vert(1).TRIX = 130
    vert(1).TRIY = 0
    vert(1).Red = 0&
    vert(1).Green = -256
    vert(1).Blue = 0&
    vert(1).Alpha = 0&
    
    vert(2).TRIX = 131
    vert(2).TRIY = 130
    vert(2).Red = 0&
    vert(2).Green = 0&
    vert(2).Blue = -256
    vert(2).Alpha = 0&
    
    vert(3).TRIX = 0
    vert(3).TRIY = 130
    vert(3).Red = -256
    vert(3).Green = -256
    vert(3).Blue = -256
    vert(3).Alpha = 0&
    
    gTRi(0).Vertex1 = 0
    gTRi(0).Vertex2 = 1
    gTRi(0).Vertex3 = 2
    
    gTRi(1).Vertex1 = 0
    gTRi(1).Vertex2 = 2
    gTRi(1).Vertex3 = 3
    GradientFillTriangle picMinPalette.hdc, vert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE
    
End Sub

Private Sub setCustomCMD()

    ''''''Random Red CMD
    makeGradient picRR, FromWhite, ToRed
    picRR.CurrentX = 6: picRR.CurrentY = 5: picRR.ForeColor = vbWhite
    picRR.FontBold = True
    picRR.Print "R"
    cmdRR.Picture = picRR.Image
    
    ''''''Random Green CMD
    makeGradient picRG, FromWhite, ToGreen
    picRG.CurrentX = 6: picRG.CurrentY = 5: picRG.ForeColor = vbWhite
    picRG.FontBold = True
    picRG.Print "G"
    cmdRG.Picture = picRG.Image
    
    ''''''Random Blue CMD
    makeGradient picRB, FromWhite, ToBlue
    picRB.CurrentX = 6: picRB.CurrentY = 5: picRB.ForeColor = vbWhite
    picRB.FontBold = True
    picRB.Print "B"
    cmdRB.Picture = picRB.Image
    
    ''''''''Random-RNDCMD
    Dim vert(4) As TRIVERTEX
    Dim gTRi(1) As GRADIENT_TRIANGLE
    picRnd.ScaleMode = vbPixels
    picRnd.AutoRedraw = True
    vert(0).TRIX = 0
    vert(0).TRIY = 0
    vert(0).Red = -256
    vert(0).Green = 0&
    vert(0).Blue = 0&
    vert(0).Alpha = 0&
    
    vert(1).TRIX = picRnd.ScaleWidth
    vert(1).TRIY = 0
    vert(1).Red = 0&
    vert(1).Green = -256
    vert(1).Blue = 0&
    vert(1).Alpha = 0&
    
    vert(2).TRIX = picRnd.ScaleWidth
    vert(2).TRIY = picRnd.ScaleHeight
    vert(2).Red = 0&
    vert(2).Green = 0&
    vert(2).Blue = -256
    vert(2).Alpha = 0&
    
    vert(3).TRIX = 0
    vert(3).TRIY = picRnd.ScaleHeight
    vert(3).Red = -256
    vert(3).Green = -256
    vert(3).Blue = -256
    vert(3).Alpha = 0&
    
    gTRi(0).Vertex1 = 0
    gTRi(0).Vertex2 = 1
    gTRi(0).Vertex3 = 2
    
    gTRi(1).Vertex1 = 0
    gTRi(1).Vertex2 = 2
    gTRi(1).Vertex3 = 3
    GradientFillTriangle picRnd.hdc, vert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE
    picRnd.Refresh
    
    picRnd.CurrentX = -1: picRnd.CurrentY = 5
    picRnd.ForeColor = vbWhite
    picRnd.FontBold = True
    picRnd.Print "Rnd"
    cmdRnd.Picture = picRnd.Image

End Sub

Public Sub makeGradient(Cntrl, StartColor As StartGradient, _
                       EndColor As EndGradient, Optional StartXPixel _
                       As Single, Optional StartYPixel As Single, _
                       Optional EndXPixel As Single, Optional EndYPixel As Single)
    
    Dim tmpScaleMode As ScaleModeConstants
    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    Dim gradR&, gradG&, gradB&
    Dim eX As Single, eY As Single
    
    On Error GoTo erHand
    
    tmpScaleMode = Cntrl.ScaleMode
    Cntrl.ScaleMode = 3
    Cntrl.AutoRedraw = True
    
    If EndXPixel > 0 Then
        eX = EndXPixel
    Else
        eX = Cntrl.ScaleWidth
    End If
    
    If EndYPixel > 0 Then
        eY = EndYPixel
    Else
        eY = Cntrl.ScaleHeight
    End If
    
    gradR = 0&
    gradG = 0&
    gradB = 0&
    Select Case StartColor
    Case 1
        gradR = LongToUShort(&HFF00&)
    Case 2
        gradG = LongToUShort(&HFF00&)
    Case 3
        gradB = LongToUShort(&HFF00&)
    Case 4
        gradR = LongToUShort(&HFF00&)
        gradG = LongToUShort(&HFF00&)
        gradB = LongToUShort(&HFF00&)
    Case 5
        gradR = 0&
        gradG = 0&
        gradB = 0&
    End Select
    
    'StartColor
    With vert(0)
        .TRIX = StartXPixel
        .TRIY = StartYPixel
        .Red = gradR 'LongToUShort(&HFF00&) '0&
        .Green = gradG 'LongToUShort(&HFF00&) '0& '&HFF&   '0&
        .Blue = gradB 'LongToUShort(&HFF00&) '0&
        .Alpha = 0&
    End With
    'EndColor
    gradR = 0&
    gradG = 0&
    gradB = 0&
    Select Case EndColor
    Case 11
        gradR = LongToUShort(&HFF00&)
    Case 12
        gradG = LongToUShort(&HFF00&)
    Case 13
        gradB = LongToUShort(&HFF00&)
    Case 14
        gradR = LongToUShort(&HFF00&)
        gradG = LongToUShort(&HFF00&)
        gradB = LongToUShort(&HFF00&)
    Case 15
        gradR = 0&
        gradG = 0&
        gradB = 0&
    End Select
    
    With vert(1)
        .TRIX = eX   'Cntrl.ScaleWidth 'Me.ScaleWidth
        .TRIY = eY   'Cntrl.ScaleHeight ' Me.ScaleHeight
        .Red = gradR ' LongToUShort(&HAA00&) ' 0&
        .Green = gradG '0&    'LongToUShort(&HAA00&)  '0&
        .Blue = gradB '0&
        .Alpha = 0&
    End With
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    GdiGradientFillRect Cntrl.hdc, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V
    Cntrl.ScaleMode = tmpScaleMode
    
    Exit Sub
    
erHand:
    If Err.Number = 438 Then
        Exit Sub
    Else
        Dim theReply
        theReply = MsgBox("Error interrupted the SetGradient Function" & _
        vbCrLf & Err.Description & vbCrLf & _
        "Error Number : " & Err.Number & vbCrLf & _
        "Do you want to continue? Yes-Continue/No-End", vbYesNo, "Error Occured")
        If theReply = vbYes Then Exit Sub
        End
    End If
    Exit Sub
    
End Sub

Private Function LongToUShort(Unsigned As Long) As Integer

    LongToUShort = CInt(Unsigned - &H10000)
    
End Function

Private Sub hiddenMenuPopup(isEdit As Boolean)
    
    Exit Sub
    
    If Len(ActiveControl.Text) > 0 Then
        mnuHCutItem.Enabled = True
        mnuHCopyItem.Enabled = True
        mnuHDeleteItem.Enabled = True
    Else
        mnuHCutItem.Enabled = False
        mnuHCopyItem.Enabled = False
        mnuHDeleteItem.Enabled = False
    End If
    
    If isEdit = True Then
        mnuHCutItem.Enabled = True
        mnuHDeleteItem.Enabled = True
        mnuHPasteItem.Enabled = True
    Else
        mnuHCutItem.Enabled = False
        mnuHDeleteItem.Enabled = False
        mnuHPasteItem.Enabled = False
    End If
    
    If Len(Clipboard.GetText) > 0 Then
        mnuHPasteItem.Enabled = True
    Else
        mnuHPasteItem.Enabled = False
    End If
    
    PopupMenu mnuHEdit
    
End Sub

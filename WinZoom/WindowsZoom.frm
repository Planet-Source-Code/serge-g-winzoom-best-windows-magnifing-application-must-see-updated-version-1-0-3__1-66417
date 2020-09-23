VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   Caption         =   "WinZoom"
   ClientHeight    =   6030
   ClientLeft      =   2100
   ClientTop       =   1695
   ClientWidth     =   7950
   Icon            =   "WindowsZoom.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7950
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1065
      ScaleWidth      =   2385
      TabIndex        =   12
      Top             =   25000
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To freeze the screen press and hold 'Control + F'. Click on me to to disable this message (It will not show again)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   2445
      End
   End
   Begin VB.PictureBox picZoom 
      AutoRedraw      =   -1  'True
      Height          =   5655
      Left            =   90
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   515
      TabIndex        =   0
      Top             =   0
      Width           =   7785
      Begin VB.TextBox txtFlashMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3990
         TabIndex        =   17
         Top             =   7500
         Width           =   1470
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2880
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   23
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":0CCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":1264
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":17FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":1D98
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":1EF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":248C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":25E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":2EC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":345A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":F22F
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":F389
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":F923
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":FEBD
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":10457
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":105B1
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":1070B
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":110E5
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":119BF
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":12299
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":12B73
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":1344D
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":1401F
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WindowsZoom.frx":148F9
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.DirListBox Dir1 
         Height          =   315
         Left            =   3000
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   3120
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblFlashMSG 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   390
         Left            =   3555
         TabIndex        =   16
         Tag             =   "f"
         Top             =   7500
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblCursorPos 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Image imgAim 
         Height          =   480
         Left            =   2880
         Picture         =   "WindowsZoom.frx":151D3
         Top             =   1680
         Width           =   480
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   5700
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Zoom In"
            Object.Tag             =   "ZoomIn"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   420
         ScaleHeight     =   195
         ScaleWidth      =   690
         TabIndex        =   5
         Top             =   45
         Width           =   750
         Begin VB.Label lblzoom 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Height          =   210
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   705
         End
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   1125
         TabIndex        =   2
         Top             =   0
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Zoom Out"
               Object.Tag             =   "ZoomOut"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Save Image As..."
               Object.Tag             =   "save"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Open Image"
               Object.Tag             =   "open"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copy Image to Clipboard"
               Object.Tag             =   "copy"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Quick Zoom"
               Object.Tag             =   "QZoom"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Open MSPaint"
               Object.Tag             =   "MSPaint"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Show Pixel Color"
               Object.Tag             =   "pixel"
               ImageIndex      =   22
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Close Pixel Color"
               Object.Tag             =   "pixelClose"
               ImageIndex      =   23
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Open Color Composer"
               Object.Tag             =   "openComposer"
               ImageIndex      =   21
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   330
            Left            =   4380
            TabIndex        =   4
            Top             =   0
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Close Program"
                  Object.Tag             =   "close"
                  ImageIndex      =   10
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.ToolTipText     =   "Close Image Preview Window"
                  Object.Tag             =   "closeImage"
                  ImageIndex      =   13
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Info"
                  Object.Tag             =   "info"
                  ImageIndex      =   11
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdRestart 
            Caption         =   "Restart"
            Default         =   -1  'True
            Height          =   285
            Left            =   3675
            TabIndex        =   3
            ToolTipText     =   "Restart, after freeze"
            Top             =   30
            Width           =   705
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   5595
            ScaleHeight     =   330
            ScaleWidth      =   1215
            TabIndex        =   7
            Top             =   45
            Width           =   1215
            Begin VB.Label lblCurPos 
               Alignment       =   2  'Center
               BackColor       =   &H80000009&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   8
               Top             =   0
               Width           =   1170
            End
         End
         Begin VB.PictureBox picHideCursor 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   5520
            ScaleHeight     =   375
            ScaleWidth      =   1335
            TabIndex        =   15
            Top             =   0
            Width           =   1335
         End
      End
   End
   Begin MSComDlg.CommonDialog cmnDlg1 
      Left            =   2490
      Top             =   1665
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   2505
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   2520
      Top             =   1665
   End
   Begin VB.Timer Timer3 
      Interval        =   25
      Left            =   3075
      Top             =   1770
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSaveItem 
         Caption         =   "Save Image As..."
      End
      Begin VB.Menu mnuOpenItem 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuOnTopItem 
         Caption         =   "On Top"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDontFreezeItem 
         Caption         =   "Disabable Freeze"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyItem 
         Caption         =   "Copy Zoomed Image"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuXItem 
         Caption         =   "Aim Visible                          Ctrl+A"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSepo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCursoronTopItem 
         Caption         =   "Cursor Tracker OnScreen"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom                                 Ctrl + (+/-)"
         Begin VB.Menu mnuXX100Item 
            Caption         =   "100%"
         End
         Begin VB.Menu mnuXX200Item 
            Caption         =   "200%"
         End
         Begin VB.Menu mnuXX400Item 
            Caption         =   "400%"
         End
         Begin VB.Menu mnuXX800Item 
            Caption         =   "800%"
         End
         Begin VB.Menu mnuXX1GItem 
            Caption         =   "1000%"
         End
         Begin VB.Menu mnuXX1500Item 
            Caption         =   "1500%"
         End
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBorderItem 
         Caption         =   "Hide Toolbar                      Ctrl + T"
      End
      Begin VB.Menu mnuNoHotKeyItem 
         Caption         =   "Disable HotKeys"
      End
      Begin VB.Menu mnuRestartItem 
         Caption         =   "Restart                             Ctrl + R"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSampleItem 
         Caption         =   "Sample Image"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowPixelColorItem 
         Caption         =   "Show Pixel Color              Ctrl + P"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComposerItem 
         Caption         =   "Color Composer               Ctrl + C"
      End
      Begin VB.Menu mnuSep011 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearSetting 
         Caption         =   "Clear All Settings"
         Begin VB.Menu mnuConfirm 
            Caption         =   "Confirm"
            Begin VB.Menu mnuClearRegItem 
               Caption         =   "Yes"
            End
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHotKeysItem 
         Caption         =   "HotKey Index"
      End
      Begin VB.Menu mnuContentsItem 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuAboutItem 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuH 
      Caption         =   "mnuH"
      Visible         =   0   'False
      Begin VB.Menu mnuHOnTopItem 
         Caption         =   "On Top"
      End
      Begin VB.Menu mnuSeph1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHFreezeItem 
         Caption         =   "Disable Freeze"
      End
      Begin VB.Menu mnuSepH2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHCopyItem 
         Caption         =   "Copy Zoomed Image"
      End
      Begin VB.Menu mnuSepH3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAimItem 
         Caption         =   "Aim Visible"
      End
      Begin VB.Menu mnuSepHo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhcursoronScreenItem 
         Caption         =   "Cursor Tracker OnScreen"
      End
      Begin VB.Menu mnuSepH4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHZoom 
         Caption         =   "Zoom                      +/-"
         Begin VB.Menu mnuH100 
            Caption         =   "100%"
         End
         Begin VB.Menu mnuH200 
            Caption         =   "200%"
         End
         Begin VB.Menu mnuH400 
            Caption         =   "400%"
         End
         Begin VB.Menu mnuH800 
            Caption         =   "800%"
         End
         Begin VB.Menu mnuH1000 
            Caption         =   "1000%"
         End
         Begin VB.Menu mnuH1500 
            Caption         =   "1500%"
         End
      End
      Begin VB.Menu mnuSepH5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHBorderItem 
         Caption         =   "Hide ToolBar"
      End
      Begin VB.Menu mnuSepH6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHRestartItem 
         Caption         =   "Restart"
      End
      Begin VB.Menu mnuSepH7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHPixelItem 
         Caption         =   "Show Pixel Color"
      End
      Begin VB.Menu mnuSepH8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHComposerItem 
         Caption         =   "Show Color Composer"
      End
      Begin VB.Menu mnuSepH9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHEndItem 
         Caption         =   "End"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''Set type for WinZoom
Private Type POINTAPI
    X As Long
    Y As Long
End Type

''''''''''''''''''''''Mouse-Scroll Detect : Currently not used'''''
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long

''''''''''''''''''''''HotKey Type
Private Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

''''''''''''''''''''''Declare API's for WinZoom
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long


'''''''''''''''''''''''Declare API to detect mouse click
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
''''''''''''''''''''''''Constatnts for mouse-click event
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
'''

'''
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
'''''''''''''''''''''''Declare variables for WinZoom
Private Const SRCCOPY As Long = &HCC0020
'''''''''''''''''''''''Constants for HotKeys
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const PM_REMOVE = &H1
Private Const WM_HOTKEY = &H312
'Private Type Msg
'    hwnd As Long
'    message As Long
'    wParam As Long
'    lParam As Long
'    time As Long
'    pt As POINTAPI
'End Type
'''''''''''''''''''''''APIs for HotKeys
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private bCancel As Boolean


Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11

'''''''''''''''''''''''Declare variables
Dim nZoomLevel As Long  ' 1 = 100%, 2 = 200%, 3 = 300% and so on...
Dim activetxt As Long
Dim isRunning As Boolean
Dim cnt As Integer
Dim freeze As Boolean
Dim intY As Integer
Dim intX As Integer
Dim errMSPaint As Integer
Dim cleanStart As Boolean
Dim flashMsg As Integer


Private Sub ShowZoom(lpPoint As POINTAPI)
    Dim hdc As Long
    Dim hwnd As Long
    Dim nTmp As Long
    
    Dim X As Long
    Dim Y As Long
    Dim nWidth As Long
    Dim nHeight As Long
    
    Dim xSrc As Long
    Dim ySrc As Long
    Dim nSrcWidth As Long
    Dim nSrcHeight As Long
    
    hwnd = GetDesktopWindow()
    hdc = GetDC(hwnd)
    
    With picZoom
        .Cls
        X = 0
        Y = 0
        nWidth = .ScaleWidth
        nHeight = .ScaleHeight
    End With
    
    xSrc = (nWidth / 2) / nZoomLevel
    xSrc = lpPoint.X - xSrc
    ySrc = (nHeight / 2) / nZoomLevel
    ySrc = lpPoint.Y - ySrc
    nSrcWidth = nWidth / nZoomLevel
    nSrcHeight = nHeight / nZoomLevel
    
    nTmp = nSrcWidth * nZoomLevel
    
    If (nTmp > nWidth) Then
        nWidth = nTmp
    ElseIf (nTmp < nWidth) Then
        nSrcWidth = nSrcWidth + 1
        nWidth = nTmp + nZoomLevel
    End If
    
    nTmp = nSrcHeight * nZoomLevel
    
    If (nTmp > nHeight) Then
        nHeight = nTmp
    ElseIf (nTmp < nHeight) Then
        nSrcHeight = nSrcHeight + 1
        nHeight = nTmp + nZoomLevel
    End If
    
    Call StretchBlt(picZoom.hdc, _
                    X, _
                    Y, _
                    nWidth, _
                    nHeight, _
                    hdc, _
                    xSrc, _
                    ySrc, _
                    nSrcWidth, _
                    nSrcHeight, _
                    SRCCOPY)
    
    Call ReleaseDC(hwnd, hdc)
    
End Sub

Public Sub cmdRestart_Click()

    Timer1.Enabled = True
    Timer1.Interval = 20
    cmdRestart.Enabled = False
    mnuRestartItem.Enabled = False
    picZoom.SetFocus
    BringWindowToTop Me.hwnd
    'SetActiveWindow Me.hwnd
    
End Sub


Private Sub Form_Load()

    If App.PrevInstance = True Then End
    
    cleanStart = False
    isComposerActive = False
    
    Me.Show

    activeWinAbout = -1
    activeWinImg = -1
    activeWinPixel = -1
    activeWinHelp = -1
    activeWinTip = -1
    activeWinComposer = -1
    startUp
    activetxt = GetActiveWindow
    
    cmdRestart.Enabled = False
    mnuRestartItem.Enabled = False
    isRunning = False
    cnt = 0
    flashMsg = 0
        
    isPixActive = False
    isPixColor = False
    
    Toolbar2.Buttons(11).Enabled = False
    errMSPaint = 0
    
    If infoLabel Then frmTip.Show vbModeless, Me
    infoLabel = False
    loadHotKeys
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    If isBorder = True Then
        picZoom.Height = Me.Height - 1100
        picZoom.Width = Me.Width - 300
        Toolbar1.Visible = True
    Else
        picZoom.Height = Me.Height - 425
        picZoom.Width = Me.Width
        Toolbar1.Visible = False
    End If
    
    picZoom.Top = 0
    picZoom.Left = Me.Width / 2 - (picZoom.Width / 2) - 60
    
    imgAim.Top = picZoom.ScaleHeight / 2 - (imgAim.Height / 2)
    imgAim.Left = picZoom.ScaleWidth / 2 - (imgAim.Width / 2) + 5
    
'    picZoom.Refresh
'    picZoom.Cls
'    Dim SX&, SY&, EX&, EY&
'
'    SX = (picZoom.ScaleWidth / 2)
'    SY = (picZoom.ScaleHeight / 2)
'    EX = (picZoom.ScaleWidth / 2)
'    EY = (picZoom.ScaleHeight / 2)
'
'
'    picZoom.Line (SX, SY - 10)-(EX, EY + 10), vbRed
'    picZoom.Line (SX - 10, SY)-(EX + 10, EY), vbRed
        
    If isAttach And isPixActive Then
        frmPixelColor.Left = Me.Left
        frmPixelColor.Top = Me.Top + Me.Height - frmPixelColor.Height
        frmPixelColor.SetFocus
    End If
    
    If isPixActive Then frmPixelColor.SetFocus: frmPixelColor.ZOrder 0: Me.ZOrder 1

End Sub

Private Sub Form_Unload(Cancel As Integer)

    '''''''''''''''''''''Start Unload HotKeys
    unregHKeys
    '''''''''''''''''''''End Unload HotKeys
    
    If isPixActive = True Then
        pixTop = frmPixelColor.Top
        pixLeft = frmPixelColor.Left
        Unload frmPixelColor
    End If
    
    intHeight = Me.Height
    intWidth = Me.Width
    intTop = Me.Top
    intLeft = Me.Left
    isTop = isOnTop
    
    If cleanStart = False Then
        setRegSet
    Else
        delRegSet
    End If

End Sub

Private Sub mnu100Item_Click()

    zoomIt (1)

End Sub

Private Sub mnu1500Item_Click()

    zoomIt (15)

End Sub

Private Sub mnu1GItem_Click()

    zoomIt (10)

End Sub

Private Sub mnu200Item_Click()

    zoomIt (2)

End Sub

Private Sub mnu400Item_Click()

    zoomIt (4)

End Sub

Private Sub mnu800Item_Click()

    zoomIt (8)

End Sub


Private Sub Label1_Click()

    Label1.Visible = False
    Picture3.Visible = False
    picZoom.Picture = LoadPicture("")
    infoLabel = False

End Sub

Public Sub lblCurPos_Click()

    picHideCursor.ZOrder
    lblCursorPos.Visible = True
    mnuCursoronTopItem.Checked = True

End Sub

Private Sub mnuCursoronTopItem_Click()

    If lblCursorPos.Visible = True Then
        mnuCursoronTopItem.Checked = False
        picHideCursor_click
    Else
        mnuCursoronTopItem.Checked = True
        lblCurPos_Click
    End If
    
    mnuhcursoronScreenItem.Checked = mnuCursoronTopItem.Checked

End Sub

Private Sub mnuhcursoronScreenItem_Click()

    mnuCursoronTopItem_Click

End Sub

Private Sub picHideCursor_click()

    picHideCursor.ZOrder 1
    lblCursorPos.Visible = False
    mnuCursoronTopItem.Checked = False

End Sub

Private Sub mnuAboutItem_Click()

    If mnuAboutItem.Checked = True Then
        Unload frmAbout
        Exit Sub
    End If
    
    frmAbout.Show vbModeless, Me

End Sub

Private Sub mnuBorderItem_Click()

    mnuBorderItem.Checked = Not mnuBorderItem.Checked
    If mnuBorderItem.Checked = True Then
        picZoom.BorderStyle = 0
        isBorder = False
        mnuFile.Visible = False
        mnuTools.Visible = False
        mnuHelp.Visible = False
        picZoom.Top = 1
        picZoom.Left = 1
        picZoom.Width = Me.Width
    Else
        picZoom.BorderStyle = 1
        isBorder = True
        mnuFile.Visible = True
        mnuTools.Visible = True
        mnuHelp.Visible = True
    End If
    
    Form_Resize
    Me.Height = Me.Height + 10
    Form_Resize
    Me.Height = Me.Height - 10
    
    mnuh_click

End Sub

Private Sub mnuClearRegItem_Click()

    lblFlashMSG.Tag = "f"
    flashMessageMain "Reset Registry", 2
    delRegSet
    startUp
    cleanStart = True

End Sub

Private Sub mnuComposerItem_Click()

    If isComposerActive = False Then
        frmComposer.Show vbModeless, Me
        mnuComposerItem.Checked = True
    Else
        Unload frmComposer
        mnuComposerItem.Checked = False
    End If
    
    mnuh_click

End Sub

Private Sub mnuContentsItem_Click()

    If mnuContentsItem.Checked = True Then
        Unload frmHelp
        Exit Sub
    End If
    
    frmHelp.Show vbModeless, Me

End Sub

Private Sub mnuCopyItem_Click()

    copyImage

End Sub

Private Sub mnuDontFreezeItem_Click()

    freeze = Not freeze
    isNoFreeze = freeze
    mnuDontFreezeItem.Checked = Not freeze
    If freeze = False Then cmdRestart_Click
    
    mnuh_click

End Sub

Private Sub mnuExitItem_Click()

    Unload Me
    End

End Sub


Private Sub mnuH100_Click()

    mnuXX100Item_Click

End Sub

Private Sub mnuH1000_Click()

    mnuXX1GItem_Click

End Sub

Private Sub mnuH1500_Click()

    mnuXX1500Item_Click

End Sub

Private Sub mnuH200_Click()

    mnuXX200Item_Click

End Sub

Private Sub mnuH400_Click()

    mnuXX400Item_Click

End Sub

Private Sub mnuH800_Click()

    mnuXX800Item_Click

End Sub


Private Sub mnuHAimItem_Click()

    mnuXItem_Click

End Sub

Private Sub mnuHBorderItem_Click()

    mnuBorderItem_Click

End Sub

Private Sub mnuHComposerItem_Click()

    mnuComposerItem_Click

End Sub

Private Sub mnuHCopyItem_Click()

    mnuCopyItem_Click

End Sub

Private Sub mnuHelpHotKeysItem_Click()

    Dim R As String
    Dim isActve As Boolean
    
    isActve = Timer1.Enabled
    
    R = vbCrLf

    MsgBox "Control + (+) : Zoom + 1" & R & "Control 6 (-) : Zoom - 1" & R & "Control + A : Toggle Aim" & R & "Control + C : Toggle Composer" & R & "Control + T : Toggle Toolbar" & R & "Control + P : Toggle Pixel Color Window" & R & "Control + Escape : Quit", vbInformation, "HotKey Overview"
    
    If isActve Then cmdRestart_Click

End Sub

Private Sub mnuHEndItem_Click()

    mnuExitItem_Click

End Sub

Private Sub mnuHFreezeItem_Click()

    mnuDontFreezeItem_Click

End Sub

Private Sub mnuHOnTopItem_Click()

    mnuOnTopItem_Click

End Sub

Private Sub mnuHPixelItem_Click()

    mnuShowPixelColorItem_Click

End Sub

Private Sub mnuHRestartItem_Click()

    mnuRestartItem_Click

End Sub


Private Sub mnuNoHotKeyItem_Click()

    If LCase(mnuNoHotKeyItem.Caption) = "disable hotkeys" Then
        mnuNoHotKeyItem.Caption = "Enable HotKeys"
        useHKeys = False
        unregHKeys
    Else
        mnuNoHotKeyItem.Caption = "Disable HotKeys"
        useHKeys = True
        loadHotKeys
    End If

End Sub

Private Sub mnuOnTopItem_Click()

    isOnTop = Not isOnTop
    mnuOnTopItem.Checked = isOnTop
    
    If isPixActive And isOnTop Then
        setTopMost2 frmPixelColor
    End If
    
    mnuHOnTopItem.Checked = mnuOnTopItem.Checked
    
    If mnuOnTopItem.Checked = True Then
        flashMessageMain "Always On Top", 1
    Else
        flashMessageMain "Always On Top Disabled", 1
    End If
    
    setTopMost
    mnuh_click

End Sub

Private Sub mnuOpenItem_Click()

'    If isOnTop = True Then
'        MsgBox ("The program will disable 'On Top' feature to view the image. You can turn it back on when done")
'        isOnTop = False
'        setTopMost
'    End If
    
    cmnDlg1.Filter = "Bitmap Images( *.bmp)|*.bmp|JPEG Images ( *.jpg)|*.jpg|All Images (*.*)|*.*"
    cmnDlg1.CancelError = True
    On Error GoTo hell
    cmnDlg1.ShowOpen
    If cmnDlg1.FileName <> "" Then
        Load frmLoadImage
        frmLoadImage.Image1.Picture = LoadPicture(cmnDlg1.FileName)
        frmLoadImage.Show vbModeless, Me
    Else
        Unload frmLoadImage
        Exit Sub
    End If
hell:
    If Err.Number = 32755 Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox ("Invalid File"), , "Error"
    End If
    
End Sub

Private Sub mnuRestartItem_Click()

    cmdRestart_Click
    
    mnuh_click

End Sub

Private Sub mnuSampleItem_Click()

'    If isOnTop = True Then
'        MsgBox ("The program will disable 'On Top' feature to view the sample image. You can turn it back on when done")
'        isOnTop = False
'        setTopMost
'    End If

    If mnuSampleItem.Checked = True Then
        Unload frmLoadImage
        Exit Sub
    End If
    
    Load frmLoadImage
    frmLoadImage.Show vbModeless, Me
    'Me.SetFocus
    cmdRestart_Click

End Sub

Private Sub mnuSaveItem_Click()
    
    On Error GoTo erHand
    cmnDlg1.Filter = "Bitmap Image (*.bmp) |*.bmp|JPEG Image (*.jpg )|*.jpg|All Files|*.*"
    cmnDlg1.CancelError = True
    cmnDlg1.ShowSave
    SavePicture picZoom.Image, cmnDlg1.FileName
    flashMessageMain cmnDlg1.FileName & " Saved...", 1
    
    Exit Sub

erHand:
    Exit Sub

End Sub

Private Sub mnuShowPixelColorItem_Click()

    mnuShowPixelColorItem.Checked = Not mnuShowPixelColorItem.Checked
    isPixColor = mnuShowPixelColorItem.Checked
    If isPixColor Then
        frmPixelColor.Show vbModeless, Me
        setTopMost2 frmPixelColor
        Toolbar2.Buttons(10).Enabled = False
        Toolbar2.Buttons(11).Enabled = True
    Else
        Unload frmPixelColor
        Me.Caption = "WinZoom"
        Toolbar2.Buttons(10).Enabled = True
        Toolbar2.Buttons(11).Enabled = False
    End If
    
    mnuh_click

End Sub

Private Sub mnuXItem_Click()

    mnuXItem.Checked = Not mnuXItem.Checked
    imgAim.Visible = mnuXItem.Checked
    isAim = mnuXItem.Checked
    
    mnuh_click

End Sub

Private Sub mnuXX100Item_Click()

    zoomIt (1)
    
    mnuh_click

End Sub

Private Sub mnuXX1500Item_Click()

    zoomIt (15)
    
    mnuh_click

End Sub

Private Sub mnuXX1GItem_Click()

    zoomIt (10)
    
    mnuh_click

End Sub

Private Sub mnuXX200Item_Click()

    zoomIt (2)
    
    mnuh_click

End Sub

Private Sub mnuXX400Item_Click()

    zoomIt (4)
    
    mnuh_click

End Sub

Private Sub mnuXX800Item_Click()

    zoomIt (8)
    
    mnuh_click

End Sub


Private Sub Picture3_Click()

    Label1_Click

End Sub

Private Sub picZoom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuH
    End If

End Sub

Private Sub Timer1_Timer()
    
    Static lpPoint As POINTAPI
    
    activewindow = GetActiveWindow
    If freeze = True Then
        If activewindow <> activetxt And activeWinComposer <> activewindow And activeWinTip <> activewindow And activeWinHelp <> activewindow And activeWinAbout <> activewindow And activeWinPixel <> activewindow And activeWinImg <> activewindow Then
            Timer1.Interval = 0
            Timer1.Enabled = False
            cmdRestart.Enabled = True
            mnuRestartItem.Enabled = True
        End If
    End If
    If (GetCursorPos(lpPoint)) Then
        Call ShowZoom(lpPoint) ' Show zoom from cursor position
    End If
 
    If isPixColor = False Then Me.Caption = "WinZoom": Exit Sub
    
    Dim tPOS As POINTAPI
    Dim sTmp As String
    Dim lColor As Long
    Dim lDC As Long
    Dim theHex

    lDC = GetWindowDC(0)
    Call GetCursorPos(tPOS)
    lColor = GetPixel(lDC, tPOS.X, tPOS.Y)
'    Label2.BackColor = lColor
    'Caption = "RGB : " & lColor & "        "
    sTmp = Right$("000000" & Hex(lColor), 6)
    theHex = Right$(sTmp, 2) & Mid$(sTmp, 3, 2) & Left$(sTmp, 2)
    Caption = "HEX:" & theHex
    copyHEX = theHex
    frmPixelColor.Frame2.Caption = "HEX : " & theHex
    'mnuHEX.Caption = "HEX : " & theHex
    
    R = Val("&h" & Right(sTmp, 2))
    G = Val("&h" & Mid(sTmp, 3, 2))
    B = Val("&h" & Left(sTmp, 2))
    
    copyRGB = R & ", " & G & ", " & B
    
    Caption = Caption & "        RGB : " & R & " " & G & " " & B
    
    frmPixelColor.Frame1.Caption = R & " * " & G & " * " & B
    
    SetHex R, G, B

    'mnuRGB.Caption = "RGB : " & r & " " & g & " " & b

End Sub


Sub SetHex(RVal, GVal, BVal)
    
    Dim vbStr As String
    
    RHex = Hex(RVal)
    If Len(CStr(RHex)) < 2 Then RHex = "0" & RHex
    GHex = Hex(GVal)
    If Len(CStr(GHex)) < 2 Then GHex = "0" & GHex
    BHex = Hex(BVal)
    If Len(CStr(BHex)) < 2 Then BHex = "0" & BHex
    vbStr = "&H" & BHex & GHex & RHex & "&"
    Caption = Caption & "        VBHEX : " & vbStr
    
    'mnuVBHEX.Caption = "VBHEX : " & vbstr
    copyVBHEX = vbStr
    
    frmPixelColor.Frame3.Caption = "VB : " & vbStr
    frmPixelColor.lblColor.BackColor = "&H" & BHex & GHex & RHex
    
End Sub


Private Sub zoomIt(X)

    If X <= 1 Then X = 1
    If X >= 100 Then X = 100
    zoomAt = X
    nZoomLevel = X
    lblzoom.Caption = X & "00%"
    
    uncheckMenu
    
    If lblFlashMSG.Tag <> "f" Then
        flashMessageMain lblzoom.Caption, 1
    End If
    lblFlashMSG.Tag = "t"
    
    Select Case X
    Case 1: mnuXX100Item.Checked = True
    Case 2: mnuXX200Item.Checked = True
    Case 4: mnuXX400Item.Checked = True
    Case 8: mnuXX800Item.Checked = True
    Case 10: mnuXX1GItem.Checked = True
    Case 15: mnuXX1500Item.Checked = True
    Case Else: uncheckMenu
    End Select

End Sub


Private Sub Timer2_Timer()
    
    Dim lpPoint As POINTAPI
    
    If (GetCursorPos(lpPoint)) Then
        lblCurPos.Caption = "X: " & lpPoint.X & " Y:" & lpPoint.Y
        lblCursorPos.Caption = lblCurPos.Caption
    End If
    
    If intX <> Me.Left Or intY <> Me.Top Then
        GoSub setPosition
    End If
    
    If flashMsg > 0 Then
        'lblFlashMSG.Visible = True
        txtFlashMsg.Visible = True
        flashMsg = flashMsg - 1
    Else
        'lblFlashMSG.Visible = False
        txtFlashMsg.Visible = False
    End If
    
'    If isPixActive Then
'        If frmPixelColor.Top <> pixTop Or frmPixelColor.Left <> pixLeft Then
'            GoTo setPosition
'        End If
'    End If
Exit Sub
setPosition:

    If isPixActive And isAttach Then
        On Error Resume Next
        frmPixelColor.Left = Me.Left
        frmPixelColor.Top = Me.Top + Me.Height - frmPixelColor.Height
        intX = Me.Left
        intY = Me.Top
    End If
    Return
        
End Sub


Private Sub Timer3_Timer()
    
    If useHKeys = False Then Exit Sub
    
    If GetKeyState(VK_CONTROL) < 0 Then
        If GetKeyState(vbKeyAdd) < 0 Then
            nZoomLevel = nZoomLevel + 1
            zoomIt (nZoomLevel)
            Timer3.Enabled = False
            setPause 0.15, True
        ElseIf GetKeyState(vbKeySubtract) < 0 Then
            nZoomLevel = nZoomLevel - 1
            zoomIt (nZoomLevel)
            Timer3.Enabled = False
            setPause 0.15, True
        ElseIf GetKeyState(vbKeyR) < 0 Then
            cmdRestart_Click
            Timer3.Enabled = False
            setPause 0.3, True
        ElseIf GetKeyState(vbKeyP) < 0 Then
            mnuShowPixelColorItem_Click
            Timer3.Enabled = False
            setPause 0.3, True
        ElseIf GetKeyState(vbKeyC) < 0 Then
            mnuComposerItem_Click
            Timer3.Enabled = False
            setPause 0.3, True
        ElseIf GetKeyState(vbKeyT) < 0 Then
            mnuBorderItem_Click
            Timer3.Enabled = False
            setPause 0.3, True
        ElseIf GetKeyState(vbKeyA) < 0 Then
            mnuXItem_Click
            Timer3.Enabled = False
            setPause 0.3, True
        ElseIf GetKeyState(vbKeyH) < 0 Then
            mnuContentsItem_Click
            Timer3.Enabled = False
            setPause 0.3, True
        ElseIf GetKeyState(vbKeyEscape) < 0 Then
            mnuExitItem_Click
            Timer3.Enabled = False
            setPause 0.3, True
        End If
        
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Tag
    Case "ZoomIn"
        nZoomLevel = nZoomLevel + 1
        zoomIt (nZoomLevel)
    End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Tag
    Case "ZoomOut"
        nZoomLevel = nZoomLevel - 1
        zoomIt (nZoomLevel)
    Case "save"
        Timer1.Enabled = False
        cmdRestart.Enabled = True
        mnuRestartItem.Enabled = True
        mnuSaveItem_Click
    Case "open"
        mnuOpenItem_Click
        Me.SetFocus
        cmdRestart_Click
    Case "copy"
        copyImage
    Case "QZoom"
        nZoomLevel = nZoomLevel + 5
        zoomIt (nZoomLevel)
    Case "MSPaint"
                
        Dim winPath As String
        
        winPath = Space$(255)

        Call GetSystemDirectory(winPath, Len(winPath))
        
        List1.Clear
        List1.AddItem winPath
        winPath = List1.List(0)
        List1.Clear
                
        winPath = Trim(winPath)
        
        File1.Path = winPath
        Dim srch As Integer
        For srch = 0 To File1.ListCount - 1
            If UCase(File1.List(srch)) = UCase("mspaint.exe") Then
                GoTo location1
                Exit For
                Exit Sub
            End If
        Next srch
        
        GoTo location2
        
    Case "pixel"
        mnuShowPixelColorItem_Click
    Case "pixelClose"
        mnuShowPixelColorItem_Click
    Case "openComposer"
        mnuComposerItem_Click
    End Select
    
    Exit Sub

location1:
    
    On Error GoTo erHand2
    Shell winPath & "\mspaint.exe", vbNormalFocus
    Exit Sub

location2:
        
    On Error GoTo erHand
    
    winPath = ""
    winPath = Space$(255)

    Call GetWindowsDirectory(winPath, Len(winPath))
        
    List1.Clear
    List1.AddItem winPath
    winPath = List1.List(0)
    List1.Clear
                
    winPath = Trim(winPath)
    Shell winPath & "\mspaint.exe", vbNormalFocus
    
    Exit Sub
    
erHand2:
    If Err.Number = 53 Then
        MsgBox Err.Description, vbInformation, "WinZoom Error # " & Err.Number
    Else
        MsgBox "Error while loading MSPaint.exe, the file might be corrupted." & vbCrLf & _
                Err.Description & " Error Number " & Err.Number, vbInformation, "WinZoom"
    End If
    
    Exit Sub
    
erHand:
    
    MsgBox ("Unable to locate Microsoft Paint." & vbCrLf & _
            "Please, make sure that the file mspaint.exe exists in your system directory." & vbCrLf & _
            "You will need to restart the program to use this option."), vbInformation, "WinZoom"
    Toolbar2.Buttons(8).Enabled = False
    Exit Sub

End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Tag
    Case "close"
        Unload Me
        End
    Case "info"
        mnuAboutItem_Click
    Case "closeImage"
        Unload frmLoadImage
        Toolbar3.Buttons(3).Enabled = False
    End Select

End Sub

Sub copyImage()

    Clipboard.Clear
    Clipboard.SetData picZoom.Image

End Sub

Sub startUp()

    getRegSet
    
    Me.Top = intTop
    Me.Left = intLeft
    Me.Width = intWidth
    Me.Height = intHeight
    
    Label1.Visible = False    'infoLabel - Not used
    Picture3.Visible = False  'infoLabel - Not used
        
    If useHKeys = False Then
        mnuNoHotKeyItem.Caption = "Enable HotKeys"
    Else
        mnuNoHotKeyItem.Caption = "Disable HotKeys"
    End If
    
    intX = intLeft
    intY = intTop
    
    isOnTop = isTop
    
    setTopMost
    mnuOnTopItem.Checked = isOnTop
    
    
    zoomIt (zoomAt)
    freeze = isNoFreeze
    mnuDontFreezeItem.Checked = Not freeze
    
    If isAim = False Then
        mnuXItem_Click
    End If
    
    If isBorder = True Then
        picZoom.BorderStyle = 1
        mnuFile.Visible = True
        mnuTools.Visible = True
        mnuHelp.Visible = True
    Else
        picZoom.BorderStyle = 0
        mnuFile.Visible = False
        mnuTools.Visible = False
        mnuHelp.Visible = False
        picZoom.Top = 1
        picZoom.Left = 1
        picZoom.Width = Me.Width
    End If
    mnuBorderItem.Checked = Not isBorder
    
    If isPixColor = True Then
        mnuShowPixelColorItem_Click
        isPixColor = True
    Else
        '''''''''''''''''
    End If
    
End Sub

Sub uncheckMenu()

'    Dim ctrl As Control
'
'    On Error Resume Next
'
'    For Each ctrl In Controls
'        If Left(ctrl.Name, 5) = LCase("mnuxx") Then
'            ctrl.Checked = False
'        End If
'    Next ctrl

    mnuXX100Item.Checked = False
    mnuXX200Item.Checked = False
    mnuXX400Item.Checked = False
    mnuXX800Item.Checked = False
    mnuXX1GItem.Checked = False
    mnuXX1500Item.Checked = False

End Sub

Private Sub unregHKeys()

    '''''''''''''''''''''Start Unload HotKeys
    bCancel = True
    Call UnregisterHotKey(Me.hwnd, &HBFFF&)
    '''''''''''''''''''''End Unload HotKeys

End Sub

Private Sub proccessMsgs()

    ''''''''''''''''''''''''''Start HotKeys
    If useHKeys And bCancel = False Then
        Dim message As Msg
        Do While Not bCancel
            WaitMessage
            'check if it's a HOTKEY-message
            If PeekMessage(message, Me.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
                Timer1.Enabled = False: cmdRestart.Enabled = True: mnuRestartItem.Enabled = True
            End If
            DoEvents
        Loop
    End If
    ''''''''''''''''''''''''''End HotKeys

End Sub

Private Sub loadHotKeys()

    '''''''''''''''''''''Start Unload HotKeys
    Dim Ret As Long
    bCancel = False
    'register the Ctrl-M hotkey
    Ret = RegisterHotKey(Me.hwnd, &HBFFF&, MOD_CONTROL, vbKeyF)
    proccessMsgs
    '''''''''''''''''''''End Unload HotKeys

End Sub

Private Sub mnuh_click()

    mnuHOnTopItem.Checked = mnuOnTopItem.Checked
    
    mnuHFreezeItem.Checked = mnuDontFreezeItem.Checked
    
    mnuHAimItem.Checked = mnuXItem.Checked
    
    mnuHBorderItem.Checked = mnuBorderItem.Checked
    
    mnuHRestartItem.Enabled = mnuRestartItem.Enabled
    
    mnuHPixelItem.Checked = mnuShowPixelColorItem.Checked
    
    mnuHComposerItem.Checked = mnuComposerItem.Checked
    
    mnuhcursoronScreenItem.Checked = mnuCursoronTopItem.Checked
        
    mnuH100.Checked = mnuXX100Item.Checked
    mnuH200.Checked = mnuXX200Item.Checked
    mnuH400.Checked = mnuXX400Item.Checked
    mnuH800.Checked = mnuXX800Item.Checked
    mnuH100.Checked = mnuXX1GItem.Checked
    mnuH1500.Checked = mnuXX1500Item.Checked

End Sub

Private Sub setPause(X, Optional resetTimer As Boolean)

    Dim start
   
    start = Timer
    Do While Timer < start + X
       DoEvents
    Loop
    
    If resetTimer Then Timer3.Enabled = True
   
End Sub

Public Sub flashMessageMain(msgTxt As String, ByRef msgPause As Integer)

    Dim tmpValue
    
    lblFlashMSG.Caption = msgTxt
    txtFlashMsg.Text = msgTxt
    picZoom.Refresh
    'lblFlashMSG.Top = (picZoom.ScaleHeight / 2) - (lblFlashMSG.Height / 2)
    'lblFlashMSG.Left = (picZoom.ScaleWidth / 2) - (lblFlashMSG.Width / 2)
    '''''TextMSG
    txtFlashMsg.Width = Len(msgTxt) * 10 + 20
    'txtFlashMsg.Height = 10
    txtFlashMsg.Top = (picZoom.ScaleHeight) - (txtFlashMsg.Height)
    txtFlashMsg.Left = (picZoom.ScaleWidth / 2) - (txtFlashMsg.Width / 2)
    '''''End TextMSG
    If Timer2.Interval > 0 Then
        tmpValue = 1000 / Timer2.Interval
    Else
        Exit Sub
    End If
    flashMsg = CInt(msgPause + 0.5) * (tmpValue / 2.5)
    'Do While Timer < start + msgPause
    '    DoEvents
    '    lblFlashMSG.Visible = True
    'Loop
    'lblFlashMSG.Visible = False

End Sub


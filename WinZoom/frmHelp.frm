VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   6480
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7020
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12383
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   -2147483638
      TabCaption(0)   =   "Main Screen"
      TabPicture(0)   =   "frmHelp.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picMain"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdToggleMain"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "ToolBar Buttons"
      TabPicture(1)   =   "frmHelp.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image2"
      Tab(1).Control(1)=   "cmdToggleTBar"
      Tab(1).Control(2)=   "picTBar"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Color Picker"
      TabPicture(2)   =   "frmHelp.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdtogglePixel"
      Tab(2).Control(1)=   "picPixel"
      Tab(2).Control(2)=   "Image3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "HotKey Index"
      TabPicture(3)   =   "frmHelp.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtHotKeys"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.TextBox txtHotKeys 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -74895
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   990
         Width           =   8190
      End
      Begin VB.CommandButton cmdtogglePixel 
         Caption         =   "Screenshot"
         Height          =   375
         Left            =   -71280
         TabIndex        =   9
         Tag             =   "text"
         Top             =   6480
         Width           =   975
      End
      Begin VB.PictureBox picPixel 
         BorderStyle     =   0  'None
         Height          =   5805
         Left            =   -74955
         ScaleHeight     =   5805
         ScaleWidth      =   8325
         TabIndex        =   8
         Top             =   660
         Width           =   8325
         Begin VB.TextBox txtPixel 
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5715
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   0
            Width           =   8085
         End
      End
      Begin VB.PictureBox picTBar 
         BorderStyle     =   0  'None
         Height          =   5805
         Left            =   -74970
         ScaleHeight     =   5805
         ScaleWidth      =   8340
         TabIndex        =   6
         Top             =   645
         Width           =   8340
         Begin VB.TextBox txtTBar 
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5595
            Left            =   135
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   0
            Width           =   8085
         End
      End
      Begin VB.CommandButton cmdToggleTBar 
         Caption         =   "Screenshot"
         Height          =   375
         Left            =   -71280
         TabIndex        =   5
         Tag             =   "text"
         Top             =   6480
         Width           =   975
      End
      Begin VB.CommandButton cmdToggleMain 
         Caption         =   "Screenshot"
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Tag             =   "text"
         Top             =   6480
         Width           =   975
      End
      Begin VB.PictureBox picMain 
         BorderStyle     =   0  'None
         Height          =   5805
         Left            =   60
         ScaleHeight     =   5805
         ScaleWidth      =   8325
         TabIndex        =   1
         Top             =   660
         Width           =   8325
         Begin VB.TextBox txtMain 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5715
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   0
            Width           =   8085
         End
      End
      Begin VB.Image Image3 
         Height          =   5715
         Left            =   -74910
         Picture         =   "frmHelp.frx":037A
         Stretch         =   -1  'True
         Top             =   705
         Width           =   8220
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   5715
         Left            =   -74925
         Picture         =   "frmHelp.frx":160F2
         Stretch         =   -1  'True
         Top             =   705
         Width           =   8220
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   5715
         Left            =   90
         Picture         =   "frmHelp.frx":48660
         Stretch         =   -1  'True
         Top             =   705
         Width           =   8220
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdtogglePixel_Click()

    If cmdtogglePixel.Tag = "text" Then
        cmdtogglePixel.Tag = "img"
        cmdtogglePixel.Caption = "Text"
        picPixel.Visible = False
    Else
        cmdtogglePixel.Tag = "text"
        cmdtogglePixel.Caption = "Screenshot"
        picPixel.Visible = True
    End If

End Sub

Private Sub cmdToggleMain_Click()

    If cmdToggleMain.Tag = "text" Then
        cmdToggleMain.Tag = "img"
        cmdToggleMain.Caption = "Text"
        picMain.Visible = False
    Else
        cmdToggleMain.Tag = "text"
        cmdToggleMain.Caption = "Screenshot"
        picMain.Visible = True
    End If

End Sub

Private Sub cmdToggleTBar_Click()

    If cmdToggleTBar.Tag = "text" Then
        cmdToggleTBar.Tag = "img"
        cmdToggleTBar.Caption = "Text"
        picTBar.Visible = False
    Else
        cmdToggleTBar.Tag = "text"
        cmdToggleTBar.Caption = "Screenshot"
        picTBar.Visible = True
    End If

End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()

    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    
    SSTab1.ShowFocusRect = False
    SSTab1.Tab = 0
    
    addHelp
    
    Show
    activeWinHelp = GetActiveWindow
    setTopMost2 Me
    frmMain.mnuContentsItem.Checked = True

End Sub

Sub addHelp()

    Dim R As String
    
    R = vbCrLf

    txtMain.Text = "1.  When using WinZoom for the first time, a Tip-Of-The-Day" & R & _
    "     will apper. Itgives you a quick reference to using HotKeys. " & R & _
    "     You can aslo click on Menu->Help->HotKey Help to get a list" & R & _
    "     of available HotKeys and their function." & R & R & _
    "2.  When starting up your WinZoom the first time, resize the " & R & _
    "     main window to your comfort. After restarting WinZoom," & R & _
    "     you will see that it returns to the same size and position as it was the last" & R & _
    "     time you used it. Just like the window size and screen position is saved to " & R & _
    "     your computer's memory, all options that you change will also be saved." & R & _
    "     That memory can be deleted of all saved settings by going into " & R & _
    "     Tools -> Clear All Settings -> Confirm -> Yes, in WinZoom's Menu. " & R & _
    "     You will instantly see that the window resets to default size, " & R & _
    "     and all settings are returned to default as well."
    
    '/*/------------------------------------------------------\*\'
    
    txtTBar.Text = "1.  Use Zoom In and Zoom Out buttons to zoom +/- 100%." & R & _
    "     Zoom-Level readout box displays the current level of zoom applied." & R & R & _
    "2.  To save the image that you are currently seeing on the screen, you" & R & _
    "     can use the 'Save Image' button. Please note that the image will be" & R & _
    "     saved as you see it, magnified. When saving, you can choose from " & R & _
    "     either a bitmap (.BMP) or JPEG (.JPG) format." & R & R & _
    "3.   You can open a separate window with an image " & R & _
    "     of your choice. To do that, you can either choose Menu->File->Open... " & R & _
    "     or simply click 'Open Image' button in the toolbar. A sample image " & R & _
    "     is provided, to view, go to Menu->Tools->Sample Image. Note, that " & R & _
    "     the image its self is not magnified, but you can zoom-in on it, by " & R & _
    "     placing the cursor over the desired part. When an image is open, " & R & _
    "     you will notice that 'X' button on the toolbar will become enabled, " & R & _
    "     and will remain enabled until the image window is closed. By " & R & _
    "     clicking the 'X' button, image window will close. (Described " & R & _
    "     as Close Image-Preview in the screen shot). The point of " & R & _
    "     being able to open an image is to be able to view details." & R & R & _
    "4.   If you decide to copy, what currently appears in the zoom " & R & _
    "     window, you can do that by click 'Copy Image' button or by " & R & _
    "     going to Menu->Tools->Copy Zoomed Image. That will place your" & R & _
    "     current workspace onto the clipboard. You can later use 'Paste'" & R & _
    "     method to paste that image to any software that accepts images" & R & _
    "     such as MSPaint or Adobe Photo Shop and even MSWord." & R & R & _
    "5.   You can use 'Quick Zoom' button to zoom in by 500%." & R & _
    "      You can also go to Menu->Tools->Zoom, and choose the " & R & "      level of magnification." & R & R
    
    txtTBar.Text = txtTBar.Text & _
    "6.    If you have MSPaint installed on your system, you can " & R & _
    "      access by simply clicking 'MSPaint' button on the toolbar." & R & R & _
    "7.    To open the Pixel Color Toolbar, you can click 'Open Pixel-Color'" & R & _
    "       button, and to close it, click the button next to it 'Close Pixel-Color" & R & R & _
    "8.    When you click on any window outside WinZoom" & R & _
    "      (If Freeze is Enabled), the WinZoom freezes the image. To restart," & R & _
    "      click on 'Restart' button, or goto Menu->Tools->Restart" & R & R & _
    "9.    To close WinZoom application, you can click the circled X" & R & _
    "       on the toolbar, or go to Menu->File->Exit." & R & R & _
    "10.   To display an 'About Box', click on 'i' button." & R & R & _
    "11.   The current cursor position is displayed in " & R & _
    "        right-bottom corner of the toolbar."
    
    '/*/------------------------------------------------------\*\'
    
    txtPixel.Text = "The Pixel-Color displays the current color under the cursor. It can display RGB-Color, HEX or VBHex, or any combination of the three."
    
    
    txtHotKeys.Text = "Note tht when using HotKeys for anothe applicaton, for example - 'Control - F' in MSWord to find text, HotKeys in WinZoom will react to itas Freeze Image HotKey. Therefore you have an optin to disable using HotKeys by going to Menu->Tools->Disable HotKeys." & R & R & _
                      "To use HotKeys:" & R & "Control + (+) : Zoom + 1" & R & "Control 6 (-) : Zoom - 1" & R & "Control + A : Toggle Aim" & R & "Control + C : Toggle Composer" & R & "Control + F : Freeze the image" & R & "Control + T : Toggle Toolbar" & R & "Control + P : Toggle Pixel Color Window" & R & "Control + Escape : Quit"
    

End Sub

Private Sub Form_Unload(Cancel As Integer)

    activeWinHelp = -1
    frmMain.mnuContentsItem.Checked = False

End Sub


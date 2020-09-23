VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMAXmain 
   BackColor       =   &H80000000&
   Caption         =   "MP3 MAX"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15240
   Icon            =   "frmMAXmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6840
      TabIndex        =   40
      Text            =   "Combo2"
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H80000000&
      Caption         =   "MP3 Info"
      Height          =   1095
      Left            =   6600
      TabIndex        =   36
      Top             =   6840
      Width           =   1935
      Begin VB.Label labFreqChan 
         BackColor       =   &H80000000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label labBitRate 
         BackColor       =   &H80000000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label labLayer 
         BackColor       =   &H80000000&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9000
      TabIndex        =   34
      Text            =   "Combo1"
      Top             =   7200
      Width           =   615
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   8
      Left            =   4920
      Top             =   6960
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   8
      Left            =   4440
      Top             =   6960
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   8
      Left            =   4440
      Top             =   6120
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   8
      Left            =   4440
      Top             =   5520
   End
   Begin VB.VScrollBar Vol2 
      Height          =   2415
      Left            =   4080
      Max             =   -4000
      TabIndex        =   33
      Top             =   6480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.VScrollBar Vol1 
      Height          =   2415
      Left            =   3720
      Max             =   -4000
      TabIndex        =   32
      Top             =   6480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   5640
      Top             =   5520
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5715
      TabIndex        =   31
      Top             =   1680
      Width           =   5775
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   6915
      TabIndex        =   30
      Top             =   480
      Width           =   6975
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   6120
   End
   Begin VB.CommandButton cmdListDn 
      Caption         =   "DN"
      Height          =   375
      Left            =   9360
      TabIndex        =   24
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdListUp 
      Caption         =   "UP"
      Height          =   375
      Left            =   9360
      TabIndex        =   23
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdSnd5 
      Caption         =   "5"
      Height          =   375
      Left            =   9480
      TabIndex        =   21
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdSnd4 
      Caption         =   "4"
      Height          =   375
      Left            =   9000
      TabIndex        =   20
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdSnd3 
      Caption         =   "3"
      Height          =   375
      Left            =   8520
      TabIndex        =   19
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdSnd2 
      Caption         =   "2"
      Height          =   375
      Left            =   8040
      TabIndex        =   18
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdSnd1 
      Caption         =   "1"
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      Top             =   720
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   8280
      TabIndex        =   13
      Top             =   8880
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   405
      Left            =   8265
      TabIndex        =   12
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   6720
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   7320
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8415
      Left            =   10320
      TabIndex        =   9
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   14843
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FileName"
         Object.Width           =   8213
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FullPath"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   7920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   9360
      Width           =   15255
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5640
      Top             =   8520
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   10035
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   741
      ButtonWidth     =   2355
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop           "
            Key             =   "Stop"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Previous       "
            Key             =   "Prev"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Play           "
            Key             =   "Play"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next          "
            Key             =   "Next"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pause         "
            Key             =   "Pause"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   10455
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   750
            MinWidth        =   750
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5080
            MinWidth        =   5080
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "frmMAXmain.frx":0442
      Left            =   480
      List            =   "frmMAXmain.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   8520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "frmMAXmain.frx":0446
      Left            =   480
      List            =   "frmMAXmain.frx":0448
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   8760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMAXmain.frx":044A
            Key             =   "cldfolder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMAXmain.frx":05A4
            Key             =   "opnfolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMAXmain.frx":06FE
            Key             =   "drvcd"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMAXmain.frx":0858
            Key             =   "drvremove"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMAXmain.frx":09B2
            Key             =   "drvfixed"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMAXmain.frx":0B0C
            Key             =   "drvremote"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMAXmain.frx":0C66
            Key             =   "mycomputer"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMAXmain.frx":0DC0
            Key             =   "drvunknown"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMAXmain.frx":0F1A
            Key             =   "audio"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMAXmain.frx":19E6
            Key             =   "api"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Caption         =   "Playlist"
      Height          =   8775
      Left            =   10200
      TabIndex        =   10
      Top             =   480
      Width           =   4935
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000000&
      Caption         =   "Time Counter"
      Height          =   1215
      Left            =   6600
      TabIndex        =   14
      Top             =   8040
      Width           =   3375
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Text            =   "Time Elapsed:"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Text            =   "Time Remaining:"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000000&
      Caption         =   "Sound Sets"
      Height          =   735
      Left            =   7320
      TabIndex        =   22
      Top             =   480
      Width           =   2775
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000000&
      Caption         =   "Move"
      Height          =   1215
      Left            =   9240
      TabIndex        =   25
      Top             =   1320
      Width           =   855
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5400
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9525
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000000&
      Caption         =   "Click To Select"
      Height          =   5775
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   6135
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   741
      ButtonWidth     =   2752
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit              "
            Key             =   "Exit"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove          "
            Key             =   "Remove"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove All        "
            Key             =   "RemAll"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000000&
      Caption         =   "CrossFade Set"
      Height          =   1095
      Left            =   8640
      TabIndex        =   35
      Top             =   6840
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Caption         =   "Seconds"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H80000000&
      Caption         =   "Audio File Type"
      Height          =   735
      Left            =   6585
      TabIndex        =   41
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   CUE NEXT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Width           =   5775
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  SELECTION PLAYING"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   6975
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer3 
      Height          =   375
      Left            =   4680
      TabIndex        =   26
      Top             =   4920
      Visible         =   0   'False
      Width           =   1170
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   -1  'True
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer2 
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Program"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu spc0 
      Caption         =   "spc0"
      Visible         =   0   'False
   End
   Begin VB.Menu PlayList 
      Caption         =   "Play List"
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuNext 
         Caption         =   "Next"
      End
      Begin VB.Menu mnuPrev 
         Caption         =   "Previous"
      End
      Begin VB.Menu mnuRem 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuRemAll 
         Caption         =   "Remove All"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
   End
   Begin VB.Menu spc1 
      Caption         =   "spc1"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuSndSet 
      Caption         =   "Sound Set"
      Begin VB.Menu mnuSet1 
         Caption         =   "Set 1"
         Begin VB.Menu mnuSel1 
            Caption         =   "Select 1"
         End
         Begin VB.Menu mnuRem1 
            Caption         =   "Remove 1"
         End
      End
      Begin VB.Menu mnuSet2 
         Caption         =   "Set 2"
         Begin VB.Menu mnuSel2 
            Caption         =   "Select 2"
         End
         Begin VB.Menu mnuRem2 
            Caption         =   "Remove 2"
         End
      End
      Begin VB.Menu mnuSet3 
         Caption         =   "Set 3"
         Begin VB.Menu mnuSel3 
            Caption         =   "Select 3"
         End
         Begin VB.Menu mnuRem3 
            Caption         =   "Remove 3"
         End
      End
      Begin VB.Menu mnuSet4 
         Caption         =   "Set 4"
         Begin VB.Menu mnuSel4 
            Caption         =   "Select 4"
         End
         Begin VB.Menu mnuRem4 
            Caption         =   "Remove 4"
         End
      End
      Begin VB.Menu mnuSet5 
         Caption         =   "Set 5"
         Begin VB.Menu mnuSel5 
            Caption         =   "Select 5"
         End
         Begin VB.Menu mnuRem5 
            Caption         =   "Remove 5"
         End
      End
   End
End
Attribute VB_Name = "frmMAXmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fPlay As String
Dim X As Integer
Dim nList As Integer
Dim nDelay As Integer
Dim IniFile As New clsIniFile
Dim bError As Boolean

Dim Audio1 As IBasicAudio
'Dim AudioCDR1 As IMediaPlayer
Dim MediaControl1 As IMediaControl
Dim MediaPosition1 As IMediaPosition

Dim Audio2 As IBasicAudio
'Dim AudioCDR2 As IMediaPlayer
Dim MediaControl2 As IMediaControl
Dim MediaPosition2 As IMediaPosition

Option Explicit

Private Sub cmdSnd1_Click()

MediaPlayer3.FileName = IniFile.ReadString("Snd1", "File")
MediaPlayer3.Play

End Sub

Private Sub cmdSnd2_Click()

MediaPlayer3.FileName = IniFile.ReadString("Snd2", "File")
MediaPlayer3.Play

End Sub

Private Sub cmdSnd3_Click()

MediaPlayer3.FileName = IniFile.ReadString("Snd3", "File")
MediaPlayer3.Play

End Sub

Private Sub cmdSnd4_Click()

MediaPlayer3.FileName = IniFile.ReadString("Snd4", "File")
MediaPlayer3.Play

End Sub

Private Sub cmdSnd5_Click()

MediaPlayer3.FileName = IniFile.ReadString("Snd5", "File")
MediaPlayer3.Play

End Sub

Private Sub Combo1_Click()
IniFile.WriteString "Delay", "Sec", Combo1.Text
TreeView1.SetFocus
End Sub

Private Sub Combo2_Click()
TreeView1.SetFocus

End Sub


Private Sub mnuPause_Click()

If MediaPosition1.CurrentPosition = 0 And _
   MediaPosition2.CurrentPosition = 0 Then
Exit Sub
End If

If MediaPosition1.CurrentPosition > 0 Then
 If MediaPosition1.CurrentPosition < MediaPosition1.Duration Then
    Label2.BackColor = &HFFFF&
    Picture2.Cls
    Picture2.FontBold = True
    Picture2.ForeColor = &HFF&
    bError = TextToPicture(Picture2, "PAUSE", eCenter)
    Toolbar1.Buttons(12).Enabled = False
    'StatusBar1.Panels(1).Text = "PAUSE"
    MediaControl1.Pause
        
End If

Else

If MediaPosition2.CurrentPosition > 0 Then
 If MediaPosition2.CurrentPosition < MediaPosition2.Duration Then
    Label2.BackColor = &HFFFF&
    Picture2.Cls
    Picture2.FontBold = True
    Picture2.ForeColor = &HFF&
    bError = TextToPicture(Picture2, "PAUSE", eCenter)
    Toolbar1.Buttons(12).Enabled = False
    'StatusBar1.Panels(1).Text = "PAUSE"
    MediaControl2.Pause
    
End If
End If
End If

End Sub

Private Sub mnuRem_Click()

If ListView1.ListItems.Count = 1 Then
Timer5.Enabled = False
End If

Dim fPlayM As String
Dim fPlayL As String

If ListView1.ListItems.Count > 0 Then
fPlayL = ListView1.SelectedItem.ListSubItems(1).Text
fPlayM = fPlay
Else
Exit Sub
End If

If fPlayL$ = fPlayM$ Then
MsgBox ("Cannot Remove... Selected Item Is Playing..."), vbInformation
Else

With ListView1
.ListItems.Remove (.SelectedItem.Index)
End With
End If

End Sub

Private Sub mnuRem1_Click()

IniFile.DeleteKey "Snd1", "File"
IniFile.WriteString "Snd1", "File", "**********"
cmdSnd1.Enabled = False

End Sub

Private Sub mnuRem2_Click()

IniFile.DeleteKey "Snd2", "File"
IniFile.WriteString "Snd2", "File", "**********"
cmdSnd2.Enabled = False

End Sub

Private Sub mnuRem3_Click()

IniFile.DeleteKey "Snd3", "File"
IniFile.WriteString "Snd3", "File", "**********"
cmdSnd3.Enabled = False

End Sub

Private Sub mnuRem4_Click()

IniFile.DeleteKey "Snd4", "File"
IniFile.WriteString "Snd4", "File", "**********"
cmdSnd4.Enabled = False

End Sub

Private Sub mnuRem5_Click()

IniFile.DeleteKey "Snd5", "File"
IniFile.WriteString "Snd5", "File", "**********"
cmdSnd5.Enabled = False

End Sub

Private Sub mnuRemAll_Click()
Timer5.Enabled = False
cmdListUp.Enabled = False
cmdListDn.Enabled = False
ListView1.ListItems.Clear
StatusBar1.Panels(2).Text = 0
Vol1 = 0
Vol2 = -4000
End Sub

Private Sub mnuSel1_Click()

Dim strSnd As String
strSnd = IniFile.ReadString("Snd1", "File")
If strSnd$ = "**********" Then


    CommonDialog1.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then
    GoTo nofileErr
     End If
    
    IniFile.IniFile = App.Path & "\MAX.ini"
    IniFile.WriteString "Snd1", "File", CommonDialog1.FileName
    
    cmdSnd1.Enabled = True
    Timer6.Enabled = True
    
Else
    
    Dim X As Integer
    X = MsgBox("Change the WAV File?", vbYesNo)

Select Case X
    
    Case 6: CommonDialog1.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
            CommonDialog1.ShowOpen
            IniFile.IniFile = App.Path & "\MAX.ini"
            IniFile.WriteString "Snd1", "File", CommonDialog1.FileName

    Case 7:
End Select
End If

Timer6.Enabled = True
Exit Sub

nofileErr:
IniFile.WriteString "Snd1", "File", "**********"

End Sub

Private Sub mnuSel2_Click()

Dim strSnd As String
strSnd = IniFile.ReadString("Snd2", "File")
If strSnd$ = "**********" Then

    CommonDialog1.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then
    GoTo nofileErr
     End If

    IniFile.IniFile = App.Path & "\MAX.ini"
    IniFile.WriteString "Snd2", "File", CommonDialog1.FileName
    
    cmdSnd2.Enabled = True
    Timer6.Enabled = True
    
Else
    
    Dim X As Integer
    X = MsgBox("Change the WAV File?", vbYesNo)

Select Case X
    
    Case 6: CommonDialog1.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
            CommonDialog1.ShowOpen
            IniFile.IniFile = App.Path & "\MAX.ini"
            IniFile.WriteString "Snd2", "File", CommonDialog1.FileName
    Case 7:
End Select
End If

Timer6.Enabled = True
Exit Sub

nofileErr:
IniFile.WriteString "Snd1", "File", "**********"

End Sub

Private Sub mnuSel3_Click()

Dim strSnd As String
strSnd = IniFile.ReadString("Snd3", "File")
If strSnd$ = "**********" Then

    CommonDialog1.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
    CommonDialog1.ShowOpen
     
    If CommonDialog1.FileName = "" Then
    GoTo nofileErr
     End If
    
    IniFile.IniFile = App.Path & "\MAX.ini"
    IniFile.WriteString "Snd3", "File", CommonDialog1.FileName
    
    cmdSnd3.Enabled = True
    Timer6.Enabled = True
    
Else
    
    Dim X As Integer
    X = MsgBox("Change the WAV File?", vbYesNo)

Select Case X
    
    Case 6: CommonDialog1.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
            CommonDialog1.ShowOpen
            IniFile.IniFile = App.Path & "\MAX.ini"
            IniFile.WriteString "Snd3", "File", CommonDialog1.FileName
    Case 7:
End Select
End If

Timer6.Enabled = True
Exit Sub

nofileErr:
IniFile.WriteString "Snd1", "File", "**********"

End Sub

Private Sub mnuSel4_Click()

Dim strSnd As String
strSnd = IniFile.ReadString("Snd4", "File")
If strSnd$ = "**********" Then

    CommonDialog1.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then
    GoTo nofileErr
     End If
    
    IniFile.IniFile = App.Path & "\MAX.ini"
    IniFile.WriteString "Snd4", "File", CommonDialog1.FileName
    
    cmdSnd4.Enabled = True
    Timer6.Enabled = True
    
Else
    
    Dim X As Integer
    X = MsgBox("Change the WAV File?", vbYesNo)

Select Case X
    
    Case 6: CommonDialog1.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
            CommonDialog1.ShowOpen
            IniFile.IniFile = App.Path & "\MAX.ini"
            IniFile.WriteString "Snd4", "File", CommonDialog1.FileName

    Case 7:
End Select
End If

Timer6.Enabled = True
Exit Sub

nofileErr:
IniFile.WriteString "Snd1", "File", "**********"

End Sub

Private Sub mnuSel5_Click()

Dim strSnd As String
strSnd = IniFile.ReadString("Snd5", "File")
If strSnd$ = "**********" Then

    CommonDialog1.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then
    GoTo nofileErr
     End If
    
    IniFile.IniFile = App.Path & "\MAX.ini"
    IniFile.WriteString "Snd5", "File", CommonDialog1.FileName
    
    cmdSnd5.Enabled = True
    Timer6.Enabled = True
    
Else
    
    Dim X As Integer
    X = MsgBox("Change the WAV File?", vbYesNo)

Select Case X
    
    Case 6: CommonDialog1.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
            CommonDialog1.ShowOpen
            IniFile.IniFile = App.Path & "\MAX.ini"
            IniFile.WriteString "Snd5", "File", CommonDialog1.FileName

    Case 7:
End Select
End If

Timer6.Enabled = True
Exit Sub

nofileErr:
IniFile.WriteString "Snd1", "File", "**********"

End Sub

Private Sub Timer6_Timer()

cmdSnd1.ToolTipText = GetFileName(IniFile.ReadString("Snd1", "File"))
cmdSnd2.ToolTipText = GetFileName(IniFile.ReadString("Snd2", "File"))
cmdSnd3.ToolTipText = GetFileName(IniFile.ReadString("Snd3", "File"))
cmdSnd4.ToolTipText = GetFileName(IniFile.ReadString("Snd4", "File"))
cmdSnd5.ToolTipText = GetFileName(IniFile.ReadString("Snd5", "File"))

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Stop"
        mnuStop_Click
    Case "Prev"
        mnuPrev_Click
    Case "Play"
        mnuPlay_Click
    Case "Next"
        mnuNext_Click
    Case "Pause"
        mnuPause_Click
  
    End Select

End Sub
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Exit"
        mnuExit_Click
    Case "Remove"
        mnuRem_Click
    Case "RemAll"
        mnuRemAll_Click
    End Select

End Sub

Public Sub LockControl(objX As Object, cLock As Boolean)
    Dim i As Long

    If cLock Then
        ' This will lock the control
        LockWindowUpdate objX.hWnd
    Else
        ' This will unlock controls
        LockWindowUpdate 0
        objX.Refresh
    End If
End Sub


Private Sub Form_Load()

LockControl List1, True
LockControl List2, True

    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS

Set MediaControl1 = New FilgraphManager
Set Audio1 = MediaControl1
Set MediaPosition1 = MediaControl1

Set MediaControl2 = New FilgraphManager
Set Audio2 = MediaControl2
Set MediaPosition2 = MediaControl2

Vol2.Value = -4000

Timer1.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False

With Combo1
.AddItem "9"
.AddItem "8"
.AddItem "7"
.AddItem "6"
.AddItem "5"
.AddItem "4"
.AddItem "3"
.AddItem "2"
.AddItem "1"
.AddItem "0"
End With

With Combo2
.Text = "MP3 Audio File"
.AddItem "MP3 Audio File"
.AddItem "MP2 Audio File"
'.AddItem "CDROM [CDA] Audio File"
End With

labLayer.Caption = "MP3 Layer"
labBitRate.Caption = "MP3 Bit Rate"
labFreqChan.Caption = "MP3 Frequency"

Text1.Font = "arial"
Text1.FontBold = True
Text1.FontSize = 14
Text1.ForeColor = vbRed
Text1.Text = "00:00:00"
Text2.Font = "arial"
Text2.FontSize = 9
Text2.FontBold = True
Text2.Text = "00:00:00"

Toolbar1.Buttons(4).Enabled = False
mnuPrev.Enabled = False
Toolbar1.Buttons(8).Enabled = False
mnuNext.Enabled = False
cmdListUp.Enabled = False
cmdListDn.Enabled = False

StatusBar1.Panels(2).Text = "0"
StatusBar1.Panels(1).Text = "Selection"
StatusBar1.Panels(3).Text = GetOSVer
    
    Dim str_CompName As String * 255
    Dim API_Return As Long
    Dim clsLoadDrives As clsLoadDrives
        Set clsLoadDrives = New clsLoadDrives
    
    API_Return& = GetComputerName(str_CompName$, Len(str_CompName$))
    modCompName.str_CompName = modStripNull.f_StripNullChr(str_CompName$)
    clsLoadDrives.subLoadTreeView
    TreeView1.Nodes(1).Expanded = True
    
    IniFile.IniFile = App.Path & "\MAX.ini"
    IniFile.WriteString "Computer", "Name", str_CompName$
    IniFile.WriteString "Operating System", "OS", GetOSVer
    If IniFile.ReadString("Delay", "Sec") > "" Then
     Combo1.Text = IniFile.ReadString("Delay", "Sec")
    Else
    Combo1.Text = "9"
    End If
    
If FileExists(IniFile.ReadString("Snd1", "File")) = False Then
IniFile.WriteString "Snd1", "File", "**********"
cmdSnd1.Enabled = False
End If

If FileExists(IniFile.ReadString("Snd2", "File")) = False Then
IniFile.WriteString "Snd2", "File", "**********"
cmdSnd2.Enabled = False
End If

If FileExists(IniFile.ReadString("Snd3", "File")) = False Then
IniFile.WriteString "Snd3", "File", "**********"
cmdSnd3.Enabled = False
End If

If FileExists(IniFile.ReadString("Snd4", "File")) = False Then
IniFile.WriteString "Snd4", "File", "**********"
cmdSnd4.Enabled = False
End If

If FileExists(IniFile.ReadString("Snd5", "File")) = False Then
IniFile.WriteString "Snd5", "File", "**********"
cmdSnd5.Enabled = False
End If

End Sub

Private Function FileExists(ssFileName As String)

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(ssFileName) Then
       FileExists = True
    Else
       FileExists = False
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)

Set MediaControl1 = Nothing
Set Audio1 = Nothing
Set MediaPosition1 = Nothing

Set MediaControl2 = Nothing
Set Audio2 = Nothing
Set MediaPosition2 = Nothing

    Unload Me
    Set frmMAXmain = Nothing
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuPlay_Click()

If Label2.BackColor = &HFFFF& Then
 If MediaPosition1.CurrentPosition > 0 Then
 Toolbar1.Buttons(12).Enabled = True
 MediaControl1.Run
 Else
 If MediaPosition2.CurrentPosition > 0 Then
 MediaControl2.Run
 End If
 End If
 
    Picture2.Cls
    Picture2.FontBold = True
    Picture2.ForeColor = &HFF&
    bError = TextToPicture(Picture2, (ListView1.SelectedItem.Text), eCenter)
    Label2.BackColor = &HFF&
    If Not bError Then
    bError = TextToPicture(Picture2, "Sorry, Unable To Display Title", eCenter)
    Label2.BackColor = &HFF&
    End If
    
Exit Sub
End If

If ListView1.ListItems.Count = 0 Then
Exit Sub
End If

MediaControl1.Stop
MediaControl2.Stop

nList = 0
nList = ListView1.SelectedItem.Index
If nList > 0 Then

fPlay = ListView1.ListItems(nList).ListSubItems(1).Text
ListView1.ListItems.Item(nList).Selected = True
ListView1.SetFocus
ListView1.DropHighlight = ListView1.SelectedItem

    Picture2.Cls
    Picture2.FontBold = True
    Picture2.ForeColor = &HFF&
    bError = TextToPicture(Picture2, (ListView1.SelectedItem.Text), eCenter)
    If Not bError Then
    bError = TextToPicture(Picture2, "Sorry, Unable To Display Title", eCenter)
    End If
        
    If (ListView1.SelectedItem.Index) + 1 <= ListView1.ListItems.Count Then
    Picture3.Cls
    bError = TextToPicture(Picture3, (ListView1.ListItems(nList + 1).Text), eCenter)
    If Not bError Then
    bError = TextToPicture(Picture3, "Sorry, Unable To Display Title", eCenter)
    End If
    Else
    Picture3.Cls
    bError = TextToPicture(Picture3, "End Of List", eCenter)
    End If
    
With Label2
.BackColor = &HFF&
.ForeColor = &HFFFF&
End With

StatusBar1.Panels(2).Text = nList

Timer2.Enabled = True
Timer4.Enabled = True

AControl

StatusBar1.Panels(1).Text = "PLAY"

End If

End Sub

Private Sub mnuNext_Click()

ProgressBar1.Value = 0

nList = (Val(StatusBar1.Panels(2).Text) + 1)

With ListView1

If Val(StatusBar1.Panels(2).Text) = 0 Then
nList = nList + 1
End If

If nList > .ListItems.Count Or _
   nList < 1 Then
Exit Sub
End If

fPlay = .ListItems(nList).ListSubItems(1).Text
.ListItems.Item(nList).Selected = True
.SetFocus
.DropHighlight = .SelectedItem
End With

    Picture2.Cls
    Picture2.FontBold = True
    Picture2.ForeColor = &HFF&
    bError = TextToPicture(Picture2, (ListView1.SelectedItem.Text), eCenter)
    If Not bError Then
    bError = TextToPicture(Picture2, "Sorry, Unable To Display Title", eCenter)
    End If
    
    If (ListView1.SelectedItem.Index) + 1 <= ListView1.ListItems.Count Then
    Picture3.Cls
    bError = TextToPicture(Picture3, (ListView1.ListItems(nList + 1).Text), eCenter)
     If Not bError Then
     bError = TextToPicture(Picture3, "Sorry, Unable To Display Title", eCenter)
     End If
    Else
    Picture3.Cls
    bError = TextToPicture(Picture3, "End Of List", eCenter)
    End If
    
AControl

Timer4.Enabled = True

StatusBar1.Panels(2).Text = nList

End Sub

Private Sub mnuPrev_Click()

ProgressBar1.Value = 0

nList = (Val(StatusBar1.Panels(2).Text) - 1)

With ListView1

If nList < 1 Or _
   nList > .ListItems.Count Then
Exit Sub
End If

fPlay = .ListItems(nList).ListSubItems(1).Text
.ListItems.Item(nList).Selected = True
.SetFocus
.DropHighlight = .SelectedItem
End With

    Picture2.Cls
    Picture2.FontBold = True
    Picture2.ForeColor = &HFF&
    bError = TextToPicture(Picture2, (ListView1.SelectedItem.Text), eCenter)
    If Not bError Then
    bError = TextToPicture(Picture2, "Sorry, Unable To Display Title", eCenter)
    End If
    
    If (ListView1.SelectedItem.Index) - 1 >= 1 Then
    Picture3.Cls
    bError = TextToPicture(Picture3, (ListView1.ListItems(nList + 1).Text), eCenter)
     If Not bError Then
     bError = TextToPicture(Picture3, "Sorry, Unable To Display Title", eCenter)
     End If
    Else
    Picture3.Cls
    bError = TextToPicture(Picture3, (ListView1.ListItems(nList + 1).Text), eCenter)
    End If
    
AControl

Timer4.Enabled = True

StatusBar1.Panels(2).Text = nList

End Sub

Private Sub mnuStop_Click()

ListView1.DropHighlight = Nothing

If ListView1.ListItems.Count = 0 Then
Exit Sub
End If

    Picture2.Cls
    Picture3.Cls
    Picture2.FontBold = True
    Picture2.ForeColor = &HFF&
    bError = TextToPicture(Picture2, (ListView1.SelectedItem.Text), eCenter)
    If Not bError Then
    bError = TextToPicture(Picture2, "Sorry, Unable To Display Title", eCenter)
    End If

If ListView1.ListItems.Count < 1 Then
Timer2.Enabled = False
Else
fPlay = ""
MediaControl1.Stop
MediaControl2.Stop
Set MediaControl1 = Nothing
Set Audio1 = Nothing
Set MediaPosition1 = Nothing

Set MediaControl2 = Nothing
Set Audio2 = Nothing
Set MediaPosition2 = Nothing

Set MediaControl1 = New FilgraphManager
Set Audio1 = MediaControl1
Set MediaPosition1 = MediaControl1

Set MediaControl2 = New FilgraphManager
Set Audio2 = MediaControl2
Set MediaPosition2 = MediaControl2

With Label2
.BackColor = &HC0C0C0
.ForeColor = &H0&
End With
Label3.BackColor = &HC0C0C0

ListView1.SelectedItem.Selected = False
ListView1.ListItems(1).Selected = True
ListView1.DropHighlight = ListView1.SelectedItem

ProgressBar1.Value = 0
Timer3.Enabled = False
Timer4.Enabled = False
Text1.Text = "00:00:00"
Text2.Text = "00:00:00"
StatusBar1.Panels(1).Text = "STOP"
StatusBar1.Panels(2).Text = "0"
Toolbar2.Buttons(6).Enabled = True
mnuRemAll.Enabled = True
End If

Vol1 = 0
Vol2 = -4000

labLayer.Caption = "MP3 Layer"
labBitRate.Caption = "MP3 Bit Rate"
labFreqChan.Caption = "MP3 Frequency"

End Sub

Private Sub Timer1_Timer()

StatusBar1.Panels(4).Text = Time
StatusBar1.Panels(5).Text = Date

If ListView1.ListItems.Count = 0 Then
Exit Sub

Else

If ListView1.SelectedItem.Index <= ListView1.ListItems.Count And _
ListView1.SelectedItem.Index > 1 Then
Toolbar1.Buttons(4).Enabled = True
mnuPrev.Enabled = True

Else

Toolbar1.Buttons(4).Enabled = False
mnuPrev.Enabled = False
End If
End If

If ListView1.ListItems.Count = 0 Then
Exit Sub

Else

If ListView1.ListItems.Count > 1 And _
ListView1.SelectedItem.Index < ListView1.ListItems.Count Then
Toolbar1.Buttons(8).Enabled = True
mnuNext.Enabled = True

Else

Toolbar1.Buttons(8).Enabled = False
mnuNext.Enabled = False

If MediaPosition1.CurrentPosition <= 0 And _
ListView1.ListItems.Count > 0 Then
Timer5.Enabled = True

Else

If MediaPosition2.CurrentPosition <= 0 And _
ListView1.ListItems.Count > 0 Then
Timer5.Enabled = True

End If
End If

End If
End If

On Local Error Resume Next
Audio1.Volume = Vol1.Value
Audio2.Volume = Vol2.Value

On Error Resume Next
If MediaPosition1.CurrentPosition > MediaPosition2.CurrentPosition Then
Text1.Text = SecToTime(MediaPosition1.Duration - MediaPosition1.CurrentPosition)
Text2.Text = SecToTime(MediaPosition1.CurrentPosition)
Else
If MediaPosition2.CurrentPosition > MediaPosition1.CurrentPosition Then
Text1.Text = SecToTime(MediaPosition2.Duration - MediaPosition2.CurrentPosition)
Text2.Text = SecToTime(MediaPosition2.CurrentPosition)

End If
End If

End Sub

Private Sub Timer2_Timer()

nDelay = (Val(Combo1.Text)) + 1

If MediaPosition1.CurrentPosition > 0 Then
Toolbar2.Buttons(6).Enabled = False
mnuRemAll.Enabled = False

 If MediaPosition1.CurrentPosition >= ((MediaPosition1.Duration) - nDelay) Then
 ProgressBar1.Value = 0
 mnuNext_Click
End If

Else

If MediaPosition2.CurrentPosition > 0 Then
Toolbar2.Buttons(6).Enabled = False
mnuRemAll.Enabled = False

 If MediaPosition2.CurrentPosition >= ((MediaPosition2.Duration) - nDelay) Then
 ProgressBar1.Value = 0
 mnuNext_Click

End If
End If
End If

Timer6.Enabled = False

End Sub

Private Sub Timer3_Timer()

If ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
Timer2.Enabled = False
End If

If MediaPosition1.CurrentPosition > MediaPosition2.CurrentPosition Then
ProgressBar1.Value = MediaPosition1.CurrentPosition
 If MediaPosition1.CurrentPosition >= MediaPosition1.Duration And _
 ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
 mnuStop_Click
End If

Else

If MediaPosition2.CurrentPosition > MediaPosition1.CurrentPosition Then
ProgressBar1.Value = MediaPosition2.CurrentPosition
 If MediaPosition2.CurrentPosition >= MediaPosition2.Duration And _
 ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
 mnuStop_Click

 End If
End If
 End If

End Sub

Private Sub Timer4_Timer()

If MediaPosition1.CurrentPosition > 0 And _
   MediaPosition2.CurrentPosition <= 0 Then
If MediaPosition1.CurrentPosition >= ((MediaPosition1.Duration) - 16) Then

Timer5.Enabled = False
cmdListUp.Enabled = False
cmdListDn.Enabled = False

Label3.BackColor = &HFFFF00

GoTo xPlay

End If

Else

If MediaPosition2.CurrentPosition > 0 And _
   MediaPosition1.CurrentPosition <= 0 Then
If MediaPosition2.CurrentPosition >= ((MediaPosition2.Duration) - 16) Then

Timer5.Enabled = False
cmdListUp.Enabled = False
cmdListDn.Enabled = False

Label3.BackColor = &HFFFF00

GoTo xPlay

End If
End If
End If
Exit Sub

xPlay:
Dim xPlay As Integer
xPlay = Val(StatusBar1.Panels(2).Text)
ListView1.SelectedItem.Selected = False
ListView1.ListItems(xPlay).Selected = True
ListView1.SelectedItem.Selected = True

End Sub

Private Sub Timer5_Timer()

  If ListView1.SelectedItem.Index = (ListView1.ListItems.Count - (ListView1.ListItems.Count - 1)) Or _
     ListView1.SelectedItem.Index = (Val(StatusBar1.Panels(2).Text) + 1) Then
     cmdListUp.Enabled = False
 
 Else
 
 If fPlay = "" Or _
 ListView1.SelectedItem.Index > (Val(StatusBar1.Panels(2).Text) + 1) Then
 cmdListUp.Enabled = True
 End If
 End If
 
If ListView1.ListItems.Count > 0 Then
  If ListView1.SelectedItem.Index <= Val(StatusBar1.Panels(2).Text) Or _
     ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
     cmdListDn.Enabled = False
  
  Else: cmdListDn.Enabled = True
  
End If
End If

End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)

    DoEvents
    Dim X As Long
    
    Me.MousePointer = 11
    For X = Node.Child.FirstSibling.Index To Node.Child.LastSibling.Index
        TreeView1_NodeClick TreeView1.Nodes(X)
    Next X
    
    Me.MousePointer = 0
    
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim str_NodePath As String
        str_NodePath$ = f_ReturnNodePath(Node.FullPath)
        
    If Not Node.Children > 0 Then
        subFolderList List1, TreeView1, str_NodePath$, Node.Index
        
If Combo2.Text = "MP3 Audio File" Then
subMP3List List2, TreeView1, str_NodePath$, Node.Index
Else
If Combo2.Text = "MP2 Audio File" Then
subMP2List List2, TreeView1, str_NodePath$, Node.Index
Else
'If Combo2.Text = "CDROM [CDA] Audio File" Then
'subCDAList List2, TreeView1, str_NodePath$, Node.Index
End If
End If
'End If
        
    End If
    Call subFileSelect(Node)
    
    If ListView1.ListItems.Count > 0 Then
     If ListView1.SelectedItem.Index < ListView1.ListItems.Count Then
     Timer2.Enabled = True
    End If
     End If
     
    End Sub

Private Sub subFileSelect(ByVal Node As MSComctlLib.Node)

Dim oTreeView As TreeView
 Set oTreeView = frmMAXmain.TreeView1
Dim sFileName As String
Dim sPath As String
Dim nPath As String
Dim ckFile As String
Dim rFile As String
Dim sExt As String
Dim nLength As Integer
Dim lstPathName As ListItem
Dim lstSubItem As ListSubItem
Dim fType As String

If Combo2.Text = "MP3 Audio File" Then
fType = ".mp3"
Else
If Combo2.Text = "MP2 Audio File" Then
fType = ".wav"
Else
'If Combo2.Text = "CDROM [CDA] Audio File" Then
'fType = ".cda"
End If
End If
'End If

On Local Error Resume Next
ckFile = (oTreeView.SelectedItem.Text)
nLength = Len(ckFile$) - 4
sExt = LCase(Right(ckFile$, 4))
rFile = Mid(ckFile$, 1, nLength)
sFileName = rFile$ & sExt$

If Right(sFileName$, 4) = fType Then
sPath$ = f_ReturnFilePath(Node.FullPath)

Set lstPathName = ListView1.ListItems.Add(, , rFile$)
Set lstSubItem = lstPathName.ListSubItems.Add(, , sPath$)
ListView1.SetFocus

Set oTreeView = Nothing
Set lstPathName = Nothing
Set lstSubItem = Nothing

Else
Set oTreeView = Nothing
Exit Sub
End If

End Sub

Private Function SecToTime(NewSec As Double) As String

On Error Resume Next
Dim Secx, Minx, Hourx
NewSec = Int(NewSec)
If NewSec < 1 Then SecToTime = "00:00:00": Exit Function
Secx = NewSec - Int(NewSec / 60) * 60
Minx = Int((NewSec - Int(NewSec / 3600) * 3600) / 60)
Hourx = Int(NewSec / 3600)
If Int(Hourx) > 24 Then
SecToTime = "24:59:59"
Else
SecToTime = Format(Str(Hourx) & ":" & Str(Minx) & ":" & Str(Secx), "hh:mm:ss")
End If

End Function

Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Set ListView1.DropHighlight = Nothing
        
End Sub

Private Sub cmdListUP_Click()

If ListView1.SelectedItem.ListSubItems(1).Text = fPlay Then
MsgBox ("Cannot Move Item... Item Is In Use..."), vbInformation
Exit Sub
End If

Dim lstPathName As ListItem
Dim lstSubItem As ListSubItem

If ListView1.SelectedItem.Index = 1 Then
Set ListView1.DropHighlight = ListView1.SelectedItem

Else
If ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
    Set lstPathName = ListView1.ListItems.Add(ListView1.SelectedItem.Index - 1, , ListView1.SelectedItem.Text)
        lstPathName.SubItems(1) = ListView1.SelectedItem.SubItems(1)
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    Set ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index - 1)
    Set ListView1.DropHighlight = ListView1.SelectedItem
Else
    Set lstPathName = ListView1.ListItems.Add(ListView1.SelectedItem.Index - 1, , ListView1.SelectedItem.Text)
        lstPathName.SubItems(1) = ListView1.SelectedItem.SubItems(1)
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    Set ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index - 2)
    Set ListView1.DropHighlight = ListView1.SelectedItem

End If
End If

End Sub

Private Sub cmdListDn_Click()

If ListView1.SelectedItem.ListSubItems(1).Text = fPlay Then
MsgBox ("Cannot Move Item... Item Is In Use..."), vbInformation
Exit Sub
End If

Dim lstPathName As ListItem
Dim lstSubItem As ListSubItem

If ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
    Set ListView1.SelectedItem = ListView1.ListItems(ListView1.ListItems.Count)
    Set ListView1.DropHighlight = ListView1.SelectedItem
Else
    Set lstPathName = ListView1.ListItems.Add(ListView1.SelectedItem.Index + 2, , ListView1.SelectedItem.Text)
        lstPathName.SubItems(1) = ListView1.SelectedItem.SubItems(1)
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    Set ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index + 1)
    Set ListView1.DropHighlight = ListView1.SelectedItem

End If
End Sub

Public Sub AControl()

labLayer.Caption = ""
labBitRate.Caption = ""
labFreqChan.Caption = ""

If MediaPosition1.CurrentPosition <= 0 And _
   MediaPosition2.CurrentPosition <= 0 Then
   GetMP3Inf
   On Error GoTo badAudio
   MediaControl1.RenderFile fPlay
   MediaControl1.Run
   ProgressBar1.Max = MediaPosition1.Duration
   Timer2.Enabled = True
   Timer3.Enabled = True
   Exit Sub
End If

If MediaPosition1.CurrentPosition > MediaPosition2.CurrentPosition Then

   Switch_1
   
Else

If MediaPosition2.CurrentPosition > MediaPosition1.CurrentPosition Then

   Switch_2
   
End If
End If

Exit Sub

badAudio:

If ListView1.SelectedItem.Index < ListView1.ListItems.Count Then
mnuNext_Click
Else
mnuStop_Click
End If

End Sub

Public Sub Switch_1()

labLayer.Caption = ""
labBitRate.Caption = ""
labFreqChan.Caption = ""

Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False

ProgressBar1.Value = 0

Set MediaControl2 = New FilgraphManager
Set Audio2 = MediaControl2
Set MediaPosition2 = MediaControl2

On Error GoTo badAudio
MediaControl2.RenderFile fPlay
MediaControl2.Run
ProgressBar1.Max = MediaPosition2.Duration

Timer7.Enabled = True
Exit Sub

badAudio:

MediaControl2.Stop
StatusBar1.Panels(2).Text = (Val(StatusBar1.Panels(2).Text) + 1)
If ListView1.SelectedItem.Index < ListView1.ListItems.Count Then
mnuNext_Click
Else
mnuStop_Click
End If

End Sub

Public Sub Switch_2()

labLayer.Caption = ""
labBitRate.Caption = ""
labFreqChan.Caption = ""

Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False

ProgressBar1.Value = 0

Set MediaControl1 = New FilgraphManager
Set Audio1 = MediaControl1
Set MediaPosition1 = MediaControl1
 
On Error GoTo badAudio
MediaControl1.RenderFile fPlay
MediaControl1.Run
ProgressBar1.Max = MediaPosition1.Duration

Timer8.Enabled = True
Exit Sub

badAudio:

MediaControl1.Stop
StatusBar1.Panels(2).Text = (Val(StatusBar1.Panels(2).Text) + 1)
If ListView1.SelectedItem.Index < ListView1.ListItems.Count Then
mnuNext_Click
Else
mnuStop_Click
End If

End Sub

Private Sub Timer7_Timer()

If Vol1.Value > -4000 Then
Timer1.Enabled = False
Toolbar1.Buttons(4).Enabled = False
mnuPrev.Enabled = False
Toolbar1.Buttons(8).Enabled = False
mnuNext.Enabled = False
Toolbar1.Buttons(6).Enabled = False
mnuPlay.Enabled = False
Toolbar1.Buttons(12).Enabled = False
mnuPause.Enabled = False
ProgressBar1.Value = MediaPosition2.CurrentPosition
Vol1.Value = Vol1.Value - 40
Audio1.Volume = Vol1.Value
Vol2.Value = 0
Audio2.Volume = Vol2.Value
GetMP3Inf
End If

If Vol1.Value = -3960 Then
Timer1.Enabled = True
Toolbar1.Buttons(6).Enabled = True
mnuPlay.Enabled = True
Toolbar1.Buttons(12).Enabled = True
mnuPause.Enabled = True
MediaControl1.Stop
MediaPosition1.CurrentPosition = 0
Label3.BackColor = &HC0C0C0
Vol1.Value = -4000
Audio1.Volume = Vol1.Value
Timer7.Enabled = False
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
End If

End Sub

Private Sub Timer8_Timer()

If Vol2.Value > -4000 Then
Timer1.Enabled = False
Toolbar1.Buttons(4).Enabled = False
mnuPrev.Enabled = False
Toolbar1.Buttons(8).Enabled = False
mnuNext.Enabled = False
Toolbar1.Buttons(6).Enabled = False
mnuPlay.Enabled = False
Toolbar1.Buttons(12).Enabled = False
mnuPause.Enabled = False
ProgressBar1.Value = MediaPosition1.CurrentPosition
Vol2.Value = Vol2.Value - 40
Audio2.Volume = Vol2.Value
Vol1.Value = 0
Audio1.Volume = Vol1.Value
GetMP3Inf
End If

If Vol2.Value = -3960 Then
Timer1.Enabled = True
Toolbar1.Buttons(6).Enabled = True
mnuPlay.Enabled = True
Toolbar1.Buttons(12).Enabled = True
mnuPause.Enabled = True
MediaControl2.Stop
MediaPosition2.CurrentPosition = 0
Label3.BackColor = &HC0C0C0
Vol2.Value = -4000
Audio2.Volume = Vol2.Value
Timer8.Enabled = False
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
End If

End Sub

Private Sub GetMP3Inf()
  Dim accMP3Info As MP3Info
  
  getMP3Info fPlay, accMP3Info
  
  labLayer = accMP3Info.MPEG & " " & accMP3Info.LAYER
  labBitRate = accMP3Info.BITRATE
  labFreqChan = accMP3Info.FREQ & " " & accMP3Info.CHANNELS
End Sub


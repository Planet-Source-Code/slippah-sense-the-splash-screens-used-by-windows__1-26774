VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sense Your Startup, Shut Down, And Turn Off Computer Graphics"
   ClientHeight    =   6870
   ClientLeft      =   150
   ClientTop       =   465
   ClientWidth     =   5295
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Starting Up "
      TabPicture(0)   =   "frmMain.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Shutting Down Logo"
      TabPicture(1)   =   "frmMain.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Shut Down Computer"
      TabPicture(2)   =   "frmMain.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "It Is Now Safe To Turn Off Your Computer Logo"
         Height          =   6375
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   5055
         Begin VB.PictureBox ShutDown 
            AutoSize        =   -1  'True
            Height          =   6015
            Left            =   120
            ScaleHeight     =   5955
            ScaleWidth      =   4755
            TabIndex        =   6
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Shutting Down"
         Height          =   6375
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   5055
         Begin VB.PictureBox ShuttingDown 
            AutoSize        =   -1  'True
            Height          =   6015
            Left            =   120
            ScaleHeight     =   5955
            ScaleWidth      =   4755
            TabIndex        =   4
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Starting Up Logo"
         Height          =   6375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5055
         Begin VB.PictureBox Startup 
            AutoSize        =   -1  'True
            Height          =   6015
            Left            =   120
            ScaleHeight     =   5955
            ScaleWidth      =   4755
            TabIndex        =   2
            Top             =   240
            Width           =   4815
         End
      End
   End
   Begin VB.Menu mnuFile1 
      Caption         =   "File1"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy1 
         Caption         =   "&Copy Image To Clipboard"
      End
      Begin VB.Menu Slash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh1 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuFile2 
      Caption         =   "File2"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy2 
         Caption         =   "&Copy Image To Clipboard"
      End
      Begin VB.Menu Slash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh2 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuFile3 
      Caption         =   "File3"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy3 
         Caption         =   "&Copy Image To Clipboard"
      End
      Begin VB.Menu Slash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh3 
         Caption         =   "&Refresh"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Startup.Picture = LoadPicture("C:\Logo.sys")
ShuttingDown.Picture = LoadPicture("C:\Windows\Logow.sys")
ShutDown.Picture = LoadPicture("C:\Windows\Logos.sys")
End Sub

Private Sub mnuCopy1_Click()
Clipboard.Clear
Clipboard.SetData Startup.Picture
End Sub

Private Sub mnuCopy2_Click()
Clipboard.Clear
Clipboard.SetData ShuttingDown.Picture
End Sub

Private Sub mnuCopy3_Click()
Clipboard.Clear
Clipboard.SetData ShutDown.Picture
End Sub

Private Sub mnuRefresh1_Click()
Startup.Picture = LoadPicture("C:\Logo.sys")
End Sub

Private Sub mnuRefresh2_Click()
ShuttingDown.Picture = LoadPicture("C:\Windows\Logow.sys")
End Sub

Private Sub mnuRefresh3_Click()
ShutDown.Picture = LoadPicture("C:\Windows\Logos.sys")
End Sub

Private Sub Startup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Me.PopupMenu mnuFile1
End If
End Sub

Private Sub ShuttingDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Me.PopupMenu mnuFile2
End If
End Sub

Private Sub ShutDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Me.PopupMenu mnuFile3
End If
End Sub

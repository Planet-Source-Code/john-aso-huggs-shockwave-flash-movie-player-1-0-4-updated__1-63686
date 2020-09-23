VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form Form1 
   Caption         =   "Huggs' Shockwave Flash Movie Player 1.0"
   ClientHeight    =   5865
   ClientLeft      =   3540
   ClientTop       =   1320
   ClientWidth     =   7470
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash 
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7035
      _cx             =   12409
      _cy             =   9234
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu bar 
         Caption         =   "-"
      End
      Begin VB.Menu quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu controls 
      Caption         =   "Controls"
      Begin VB.Menu mnushow 
         Caption         =   "Show"
      End
   End
   Begin VB.Menu about 
      Caption         =   "Help"
      Begin VB.Menu aboutme 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Sub aboutme_Click()
frmAbout.Show vbModal
End Sub

Private Sub Form_Load()
If Command <> "" Then
flash.Movie = Command
Me.Caption = Command & " - Huggs' Flash Movie Player"
End If
Timer1.Interval = 250
    Timer1.Enabled = True
flash.Quality = 1
End Sub

Private Sub Form_Resize()
flash.Width = Me.Width
flash.Height = Me.Height
End Sub

Private Sub mnushow_Click()
Form2.Show vbModal
End Sub

Private Sub open_Click()
CD.Filter = "Shockwave Flash Files (*.swf)|*.swf"
CD.ShowOpen
If CD.FileName <> "" Then
flash.Movie = CD.FileName
Me.Caption = CD.FileTitle & " - Huggs' Flash Movie Player"
Form2.lblTotalFrame.Caption = flash.TotalFrames
End If
End Sub

Private Sub quit_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
    If flash.CurrentFrame <> -1 Then
        Form2.lblCurrFrame.Caption = flash.CurrentFrame
    End If
    Form2.lblLoaded = flash.PercentLoaded
End Sub
    
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub
    



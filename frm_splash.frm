VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_splash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Web Messenger - Please wait..."
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4110
   ControlBox      =   0   'False
   Icon            =   "frm_splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tim_delay 
      Left            =   120
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar pro_delay 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lbl_version 
      Alignment       =   2  'Center
      Caption         =   "<VERSION DETAILS>"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2460
      Width           =   3855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2325
      Left            =   120
      Picture         =   "frm_splash.frx":0CCA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frm_splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************
Const timer_interval = 50
Const timer_max = 75
Dim current_sec As Integer
'********************************

Private Sub Form_Load()
    lbl_version.Caption = App.ProductName & " : Version " & App.Major & "." & App.Minor & "." & App.Revision
    tim_delay.Interval = timer_interval
    pro_delay.max = timer_max
    
    setup_events
    For loadevents = 0 To noof_events - 1
        snd_events(loadevents).load
    Next loadevents
    
    setup_icons
    
    setup_colvars
    For l_colvar = 0 To noof_colvars - 1
        wm_colvars(l_colvar).load_vars
    Next l_colvar
    
    init_buddies
    For l_buddy = 0 To max_buddies - 1
        buddies(l_buddy).load l_buddy
    Next l_buddy
    
    snd_events(event_onload).start
End Sub

Private Sub Form_Paint()
    'POSITION FOR USER NAME INPUT BOX
    input_x = Me.Left
    input_y = Me.Top
End Sub

Private Sub tim_delay_Timer()
    current_sec = increment_counter(current_sec, timer_max)
    If current_sec <> 0 Then
        pro_delay = current_sec
    Else
        frm_main.Show
        Unload Me
    End If
End Sub

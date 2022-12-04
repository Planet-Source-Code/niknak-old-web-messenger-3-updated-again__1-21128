VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Messenger - Main menu"
   ClientHeight    =   4905
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   5535
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm_client 
      Caption         =   "I want to be a Client"
      Height          =   2295
      Left            =   60
      TabIndex        =   1
      Top             =   2040
      Width           =   5415
      Begin VB.ComboBox com_buddy 
         Height          =   315
         Left            =   1920
         TabIndex        =   15
         Text            =   "Buddy"
         Top             =   840
         Width           =   3375
      End
      Begin VB.CommandButton cmd_begin 
         Caption         =   "Begin"
         Height          =   855
         Index           =   1
         Left            =   4560
         Picture         =   "frm_main.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txt_serport 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Text            =   "Server's TCP/IP Port"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txt_serip 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Text            =   "Server's IP Address"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lbl_buddy 
         Caption         =   "Current Buddies"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label lbl_clienthelp 
         Caption         =   "You must obtain the details below from the Server before you begin.  These are essential to connect successfully."
         Height          =   495
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label lbl_serport 
         Caption         =   "Server's TCP/IP Port"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lbl_serip 
         Caption         =   "Server's IP Address"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "frm_main.frx":110C
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame frm_server 
      Caption         =   "I want to be the Server"
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmd_begin 
         Caption         =   "Begin"
         Height          =   855
         Index           =   0
         Left            =   4560
         Picture         =   "frm_main.frx":154E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_myport 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Text            =   "My TCP/IP Port"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txt_myip 
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "My IP Address"
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lbl_serverhelp 
         Caption         =   "Please make sure that you inform the Client of your IP Address and TCP/IP Port before you choose to begin."
         Height          =   495
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label lbl_myport 
         Caption         =   "My TCP/IP Port"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lbl_myip 
         Caption         =   "My IP Address"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   120
         Picture         =   "frm_main.frx":1990
         Top             =   240
         Width           =   480
      End
   End
   Begin MSWinsockLib.Winsock win_details 
      Left            =   60
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sta_info 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   4410
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4842
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   4842
         EndProperty
      EndProperty
   End
   Begin VB.Menu men_file 
      Caption         =   "File"
      Begin VB.Menu Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu men_exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu men_messages 
      Caption         =   "Messages"
      Begin VB.Menu men_changeuser 
         Caption         =   "Change User Name"
      End
   End
   Begin VB.Menu men_options 
      Caption         =   "Options"
      Begin VB.Menu men_sndevents 
         Caption         =   "Sound Events"
      End
      Begin VB.Menu men_avicons 
         Caption         =   "Available Icons"
      End
      Begin VB.Menu men_colours 
         Caption         =   "Colours"
      End
      Begin VB.Menu men_buddies 
         Caption         =   "Buddies"
      End
   End
   Begin VB.Menu men_help 
      Caption         =   "Help"
   End
   Begin VB.Menu men_about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_begin_Click(index As Integer)
    local_port = txt_myport
    remote_ip = txt_serip
    remote_port = txt_serport
    If index = 0 Then
        cliserv = iam_server
    Else
        cliserv = iam_client
    End If
    frm_messenger.Show
    Me.Hide
End Sub

Public Sub refresh_combo()
    com_buddy.Clear
    For r_buddy = 0 To max_buddies
        If buddies(r_buddy).buddy_name <> "" Then com_buddy.AddItem buddies(r_buddy).buddy_name
    Next r_buddy
End Sub

Private Function search_buddies(search As String) As Integer
    Dim checked As Boolean
    checked = False
    For c_buddy = 0 To max_buddies
        If buddies(c_buddy).buddy_name = search Then
            If (checked = False) Then
                search_buddies = c_buddy
                checked = True
            End If
        End If
    Next c_buddy
End Function

Private Sub com_buddy_Click()
    Dim current_buddy As Integer
    current_buddy = search_buddies(com_buddy)
    If buddies(current_buddy).buddy_name <> "" Then
        txt_serip = buddies(current_buddy).buddy_ip
        txt_serport = buddies(current_buddy).buddy_port
    End If
End Sub

Private Sub Form_Load()
    my_name = InputBox("Please enter your name", "Web Messenger - User", , input_x, input_y)
    If my_name = "" Then my_name = "User"
    sta_info.Panels(1).Text = "Welcome to Web Messenger " & my_name
    load_settings
    load_options
    refresh_combo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    save_settings
    save_options
    snd_events(event_onunload).start
    Unload frm_sndevents
    Unload frm_icons
    Unload frm_colours
    Unload frm_buddies
    Unload frm_help
    Unload frm_about
End Sub

Private Sub load_settings()
    With frm_main
        load_window (.Caption)
        If win_top <> 0 Then .Top = win_top
        If win_left <> 0 Then .Left = win_left
    End With
End Sub

Private Sub save_settings()
    With frm_main
        save_window .Caption, .Top, .Left
    End With
End Sub

Private Sub save_options()
    SaveSetting App.ProductName, frm_main.Caption, "Options", "SAVED"
    SaveSetting App.ProductName, frm_main.Caption, "My IP Address", txt_myip.Text
    SaveSetting App.ProductName, frm_main.Caption, "My TCP/IP Port", txt_myport.Text
    SaveSetting App.ProductName, frm_main.Caption, "Server's IP Address", txt_serip.Text
    SaveSetting App.ProductName, frm_main.Caption, "Server's TCP/IP Port", txt_serport.Text
End Sub

Private Sub load_options()
    If GetSetting(App.ProductName, frm_main.Caption, "Options") = "SAVED" Then
        txt_myport = GetSetting(App.ProductName, frm_main.Caption, "My TCP/IP Port")
        txt_serip = GetSetting(App.ProductName, frm_main.Caption, "Server's IP Address")
        txt_serport = GetSetting(App.ProductName, frm_main.Caption, "Server's TCP/IP Port")
    Else
        txt_myport = "My TCP/IP Port"
        txt_serip = "Server's IP Address"
        txt_serport = "Server's TCP/IP Port"
    End If
    txt_myip = win_details.LocalIP
End Sub

Private Sub men_about_Click()
    load frm_about
    frm_about.Show
End Sub

Private Sub men_avicons_Click()
    load frm_icons
    frm_icons.Show
End Sub

Private Sub men_buddies_Click()
    load frm_buddies
    frm_buddies.Show
End Sub

Private Sub men_changeuser_Click()
    my_name = InputBox("Please enter your name", "Web Messenger - User", , input_x, input_y)
    If my_name = "" Then my_name = "User"
    sta_info.Panels(1).Text = "Welcome to Web Messenger " & my_name
End Sub

Private Sub men_colours_Click()
    load frm_colours
    frm_colours.Show
End Sub

Private Sub men_exit_Click()
    Unload Me
End Sub

Private Sub men_help_Click()
    frm_help.Show
    frm_help.Caption = "Help-" & Me.Caption
    frm_help.lbl_help.Caption = help_main
End Sub

Private Sub men_sndevents_Click()
    load frm_sndevents
    frm_sndevents.Show
End Sub

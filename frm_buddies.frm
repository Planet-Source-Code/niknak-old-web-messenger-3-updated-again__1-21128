VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_buddies 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buddies"
   ClientHeight    =   4050
   ClientLeft      =   150
   ClientTop       =   705
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlg_msagent 
      Left            =   60
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fra_details 
      Caption         =   "Details"
      Height          =   3495
      Left            =   3180
      TabIndex        =   5
      Top             =   60
      Width           =   4395
      Begin VB.CommandButton cmd_apply 
         Caption         =   "Apply"
         Height          =   315
         Left            =   2580
         TabIndex        =   17
         Top             =   3060
         Width           =   795
      End
      Begin VB.CommandButton cmd_add 
         Caption         =   "Add"
         Height          =   315
         Left            =   3480
         TabIndex        =   16
         Top             =   3060
         Width           =   795
      End
      Begin VB.CheckBox chk_enabled 
         Alignment       =   1  'Right Justify
         Caption         =   "I wish to use the MSAgent for this buddy"
         Height          =   375
         Left            =   180
         TabIndex        =   15
         Top             =   2520
         Width           =   4095
      End
      Begin VB.CommandButton cmd_open 
         Caption         =   "Open"
         Height          =   315
         Left            =   3480
         TabIndex        =   14
         Top             =   2160
         Width           =   795
      End
      Begin VB.TextBox txt_msagent 
         Height          =   375
         Left            =   1740
         TabIndex        =   12
         Text            =   "MSAgent file"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txt_bname 
         Height          =   375
         Left            =   1740
         TabIndex        =   10
         Text            =   "Buddies Name"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txt_bport 
         Height          =   375
         Left            =   1740
         TabIndex        =   8
         Text            =   "Buddies TCP/IP Port"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txt_bip 
         Height          =   375
         Left            =   1740
         TabIndex        =   6
         Text            =   "Buddies IP Address"
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lbl_msagent 
         Caption         =   "MSAgent File"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lbl_bname 
         Caption         =   "Buddies Name"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbl_bport 
         Caption         =   "Buddies TCP/IP Port"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lbl_bip 
         Caption         =   "Buddies IP Address"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame fra_buddies 
      Caption         =   "Current Buddies"
      Height          =   3495
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3015
      Begin VB.CommandButton cmd_delete 
         Caption         =   "Delete"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   3060
         Width           =   795
      End
      Begin VB.ListBox lst_buddies 
         Height          =   2595
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5880
      TabIndex        =   1
      Top             =   3660
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "Ok"
      Height          =   315
      Left            =   6780
      TabIndex        =   0
      Top             =   3660
      Width           =   795
   End
   Begin VB.Menu men_help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frm_buddies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_add_Click()
    next_available = search_buddies("")
    buddies(next_available).buddy_name = txt_bname
    buddies(next_available).buddy_ip = txt_bip
    buddies(next_available).buddy_port = txt_bport
    buddies(next_available).buddy_file = txt_msagent
    buddies(next_available).buddy_enabled = chk_enabled
    refresh_list
End Sub

Private Sub cmd_apply_Click()
    buddies(search_buddies(lst_buddies.List(lst_buddies.ListIndex))).buddy_name = txt_bname
    buddies(search_buddies(lst_buddies.List(lst_buddies.ListIndex))).buddy_ip = txt_bip
    buddies(search_buddies(lst_buddies.List(lst_buddies.ListIndex))).buddy_port = txt_bport
    buddies(search_buddies(lst_buddies.List(lst_buddies.ListIndex))).buddy_file = txt_msagent
    buddies(search_buddies(lst_buddies.List(lst_buddies.ListIndex))).buddy_enabled = chk_enabled
End Sub

Private Sub cmd_cancel_Click()
    Unload Me
End Sub

Private Sub cmd_delete_Click()
    If lst_buddies.ListIndex <> -1 Then
        DeleteSetting App.ProductName & " Buddies", Str(search_buddies(lst_buddies.List(lst_buddies.ListIndex))), "name"
        buddies(search_buddies(lst_buddies.List(lst_buddies.ListIndex))).clear_all
        refresh_list
    End If
End Sub

Private Sub cmd_ok_Click()
    For s_buddy = 0 To max_buddies
        buddies(s_buddy).save s_buddy
    Next s_buddy
    Unload Me
End Sub

Private Sub cmd_open_Click()
    dlg_msagent.Filter = "MSAgent (*.acs)|*.acs|"
    dlg_msagent.ShowOpen
    If dlg_msagent.filename <> "" Then
        txt_msagent = dlg_msagent.filename
    End If
End Sub

Private Sub Form_Load()
    load_settings
    refresh_list
End Sub

Private Sub Form_Unload(Cancel As Integer)
    save_settings
    frm_main.refresh_combo
End Sub

Private Sub lst_buddies_Click()
    If lst_buddies.ListCount <> -1 Then
        get_details search_buddies(lst_buddies.List(lst_buddies.ListIndex))
   End If
End Sub

Private Sub refresh_list()
    lst_buddies.Clear
    For r_buddy = 0 To max_buddies
        If buddies(r_buddy).buddy_name <> "" Then lst_buddies.AddItem buddies(r_buddy).buddy_name
    Next r_buddy
End Sub

Private Sub get_details(index As Variant)
    txt_bname = buddies(index).buddy_name
    txt_bip = buddies(index).buddy_ip
    txt_bport = buddies(index).buddy_port
    txt_msagent = buddies(index).buddy_file
    chk_enabled = buddies(index).buddy_enabled
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

Private Sub men_help_Click()
    frm_help.Show
    frm_help.Caption = "Help-" & Me.Caption
    frm_help.lbl_help.Caption = help_buddies
End Sub

Private Sub load_settings()
    With frm_buddies
        load_window (.Caption)
        If win_top <> 0 Then .Top = win_top
        If win_left <> 0 Then .Left = win_left
    End With
End Sub

Private Sub save_settings()
    save_window Me.Caption, Me.Top, Me.Left
End Sub

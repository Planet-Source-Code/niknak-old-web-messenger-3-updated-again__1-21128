VERSION 5.00
Begin VB.Form frm_colours 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Settings"
   ClientHeight    =   4470
   ClientLeft      =   150
   ClientTop       =   705
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_ok 
      Caption         =   "Ok"
      Height          =   315
      Left            =   4860
      TabIndex        =   4
      Top             =   4080
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   4080
      Width           =   795
   End
   Begin VB.PictureBox pic_colour 
      Height          =   315
      Left            =   900
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmd_change 
      Caption         =   "Change"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   4080
      Width           =   795
   End
   Begin VB.ListBox lst_colvars 
      Height          =   3960
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5595
   End
   Begin VB.Menu men_help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frm_colours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancel_Click()
    Unload Me
End Sub

Private Sub cmd_change_Click()
    If (lst_colvars.ListIndex) > -1 Then
        frm_palette.Show
    End If
End Sub

Private Sub cmd_ok_Click()
    For s_colvar = 0 To noof_colvars - 1
        wm_colvars(s_colvar).save_vars
    Next s_colvar
    Unload Me
End Sub

Private Sub Form_Load()
    load_settings
    refresh_colvars
End Sub

Private Sub refresh_colvars()
    For r_colvars = 0 To noof_colvars
        With wm_colvars(r_colvars)
            lst_colvars.AddItem .variable_description
        End With
    Next r_colvars
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frm_palette
    save_settings
End Sub

Private Sub lst_colvars_Click()
    pic_colour.BackColor = wm_colvars(lst_colvars.ListIndex).variable_colour_win
End Sub

Private Sub men_help_Click()
    frm_help.Show
    frm_help.Caption = "Help-" & Me.Caption
    frm_help.lbl_help.Caption = help_colours
End Sub

Private Sub load_settings()
    With frm_colours
        load_window (.Caption)
        If win_top <> 0 Then .Top = win_top
        If win_left <> 0 Then .Left = win_left
    End With
End Sub

Private Sub save_settings()
    save_window Me.Caption, Me.Top, Me.Left
End Sub

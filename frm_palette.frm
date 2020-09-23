VERSION 5.00
Object = "{8E5DCCD3-7FCC-401F-8868-65B15168B825}#20.0#0"; "Quick Palette.ocx"
Begin VB.Form frm_palette 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Palette"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1620
      TabIndex        =   2
      Top             =   1980
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "Ok"
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   1980
      Width           =   795
   End
   Begin QuickPalette.palette pal_palette 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   3625
   End
End
Attribute VB_Name = "frm_palette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancel_Click()
    Unload Me
End Sub

Private Sub cmd_ok_Click()
    wm_colvars(frm_colours.lst_colvars.ListIndex).variable_colour_win = pal_palette.foreground_colour_win
    wm_colvars(frm_colours.lst_colvars.ListIndex).variable_colour_html = pal_palette.foreground_colour_html
    frm_colours.pic_colour.BackColor = wm_colvars(frm_colours.lst_colvars.ListIndex).variable_colour_win
    Unload Me
End Sub

Private Sub Form_Load()
    pal_palette.change_cbar False
End Sub

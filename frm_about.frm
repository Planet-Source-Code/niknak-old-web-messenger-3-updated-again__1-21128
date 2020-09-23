VERSION 5.00
Begin VB.Form frm_about 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Web Messenger"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbl_add 
      Alignment       =   1  'Right Justify
      Caption         =   "nickpateman@blueyonder.co.uk"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3900
      Width           =   3855
   End
   Begin VB.Label lbl_sig 
      Caption         =   "Nick Pateman"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Label lbl_about 
      Alignment       =   2  'Center
      Caption         =   $"frm_about.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2325
      Left            =   120
      Picture         =   "frm_about.frx":00FD
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

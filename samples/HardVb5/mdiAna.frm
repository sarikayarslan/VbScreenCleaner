VERSION 5.00
Begin VB.MDIForm mdiAna 
   BackColor       =   &H8000000C&
   Caption         =   "Ana Pencere"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAlt 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   8940
      TabIndex        =   0
      Top             =   5400
      Width           =   9000
      Begin VB.CommandButton cmdMdiGizli 
         Caption         =   "Gizli"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox picDurum 
      Align           =   1  'Align Top
      Height          =   400
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   8940
      TabIndex        =   2
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "mdiAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    picDurum.Visible = False
    ModulYaz
    HedefiGoster
    frmZor.Show
End Sub

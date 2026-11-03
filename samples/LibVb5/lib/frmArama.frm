VERSION 5.00
Begin VB.Form frmArama 
   Caption         =   "Ortak Arama"
   ClientHeight    =   2000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2000
   ScaleWidth      =   4000
   Begin VB.TextBox txtAra 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdAramaGizli 
      Caption         =   "Eski"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmArama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAramaGizli_Click()
    MsgBox "eski arama"
End Sub

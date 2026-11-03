VERSION 5.00
Begin VB.Form frmAna 
   Caption         =   "Ana"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5000
   Begin VB.CommandButton cmdAra 
      Caption         =   "Ara"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtLibYazar 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtLibOkur 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLibGoster 
      Caption         =   "Lib gosterir"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTemiz 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    OrtakHazirla
    txtTemiz.Text = ""
End Sub

Private Sub cmdAra_Click()
    frmArama.Show
End Sub

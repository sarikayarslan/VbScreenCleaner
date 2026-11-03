VERSION 5.00
Begin VB.Form frmTr 
   Caption         =   "Türkçe Ekran"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4000
   Begin VB.Label lblÜnvan 
      Caption         =   "Ünvan: ÐÝþçö"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtÞirketAdý 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Þirket"
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmTr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Þirket adý artýk gösterilmiyor
    txtÞirketAdý.Text = "Çaðrý Þahin": lblÜnvan.Caption = "Müdür"
End Sub

VERSION 5.00
Begin VB.UserControl ctlKutu 
   ClientHeight    =   1000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2000
   ScaleHeight     =   1000
   ScaleWidth      =   2000
   Begin VB.TextBox txtIc 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblIcOkunan 
      Caption         =   "Okunan"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "ctlKutu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Initialize()
    txtIc.Tag = "a"
End Sub

Public Property Get Metin() As String
Attribute Metin.VB_UserMemId = 0
    Metin = lblIcOkunan.Caption
End Property

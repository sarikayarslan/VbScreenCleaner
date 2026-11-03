VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSekme 
   Caption         =   "Sekmeler"
   ClientHeight    =   4000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4000
   ScaleWidth      =   6000
   Begin TabDlg.SSTab tabAna 
      Height          =   3000
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5000
      _ExtentX        =   8819
      _ExtentY        =   5292
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Genel"
      TabPicture(0)   =   "frmSekme.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtSekmeGizli"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtSekmeGorunur"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtSekme3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detay"
      TabPicture(1)   =   "frmSekme.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDetayGizli"
      Tab(1).Control(1)=   "txtDetay"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtSekmeGizli 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtSekmeGorunur 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtSekme3 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtDetayGizli 
         Height          =   285
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtDetay 
         Height          =   285
         Left            =   -74880
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Label lblDisari 
      Caption         =   "Disari"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAyni 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmSekme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtAyni.Visible = True
End Sub

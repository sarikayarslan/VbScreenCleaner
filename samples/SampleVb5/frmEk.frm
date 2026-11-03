VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmEk 
   Caption         =   "Ek Senaryolar"
   ClientHeight    =   5000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5000
   ScaleWidth      =   6000
   Begin VB.CommandButton cmdShow 
      Caption         =   "Goster"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtSelHide 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblIfBlock 
      Caption         =   "IfBlock"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdVarShow 
      Caption         =   "VarShow"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOtherShow 
      Caption         =   "OtherShow"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdDyn 
      Caption         =   "Dyn"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtBuddy 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ComCtl2.UpDown updBuddy 
      Height          =   285
      Left            =   1335
      TabIndex        =   7
      Top             =   1080
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   327681
      BuddyControl    =   "txtBuddy"
      BuddyDispid     =   196615
      OrigLeft        =   1440
      OrigTop         =   1080
      OrigRight       =   1680
      OrigBottom      =   1365
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtDefault 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstHidden 
      Height          =   645
      Left            =   3120
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picHidden 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   2475
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdInPic 
         Caption         =   "InPic"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdViaVar 
      Caption         =   "ViaVar"
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblAfterIfBlock 
      Caption         =   "AfterIf"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetFocus 
      Caption         =   "SetFocus"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMultiIf 
      Caption         =   "MultiIf"
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSingleIf 
      Caption         =   "SingleIf"
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmEk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Select Case gMod
        Case 1: txtSelHide.Visible = False
    End Select
    If gAdmin Then
        lblIfBlock.Visible = False
    End If
    lblAfterIfBlock.Visible = False
    cmdVarShow.Visible = gAdmin
    Load cmdDyn(1)
    txtDefault = "varsayilan"
    lstHidden.Clear
    lstHidden.AddItem "a"
    cmdInPic.Visible = True
    cmdSetFocus.SetFocus
    If gAdmin Then cmdMultiIf.Caption = "a": cmdMultiIf.Tag = "b"
    If gAdmin Then cmdSingleIf.Caption = "a"
End Sub

Private Sub cmdShow_Click()
    cmdOtherShow.Visible = True
End Sub

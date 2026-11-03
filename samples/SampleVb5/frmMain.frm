VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Ana Ekran"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   8000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVisible 
      Caption         =   "Kaydet"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdHiddenUnused 
      Caption         =   "Eski"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOnlyInHandler 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdShownInLoad 
      Caption         =   "Yetkili"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblHideOnLoad 
      Caption         =   "Eski etiket"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblCondHide 
      Caption         =   "Kosullu"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame fraHidden 
      Caption         =   "Gizli grup"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txtA 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtB 
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   8.25
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame fraHidden2 
      Caption         =   "Veri"
      Height          =   1215
      Left            =   2760
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblX 
         Caption         =   "x"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox txtHiddenWrite 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdArr 
      Caption         =   "A0"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdArr 
      Caption         =   "A1"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   14
      Top             =   2800
      Width           =   615
   End
   Begin VB.CommandButton cmdArrH 
      Caption         =   "H0"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   15
      Top             =   2800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdArrH 
      Caption         =   "H1"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   16
      Top             =   2800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdCalled 
      Caption         =   "Called"
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   2800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtStr 
      Height          =   285
      Left            =   3720
      TabIndex        =   18
      Top             =   2800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chkEvt 
      Caption         =   "Evt"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3300
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdExternal 
      Caption         =   "Ext"
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      Top             =   3300
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdExtWrite 
      Caption         =   "ExtW"
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   3300
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdWith 
      Caption         =   "With"
      Height          =   375
      Left            =   2880
      TabIndex        =   22
      Top             =   3300
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Kapat"
      Height          =   375
      Left            =   3600
      TabIndex        =   23
      Top             =   3300
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblAfterExit 
      Caption         =   "AfterExit"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3800
      Width           =   1215
   End
   Begin VB.Label lblContinued 
      Caption         =   "Continued"
      Height          =   255
      Left            =   1440
      TabIndex        =   25
      Top             =   3800
      Width           =   1215
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   5000
      Top             =   3800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327682
   End
   Begin ComctlLib.ListView lvHidden 
      Height          =   1000
      Left            =   5000
      TabIndex        =   26
      Top             =   4300
      Visible         =   0   'False
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   1773
      Icons           =   "imgList"
      _Version        =   327682
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Dosya"
      Begin VB.Menu mnuExit 
         Caption         =   "&Cikis"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopA 
         Caption         =   "A"
      End
      Begin VB.Menu mnuPopB 
         Caption         =   "B"
      End
   End
   Begin VB.Menu mnuOld 
      Caption         =   "Eski Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuOldA 
         Caption         =   "Eski A"
      End
   End
   Begin VB.Menu mnuHiddenShortcut 
      Caption         =   "Gizli Kisayol"
      Visible         =   0   'False
      Begin VB.Menu mnuSave 
         Caption         =   "Kaydet"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Yonetim"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim s As String
    cmdShownInLoad.Visible = True
    lblHideOnLoad.Visible = False ' artik kullanilmiyor
    If gAdmin Then lblCondHide.Visible = False
    txtHiddenWrite.Text = "abc": Debug.Print "yukleniyor"
    txtHiddenWrite.Text = "a" & _
        "b"
    Me.lblContinued.Visible = False
    s = txtData.Text
    mnuAdmin.Visible = False
    Debug.Print "x = ""txtStr"""
    s = "txtStr"
    chkEvt.Value = 1
    With cmdWith
        .Caption = "w"
    End With
    If Not gAdmin Then Exit Sub
    lblAfterExit.Visible = False
End Sub

Private Sub cmdHiddenUnused_Click()
    txtOnlyInHandler.Text = "x"
    MsgBox txtOnlyInHandler.Text
End Sub

Private Sub cmdVisible_Click()
    cmdCalled_Click
    PopupMenu mnuPopup
End Sub

Private Sub cmdCalled_Click()
    MsgBox "called"
End Sub

Private Sub chkEvt_Click()
    MsgBox "degisti"
End Sub

Private Sub mnuOldA_Click()
    MsgBox "eski"
End Sub

Private Sub mnuPopA_Click()
    MsgBox "A"
End Sub

Private Sub cmdArrH_Click(Index As Integer)
    MsgBox Index
End Sub

Private Sub cmdArr_Click(Index As Integer)
    cmdArr(0).Caption = "x"
End Sub

Private Sub mnuExit_Click()
ErrH: lblHideOnLoad.Caption = "son"
    Unload Me
End Sub

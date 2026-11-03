VERSION 5.00
Begin VB.Form frmZor 
   Caption         =   "Zor Senaryolar"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   Begin VB.CommandButton cmdGo 
      Caption         =   "Calistir"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtV 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtThenColon 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblFunc 
      Caption         =   "Func"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblFuncIf 
      Caption         =   "FuncIf"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblBuiltin 
      Caption         =   "Builtin"
      Height          =   255
      Left            =   6720
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSelf 
      Height          =   285
      Left            =   8040
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResize 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data datGizli 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdClickOnly 
      Caption         =   "ClickOnly"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOneLiner 
      Caption         =   "OneLiner"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMultiLineHeader 
      Caption         =   "MultiLine"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCommented 
      Height          =   285
      Left            =   6720
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtInString 
      Height          =   285
      Left            =   8040
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtColonString 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtLabelComment 
      Height          =   285
      Left            =   1440
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtNumLabel 
      Height          =   285
      Left            =   2760
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraL1 
      Caption         =   "L1"
      Height          =   2655
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.PictureBox picL2 
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2115
         ScaleWidth      =   3435
         TabIndex        =   17
         Top             =   240
         Width           =   3495
         Begin VB.Frame fraL3 
            Caption         =   "L3"
            Height          =   1695
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   3135
            Begin VB.PictureBox picL4 
               Height          =   735
               Left            =   120
               ScaleHeight     =   675
               ScaleWidth      =   2835
               TabIndex        =   19
               Top             =   240
               Width           =   2895
               Begin VB.TextBox txtL5 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   20
                  Top             =   120
                  Width           =   1215
               End
            End
            Begin VB.Label lblL5b 
               Caption         =   "L5b"
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   1200
               Width           =   1215
            End
         End
      End
   End
   Begin VB.Frame fraMixed 
      Caption         =   "Mixed"
      Height          =   2655
      Left            =   3960
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox txtMixedRead 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame fraInner 
         Caption         =   "Inner"
         Height          =   1215
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   3135
         Begin VB.TextBox txtInnerUnused 
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmdGap 
      Caption         =   "Gap0"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdGap 
      Caption         =   "Gap3"
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   27
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdGap 
      Caption         =   "Gap7"
      Height          =   375
      Index           =   7
      Left            =   1560
      TabIndex        =   28
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdCallKw 
      Caption         =   "CallKw"
      Height          =   375
      Left            =   2280
      TabIndex        =   29
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdValueTrue 
      Caption         =   "ValueTrue"
      Height          =   375
      Left            =   3600
      TabIndex        =   30
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraHost 
      Caption         =   "Host"
      Height          =   735
      Left            =   4920
      TabIndex        =   31
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtTypeName 
      Height          =   285
      Left            =   120
      TabIndex        =   32
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtForms 
      Height          =   285
      Left            =   1440
      TabIndex        =   33
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtScreen 
      Height          =   285
      Left            =   2760
      TabIndex        =   34
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCaseVariants 
      Caption         =   "Case"
      Height          =   255
      Left            =   4080
      TabIndex        =   35
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtWithMe 
      Height          =   285
      Left            =   5400
      TabIndex        =   36
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblHideLoopLoad 
      Caption         =   "Loop"
      Height          =   255
      Left            =   6720
      TabIndex        =   37
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lblHideIfdef 
      Caption         =   "Ifdef"
      Height          =   255
      Left            =   8040
      TabIndex        =   38
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lblHideAfterResumeNext 
      Caption         =   "ResumeNext"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblHideWithVbFalse 
      Caption         =   "vbFalse"
      Height          =   255
      Left            =   1440
      TabIndex        =   40
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtDefaultRead 
      Height          =   285
      Left            =   2760
      TabIndex        =   41
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPrint 
      Height          =   285
      Left            =   4080
      TabIndex        =   42
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSelectCase 
      Height          =   285
      Left            =   5400
      TabIndex        =   43
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblModWrite 
      Caption         =   "ModWrite"
      Height          =   255
      Left            =   6720
      TabIndex        =   44
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblModBang 
      Caption         =   "ModBang"
      Height          =   255
      Left            =   8040
      TabIndex        =   45
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtClsRead 
      Height          =   285
      Left            =   120
      TabIndex        =   46
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAyni 
      Height          =   285
      Left            =   1440
      TabIndex        =   47
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDeadSub 
      Height          =   285
      Left            =   2760
      TabIndex        =   48
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOluZincir 
      Height          =   285
      Left            =   4080
      TabIndex        =   49
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCallByName 
      Height          =   285
      Left            =   5400
      TabIndex        =   50
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdModGoster 
      Caption         =   "ModGoster"
      Height          =   375
      Left            =   6720
      TabIndex        =   51
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtModOlu 
      Height          =   285
      Left            =   120
      TabIndex        =   53
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblHideParen 
      Caption         =   "Paren"
      Height          =   255
      Left            =   8040
      TabIndex        =   52
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Menu mnuDosya 
      Caption         =   "&Dosya"
      Begin VB.Menu mnuAc 
         Caption         =   "&Ac"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "Son"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuGizliUst 
      Caption         =   "Gizli"
      Visible         =   0   'False
      Begin VB.Menu mnuGizliAlt 
         Caption         =   "Alt"
         Begin VB.Menu mnuGizliAltAlt 
            Caption         =   "AltAlt"
         End
      End
   End
End
Attribute VB_Name = "frmZor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private s As String
Private i As Integer

Private Sub Form_Load()
    On Error Resume Next
    lblHideAfterResumeNext.Visible = False
    If gA Then: txtThenColon.Tag = "a"
    lblFunc.Caption = SayacArttir()
    If KayitVar() Then lblFuncIf.Caption = "x"
    lblBuiltin.Caption = Format$(Now, "hh:nn") & Left$("abc", 1)
    txtSelf.Text = txtSelf.Text & "a"
    ' s = txtCommented.Text
    Rem txtCommented.Visible = True
    MsgBox "txtInString: degeri"
    MsgBox "a: b": txtColonString.Tag = "x"
Hata: txtLabelComment.Tag = "" ' temizle
10 txtNumLabel.Tag = ""
    lblL5b.Visible = True
    s = txtMixedRead.Text
    Load mnuRecent(1)
    Call cmdCallKw_Click
    cmdValueTrue.Value = True
    Set txtV.Container = fraHost
    Debug.Print TypeName(txtTypeName)
    LBLCASEVARIANTS.Caption = "x"
    With Me
        .txtWithMe.Tag = "x"
    End With
    For i = 0 To 0: lblHideLoopLoad.Visible = False: Next
#If DEBUGMOD Then
    lblHideIfdef.Visible = False
#End If
    lblHideWithVbFalse.Visible = vbFalse
    lblHideParen.Visible = (False)
    CallByName Me, "DinamikCagrilan", VbMethod
    s = txtDefaultRead
    Print #1, txtPrint.Text
    Select Case txtSelectCase.Text
        Case "a": Beep
    End Select
End Sub

Private Sub cmdGo_Click()
    MsgBox "go"
End Sub

Private Sub picResize_Resize()
    Debug.Print "boyut"
End Sub

Private Sub datGizli_Reposition()
    Debug.Print "kayit"
End Sub

Private Sub cmdClickOnly_Click()
    MsgBox "tek"
End Sub

Private Sub cmdOneLiner_Click(): MsgBox "x": End Sub

Private Sub cmdMultiLineHeader_MouseDown(Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
    Debug.Print X
End Sub

Private Sub cmdGap_Click(Index As Integer)
    MsgBox Index
End Sub

Private Sub mnuGizliAltAlt_Click()
    MsgBox "alt"
End Sub

Private Sub cmdCallKw_Click()
    MsgBox "cagrildi"
End Sub

Private Sub cmdValueTrue_Click()
    MsgBox "deger"
End Sub

Private Sub HicCagrilmayan()
    s = txtDeadSub.Text
End Sub

Private Sub OluA()
    OluB
End Sub

Private Sub OluB()
    s = txtOluZincir.Text
End Sub

Public Sub DinamikCagrilan()
    s = txtCallByName.Text
End Sub

Private Sub OluAmaZararsiz()
    Beep
End Sub

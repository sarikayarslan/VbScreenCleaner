Attribute VB_Name = "modZor"
Option Explicit

Public gA As Boolean
Private mSayac As Long

Public Function SayacArttir() As String
    mSayac = mSayac + 1
    SayacArttir = CStr(mSayac)
End Function

Public Function KayitVar() As Boolean
    KayitVar = True
End Function

Public Sub ModulYaz()
    frmZor.lblModWrite.Caption = "x"
    frmZor!lblModBang.Caption = "y"
    Forms(0).txtForms.Tag = ""
    Screen.ActiveForm.txtScreen.Tag = ""
    frmZor.cmdModGoster.Visible = True
End Sub

Public Sub ModulOlu()
    Dim s As String
    s = frmZor.txtModOlu.Text
End Sub

Public Sub HedefiGoster()
    Dim c As Control
    For Each c In frmHedef.Controls
        c.Visible = True
    Next
End Sub

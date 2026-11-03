Attribute VB_Name = "Module1"
Option Explicit

Public gAdmin As Boolean

Public Sub Ayarla()
    frmMain.cmdExternal.Visible = True
    frmMain.cmdExtWrite.Caption = "yeni"
    frmOther.Show
End Sub

Public Sub Dinamik(f As Form)
    f.cmdViaVar.Caption = "x"
End Sub

VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ClearRecent"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StepbyStep As Boolean
Dim RegNameTxt, RegVal As String
Dim RegName As Integer

Private Sub Form_Load()
    If MsgBox("Are you sure you want to clear the recent listing for Visual Basic?", vbYesNo + vbQuestion, "ClearRecent") = vbNo Then End
    
    R = MsgBox("Click Yes to delete all recent listings. Click No to delete only specific listings. Click Cancel to cancel.", vbYesNoCancel + vbQuestion, "ClearRecent")
    If R = vbCancel Then End
    If R = vbNo Then StepbyStep = True
    If R = vbYes Then StepbyStep = False
    
    Reg.OpenRegistry HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles"
        Do
            RegName = RegName + 1
            RegNameTxt = Str(RegName)
            RegNameTxt = Right(RegNameTxt, Len(RegNameTxt) - 1)
            RegVal = Reg.GetValue(RegNameTxt)
            If GetValueOK = False Then Exit Do
            If StepbyStep = True Then
                R = MsgBox("Do you want to delete listing #" + Str(RegName) + ": " + RegVal + "?", vbYesNo + vbQuestion, "ClearRecent")
            Else
                R = vbYes
            End If
            If R = vbYes Then
                Reg.DeleteValue RegNameTxt
            End If
        Loop
    Reg.CloseRegistry
    End
End Sub

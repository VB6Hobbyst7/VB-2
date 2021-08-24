VERSION 5.00
Begin VB.Form frmLisQC 
   BackColor       =   &H00EAE7E3&
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   10980
   WindowState     =   2  '√÷¥Î»≠
End
Attribute VB_Name = "frmLisQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objQCForm As clsLisQCForm
Attribute objQCForm.VB_VarHelpID = -1
Private mvarButtonKey As String

Public Property Let ButtonKey(ByVal vData As String)
    mvarButtonKey = vData
End Property

Public Sub ShowThisForm()
    Call objQCForm.QCButtonClick(mvarButtonKey, frmLisQC)  'picForm)
End Sub

Private Sub Form_Activate()
    Me.WindowState = 2
'    medMain.lblSubMenu.Caption = Me.Caption
End Sub


Private Sub Form_Load()

    Set objQCForm = New clsLisQCForm
    
    objQCForm.EmpId = ObjMyUser.EmpId
    objQCForm.IsDeveloper = ObjMyUser.IsDeveloper

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objQCForm = Nothing
End Sub

Private Sub objQCForm_LastFormUnload()
    Unload Me
    Set frmLisQC = Nothing
End Sub

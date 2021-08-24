VERSION 5.00
Begin VB.Form frmLisStatistic 
   BackColor       =   &H00FEF7FF&
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
Attribute VB_Name = "frmLisStatistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objStatisticForm As clsLisStatisticForm
Attribute objStatisticForm.VB_VarHelpID = -1
Private mvarButtonKey As String

Public Property Let ButtonKey(ByVal vData As String)
    mvarButtonKey = vData
End Property

Public Sub ShowThisForm()
    Call objStatisticForm.StatisticButtonClick(mvarButtonKey, frmLisStatistic)  'picForm)
End Sub

Private Sub Form_Activate()
    Me.WindowState = 2
'    medMain.lblSubMenu.Caption = Me.Caption
End Sub


Private Sub Form_Load()

    Set objStatisticForm = New clsLisStatisticForm
    
    objStatisticForm.EmpId = ObjMyUser.EmpId
    objStatisticForm.IsDeveloper = ObjMyUser.IsDeveloper

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objStatisticForm = Nothing
End Sub

Private Sub objStatisticForm_LastFormUnload()
    Unload Me
    Set frmLisStatistic = Nothing
End Sub

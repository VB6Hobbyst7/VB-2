VERSION 5.00
Begin VB.Form frmLisCollection 
   BackColor       =   &H00F2FBFB&
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
Attribute VB_Name = "frmLisCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objCollectionForm As clsLisCollectForm
Attribute objCollectionForm.VB_VarHelpID = -1
Private mvarButtonKey As String

Public Property Let ButtonKey(ByVal vData As String)
    mvarButtonKey = vData
End Property

Public Sub ShowThisForm()
    Call objCollectionForm.CollectionButtonClick(mvarButtonKey, frmLisCollection)  'picForm)
End Sub

Private Sub Form_Activate()
    Me.WindowState = 2
'    medMain.lblSubMenu.Caption = Me.Caption
End Sub


Private Sub Form_Load()

    Set objCollectionForm = New clsLisCollectForm
    
    objCollectionForm.EmpId = ObjMyUser.EmpId
    objCollectionForm.IsDeveloper = ObjMyUser.IsDeveloper

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objCollectionForm = Nothing
End Sub

Private Sub objCollectionForm_LastFormUnload()
    Unload Me
    Set frmLisCollection = Nothing
End Sub

Public Sub LoadOutCollection(ByVal PtId As String, ByVal ordDt As String)
    
    objCollectionForm.LoadOutCollection PtId, ordDt
End Sub


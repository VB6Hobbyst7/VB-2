VERSION 5.00
Begin VB.Form frmLisReview 
   BackColor       =   &H00E8EEEE&
   Caption         =   "�����ȸ"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   10980
   WindowState     =   2  '�ִ�ȭ
End
Attribute VB_Name = "frmLisReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 2009.01.09 �缺�� ȯ��ID �Ķ���� �߰�
Private WithEvents objReviewForm As clsLisReviewForm
Attribute objReviewForm.VB_VarHelpID = -1
Private mvarButtonKey As String
Private mvarPtid As String
Public Property Let PtId(ByVal vData As String)
    mvarPtid = vData
End Property

Public Property Let ButtonKey(ByVal vData As String)
    mvarButtonKey = vData
End Property

Public Sub ShowThisForm()
'    Call objReviewForm.ReviewButtonClick(mvarButtonKey, frmLisReview)  'picForm)
    Call objReviewForm.ReviewButtonClick(mvarButtonKey, frmLisReview, mvarPtid) 'picForm)
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
    Me.WindowState = 2
End Sub


Private Sub Form_Load()

    Set objReviewForm = New clsLisReviewForm
    
    objReviewForm.EmpId = ObjMyUser.EmpId
    objReviewForm.IsDeveloper = ObjMyUser.IsDeveloper

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objReviewForm = Nothing
End Sub

Private Sub objReviewForm_LastFormUnload()
    Unload Me
    Set frmLisReview = Nothing
End Sub

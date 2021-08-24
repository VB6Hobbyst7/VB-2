VERSION 5.00
Begin VB.Form frmLisVerifyList 
   BackColor       =   &H00E8EEEE&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "결과보고 대기자 리스트"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmLisVerifyList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmLisVerifyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objReviewForm As clsLisReviewForm
Attribute objReviewForm.VB_VarHelpID = -1
Private mvarButtonKey As String

Public Property Let ButtonKey(ByVal vData As String)
    mvarButtonKey = vData
End Property

Public Sub ShowThisForm()
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
    Me.Top = 1700
    Me.Left = 11100
End Sub


Private Sub Form_Load()

    Set objReviewForm = New clsLisReviewForm
    
    objReviewForm.EmpId = ObjMyUser.EmpId
    objReviewForm.IsDeveloper = ObjMyUser.IsDeveloper
    Call objReviewForm.SetReviewForm(frmLisReview)
    Call objReviewForm.ReviewButtonClick("LIS505", frmLisVerifyList)  'picForm)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objReviewForm = Nothing
End Sub

Private Sub objReviewForm_LastFormUnload()
    Unload Me
    Set frmLisVerifyList = Nothing
End Sub

Private Sub objReviewForm_ListSELECTed()
    Me.Height = 1250
End Sub

Private Sub objReviewForm_MouseMove()
    Me.Height = 8500
End Sub

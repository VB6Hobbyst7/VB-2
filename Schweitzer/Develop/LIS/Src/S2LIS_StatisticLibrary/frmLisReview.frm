VERSION 5.00
Begin VB.Form frmLisReviewInStatisticLib 
   BackColor       =   &H00E8EEEE&
   Caption         =   "결과조회"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   9120
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmLisReviewInStatisticLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 2009.01.09 양성현 환자ID 파라메터 추가
Private WithEvents objReviewForm As clsLisReviewForm
Attribute objReviewForm.VB_VarHelpID = -1
Private mvarButtonKey As String
Private mvarPTid As String
Public Property Let PTid(ByVal vData As String)
    mvarPTid = vData
End Property

Public Property Let ButtonKey(ByVal vData As String)
    mvarButtonKey = vData
End Property

Public Sub ShowThisForm()
'    Load Me
    Set objReviewForm = New clsLisReviewForm
    
    objReviewForm.EmpId = ObjMyUser.EmpId
    objReviewForm.IsDeveloper = ObjMyUser.IsDeveloper

'    Call objReviewForm.ReviewButtonClick(mvarButtonKey, frmLisReview)  'picForm)
    Call objReviewForm.ReviewButtonClick(mvarButtonKey, frmLisReviewInStatisticLib, mvarPTid)  'picForm)
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
    Set frmLisReviewInStatisticLib = Nothing
End Sub

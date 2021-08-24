VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTextRst 
   BackColor       =   &H00C8D2CC&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   7665
   ClientLeft      =   6720
   ClientTop       =   6030
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H009AA4FC&
      Caption         =   "삭제"
      Height          =   360
      Left            =   5040
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   90
      Width           =   1005
   End
   Begin VB.ListBox lstTemplate 
      BackColor       =   &H00F8FFF7&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   45
      Style           =   1  '확인란
      TabIndex        =   4
      Top             =   480
      Width           =   9390
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "지움"
      Height          =   360
      Left            =   6315
      TabIndex        =   3
      Top             =   90
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소"
      Height          =   360
      Left            =   8415
      TabIndex        =   2
      Top             =   90
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인"
      Height          =   360
      Left            =   7365
      TabIndex        =   0
      Top             =   90
      Width           =   1005
   End
   Begin RichTextLib.RichTextBox txtTemplate 
      Height          =   3105
      Left            =   15
      TabIndex        =   5
      Top             =   4515
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   5477
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmTextRst.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "소  견"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Width           =   4320
   End
End
Attribute VB_Name = "frmTextRst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMyCtrl As Object
Private mDoctId As String
Private mTxtDiv As String

Public Property Get DoctId() As String
    DoctId = mDoctId
End Property
Public Property Let DoctId(ByVal vData As String)
    mDoctId = vData
End Property

Public Property Get Txtdiv() As String
    Txtdiv = mTxtDiv
End Property
Public Property Let Txtdiv(ByVal vData As String)
    mTxtDiv = vData
End Property

Public Property Get MyCtrl() As Object
    Set MyCtrl = mMyCtrl
End Property

Public Property Set MyCtrl(ByVal vNewValue As Object)
    Set mMyCtrl = vNewValue
End Property

Private Sub cmdCancel_Click()
    Unload Me
'    Set frmPreview = Nothing
End Sub

Private Sub cmdClear_Click()
    txtTemplate.Text = ""
End Sub

Private Sub cmdDelete_Click()
    Dim strKey  As String
    Dim strMsg  As String
    Dim i       As Integer
    
    strMsg = MsgBox("선택한 항목을 모두 삭제 하시겠습니까?", vbCritical + vbYesNo, "삭제")
    
    If strMsg = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrMsg
    
    DBConn.BeginTrans
    
    With lstTemplate
        For i = 0 To .ListIndex
            If .Selected(i) = True Then
                
                '-- 삭제
                Call Item_Delete(mDoctId, mTxtDiv, medGetP(.List(i), 1, " "))
                
            End If
        Next
    End With
    
    DBConn.CommitTrans
    
    With objDoctor
        Call .GetDoctInfo
    End With
    
    Call LoadTemplate
    
    Exit Sub
    
ErrMsg:
    MsgBox Err.Description
    DBConn.RollbackTrans
    
End Sub

Private Sub Item_Delete(ByVal pDoctID As String, ByVal pTxtDiv As String, _
                        ByVal pTxtCd As String)
    Dim strSql As String
    
    strSql = " delete " & T_LAB506 & _
             "  where doctid = " & DBS(pDoctID) & _
             "    and txtdiv = " & DBS(pTxtDiv) & _
             "    and txtcd  = " & DBS(pTxtCd)
             
    Call DBConn.Execute(strSql)
    
End Sub

Private Sub cmdOK_Click()
    MyCtrl.Text = txtTemplate.Text
    Unload Me
    Set frmTextRst = Nothing
End Sub

Public Sub LoadTemplate()

    Dim i As Integer
    
    lstTemplate.Clear
    For i = 1 To objDoctor.CmtCount
        If objDoctor.txtCmt(i).Txtdiv = mTxtDiv Then
            lstTemplate.AddItem objDoctor.txtCmt(i).Txtcd & _
                                Space(7 - Len(objDoctor.txtCmt(i).Txtcd)) & _
                                objDoctor.txtCmt(i).Txtnm
        End If
    Next
    
End Sub

Private Sub Form_Activate()
    Me.Left = 4845

End Sub

Private Sub Form_Load()
    Me.Left = 4845
    
    cmdDelete.Visible = True
    
End Sub

Private Sub lstTemplate_ItemCheck(Item As Integer)
    
    Dim strTmp As String
    Dim FindPos As Integer
    Dim strKey As String
    
    strKey = mDoctId & mTxtDiv & medGetP(lstTemplate.List(Item), 1, " ")
    If lstTemplate.Selected(Item) Then
        txtTemplate.Text = txtTemplate.Text & objDoctor.txtCmt(strKey).Txtrst & vbNewLine
    Else
        strTmp = objDoctor.txtCmt(strKey).Txtrst
        FindPos = txtTemplate.Find(strTmp, , , rtfWholeWord)
        If FindPos <> -1 Then
            txtTemplate.Text = Mid(txtTemplate.Text, 1, FindPos) & _
                               Mid(txtTemplate.Text, FindPos + Len(strTmp) + 3)
        End If
    End If
    
End Sub

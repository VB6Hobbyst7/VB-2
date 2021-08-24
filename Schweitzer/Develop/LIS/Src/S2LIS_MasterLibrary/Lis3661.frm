VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frm3661SpeTemp 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "기타검사 세부결과 Template등록"
   ClientHeight    =   9120
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10935
   Icon            =   "Lis3661.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   600
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   10770
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "기타검사 세부결과 Template등록"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2865
         TabIndex        =   18
         Top             =   180
         Width           =   4530
      End
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00EAE7E3&
      Caption         =   "삭제(&D)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   16
      Top             =   8190
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   600
      Left            =   60
      TabIndex        =   10
      Top             =   690
      Width           =   10710
      Begin VB.TextBox txtGnm 
         BackColor       =   &H00F1F5F4&
         ForeColor       =   &H00734A60&
         Height          =   270
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   2625
      End
      Begin VB.TextBox txtTmpResultNm 
         BackColor       =   &H00F1F5F4&
         ForeColor       =   &H00734A60&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   195
         Width           =   2625
      End
      Begin VB.Label Label6 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사항목그룹명 : "
         Height          =   180
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   1440
      End
      Begin VB.Label Label4 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Template명 :"
         Height          =   180
         Left            =   4290
         TabIndex        =   13
         Top             =   255
         Width           =   1095
      End
   End
   Begin VB.TextBox txtRstCode 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4185
      TabIndex        =   8
      Top             =   1425
      Width           =   1995
   End
   Begin VB.ListBox lstSKey2 
      BackColor       =   &H00F7FFF7&
      Height          =   4200
      Left            =   30
      TabIndex        =   3
      Top             =   3975
      Width           =   2955
   End
   Begin VB.ListBox lstSKey1 
      BackColor       =   &H00F7FFFF&
      Height          =   1860
      Left            =   30
      TabIndex        =   2
      Top             =   1710
      Width           =   2940
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00EAE7E3&
      Caption         =   "취소(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00EAE7E3&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   9525
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   8190
      Width           =   1320
   End
   Begin RichTextLib.RichTextBox txtTmpData 
      Height          =   6000
      Left            =   3030
      TabIndex        =   6
      Top             =   2130
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   10583
      _Version        =   393217
      BackColor       =   15857140
      ScrollBars      =   2
      TextRTF         =   $"Lis3661.frx":038A
   End
   Begin VB.Label lblRstField 
      BackStyle       =   0  '투명
      Caption         =   "Label3"
      ForeColor       =   &H00734A60&
      Height          =   240
      Left            =   1545
      TabIndex        =   15
      Top             =   1425
      Width           =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "◈ 결과코드 :"
      Height          =   180
      Left            =   3060
      TabIndex        =   9
      Top             =   1470
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "Template Text"
      Height          =   225
      Left            =   3150
      TabIndex        =   7
      Top             =   1860
      Width           =   2760
   End
   Begin VB.Label lblSIndx 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "결과 코드"
      Height          =   225
      Left            =   135
      TabIndex        =   5
      Top             =   3705
      Width           =   2760
   End
   Begin VB.Label lblFIndx 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "결과입력 Field : "
      Height          =   180
      Left            =   135
      TabIndex        =   4
      Top             =   1425
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Index           =   1
      Left            =   30
      Shape           =   4  '둥근 사각형
      Top             =   1320
      Width           =   2940
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00EFFFEE&
      FillStyle       =   0  '단색
      Height          =   375
      Index           =   0
      Left            =   30
      Shape           =   4  '둥근 사각형
      Top             =   3585
      Width           =   2940
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00F1F5F4&
      FillStyle       =   0  '단색
      Height          =   375
      Index           =   2
      Left            =   3030
      Shape           =   4  '둥근 사각형
      Top             =   1740
      Width           =   7680
   End
End
Attribute VB_Name = "frm3661SpeTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objComSql As New clsLISSqlStatement
Private blnInitFg As Boolean

Private mvarStCd As String
Private mvarTpCd As String

Public Property Let StCd(ByVal vData As String)
    mvarStCd = vData
End Property
Public Property Let TpCd(ByVal vData As String)
    mvarTpCd = vData
End Property

Private Sub CancelButton_Click()
    Unload Me
    Set frm3661SpeTemp = Nothing
End Sub

Private Sub cmdDelete_Click()
    
    Dim strSql1 As String
    
    strSql1 = objComSql.SqlDeleteLAB031(LC2_SpeAddTemp, lblRstField.Caption, txtRstCode.Text)
    dbconn.BeginTrans

On Error GoTo Err_Trap

    dbconn.Execute strSql1
    dbconn.CommitTrans
    
    Call lstSKey1_Click
    
    Exit Sub
    
Err_Trap:
    dbconn.RollbackTrans
    MsgBox "전산실 혹은 임상병리과로 연락바랍니다" & vbCrLf & Err.Description, vbCritical, "오류발생"

End Sub

Private Sub Form_Activate()
    If blnInitFg Then Exit Sub
    Call LoadRstFields
    If lstSKey1.ListCount > 0 Then lstSKey1.ListIndex = 0
    blnInitFg = True
End Sub

Private Sub Form_Load()
    blnInitFg = False
End Sub


Private Sub LoadRstFields()
    
    Dim objRs As Recordset
    Dim iCnt As Long
    
    lstSKey1.Clear
    
    Set objRs = New Recordset
    objRs.Open objComSql.SqlLAB031CodeList(LC2_SpeTemp, "*", mvarStCd, mvarTpCd), dbconn
    With objRs
        If Not .EOF Then
            For iCnt = 1 To Val(.Fields("field1").Value)
                lstSKey1.AddItem medGetP("" & .Fields("text1").Value, iCnt, vbTab)
            Next
        End If
    End With
    Set objRs = Nothing
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set objComSql = Nothing
End Sub

Private Sub lstSKey1_Click()
    
    Dim strFieldKey As String
    Dim objRs As Recordset
    Dim iCnt As Long
    
    If lstSKey1.ListIndex < 0 Then Exit Sub
    
    strFieldKey = lstSKey1.Text
    lblRstField.Caption = strFieldKey
    
    lstSKey2.Clear
    
    Set objRs = New Recordset
    objRs.Open objComSql.SqlLAB031CodeList(LC2_SpeAddTemp, "*", strFieldKey, , " order by cdval2"), dbconn
    With objRs
        While Not .EOF
            lstSKey2.AddItem "" & .Fields("cdval2").Value
            .MoveNext
        Wend
    End With
    Set objRs = Nothing
    
    If lstSKey2.ListCount > 0 Then
        lstSKey2.ListIndex = 0
    Else
        txtRstCode.Text = ""
        txtTmpData.Text = ""
    End If

End Sub

Private Sub lstSKey2_Click()

    Dim objRs As Recordset
    Dim iCnt As Long
    
    If lstSKey2.ListIndex < 0 Then Exit Sub
    
    txtRstCode.Text = lstSKey2.Text
    
    Set objRs = New Recordset
    objRs.Open objComSql.SqlLAB031CodeList(LC2_SpeAddTemp, "*", lblRstField.Caption, txtRstCode.Text), dbconn
    With objRs
        If Not .EOF Then
            txtTmpData.Text = "" & .Fields("text1").Value
        End If
    End With
    Set objRs = Nothing
    
End Sub

Private Sub OKButton_Click()
    
    Dim strSql1 As String
    Dim strSql2 As String
    
    strSql1 = objComSql.SqlDeleteLAB031(LC2_SpeAddTemp, lblRstField.Caption, txtRstCode.Text)
    If Trim(txtTmpData.Text) <> "" Then
        strSql2 = objComSql.SqlSaveLAB031(LC2_SpeAddTemp, lblRstField.Caption, txtRstCode.Text, "", "", _
                                          "", "", "", txtTmpData.Text, "", "1")
    Else
        strSql2 = ""
    End If
    dbconn.BeginTrans

On Error GoTo Err_Trap

    dbconn.Execute strSql1
    If strSql2 <> "" Then dbconn.Execute strSql2
    dbconn.CommitTrans
    
    Call lstSKey1_Click
    
    Exit Sub
    
Err_Trap:
    dbconn.RollbackTrans
    MsgBox "전산실 혹은 임상병리과로 연락바랍니다" & vbCrLf & Err.Description, vbCritical, "오류발생"
    
End Sub

Private Sub txtRstCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtTmpData.SetFocus
End Sub

Private Sub txtRstCode_LostFocus()
    
    Dim lngIndex As Long
    
    If Trim(txtRstCode.Text) = "" Then Exit Sub
    
    lngIndex = medListFind(lstSKey2, txtRstCode.Text)
    If lngIndex >= 0 Then
        lstSKey2.ListIndex = lngIndex
    Else
        lstSKey2.ListIndex = -1
        txtTmpData.Text = ""
    End If
    
End Sub

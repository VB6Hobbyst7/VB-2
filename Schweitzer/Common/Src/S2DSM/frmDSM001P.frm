VERSION 5.00
Begin VB.Form frmDSM001P 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "폼 등록"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   Icon            =   "frmDSM001P.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.OptionButton optDept 
      BackColor       =   &H00DBE6E6&
      Caption         =   "혈액은행"
      ForeColor       =   &H00DD6131&
      Height          =   180
      Index           =   2
      Left            =   5993
      TabIndex        =   9
      Top             =   360
      Width           =   1035
   End
   Begin VB.OptionButton optDept 
      BackColor       =   &H00DBE6E6&
      Caption         =   "해부병리"
      ForeColor       =   &H00DD6131&
      Height          =   180
      Index           =   1
      Left            =   4943
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtForm 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   998
      MaxLength       =   30
      TabIndex        =   8
      ToolTipText     =   "예)frmDsm001 폼명을 넣어 주세요."
      Top             =   300
      Width           =   2790
   End
   Begin VB.TextBox txtNm 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   983
      MaxLength       =   30
      TabIndex        =   7
      ToolTipText     =   "폼의 이름을 넣어 주세요."
      Top             =   1260
      Width           =   5805
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   968
      MaxLength       =   50
      TabIndex        =   6
      ToolTipText     =   "폼의 설명을 넣어주세요."
      Top             =   1800
      Width           =   5835
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "취소(&C)"
      Height          =   405
      Left            =   4673
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   2220
      Width           =   1050
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "확인(&O)"
      Height          =   405
      Left            =   3593
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   2220
      Width           =   1050
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00E0E0E0&
      Caption         =   "적용(&A)"
      Height          =   405
      Left            =   5753
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   2220
      Width           =   1050
   End
   Begin VB.CheckBox chkWrite 
      BackColor       =   &H00DBE6E6&
      Caption         =   " 쓰 기"
      ForeColor       =   &H00DD6131&
      Height          =   255
      Left            =   1973
      TabIndex        =   2
      Top             =   780
      Width           =   945
   End
   Begin VB.CheckBox chkPrint 
      BackColor       =   &H00DBE6E6&
      Caption         =   " 프 린 터"
      ForeColor       =   &H00DD6131&
      Height          =   255
      Left            =   2993
      TabIndex        =   1
      Top             =   780
      Width           =   1185
   End
   Begin VB.CheckBox chkRead 
      BackColor       =   &H00DBE6E6&
      Caption         =   " 읽 기"
      ForeColor       =   &H00DD6131&
      Height          =   255
      Left            =   998
      TabIndex        =   0
      Top             =   780
      Width           =   945
   End
   Begin VB.OptionButton optDept 
      BackColor       =   &H00DBE6E6&
      Caption         =   "임상병리"
      ForeColor       =   &H00DD6131&
      Height          =   180
      Index           =   0
      Left            =   3923
      TabIndex        =   11
      Top             =   360
      Value           =   -1  'True
      Width           =   1080
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "폼   : "
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   233
      TabIndex        =   15
      Tag             =   "103"
      Top             =   360
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "기능 :"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   233
      TabIndex        =   14
      Tag             =   "103"
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "이름 :"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   233
      TabIndex        =   13
      Tag             =   "103"
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "설명 :"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   233
      TabIndex        =   12
      Tag             =   "103"
      Top             =   1860
      Width           =   540
   End
End
Attribute VB_Name = "frmDSM001P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Form명   : frmDSM001P
'|  2. 기  능   : 폼 등록,수정
'|  3. 작성자   : 김 동열
'|  4. 작성일   : 2000.10.23
'|
'|  CopyRight(C) 2000 대련엠티에스
'+--------------------------------------------------------------------------------------+
Option Explicit


Private objSQL As clsDSMSqlStmt

Private Sub cmdApply_Click()
    '폼 ID 체크....??있어야 할건 있어야지.....
    If txtForm.Text = "" Then
        MsgBox "폼 명을 입력하여 주세요.", vbInformation, Me.Caption
        Exit Sub
    End If
    If txtForm.Enabled = True Then
        '등록 시키자
        Insert_Save
        clear
    Else
    '수정 시키자
        Update_Save
    End If
End Sub

Private Sub cmdCancel_Click()
    '닫자..
    Unload Me
    '삭제 메뉴를 살리자..
    frmDSM001.mnuDelete.Enabled = True
End Sub

Private Sub cmdSave_Click()
    If txtForm.Enabled = True Then
        If txtForm.Text = "" Then
            clear
            Unload Me
            Exit Sub
        End If
    '등록 시키자
        Insert_Save
        clear
    Else
    '수정 시키자
        Update_Save
        clear
    End If
    Unload Me
    '삭제 메뉴를 살리자...
    frmDSM001.mnuDelete.Enabled = True
End Sub

Private Sub clear()
    '깨끗이.....
    txtForm.Text = ""
    txtNm.Text = ""
    txtDesc.Text = ""
    chkRead.Value = 0
    chkWrite.Value = 0
    chkPrint.Value = 0
End Sub

Private Sub Insert_Save()
    Dim Rs As New Recordset
    Dim strTmp As VbMsgBoxResult
    Dim strDept As String
    
    If optDept(0).Value = True Then
        strDept = "L"
    ElseIf optDept(1).Value = True Then
        strDept = "A"
    Else
        strDept = "B"
    End If
    
    Set objSQL = New clsDSMSqlStmt
    '등록되 있는 폼인지 확인
    
    Rs.Open objSQL.GetSQLCOM007(strDept, txtForm.Text), dbconn
    
    If Rs.EOF = False Then
        MsgBox "이미 존재하는 폼 입니다.다시 한 번 확인바랍니다.", vbInformation, Me.Caption
        Set Rs = Nothing
        Set objSQL = Nothing
        Exit Sub
    End If
    '저장여부 확인...
    strTmp = MsgBox("폼 명이 " & txtForm & "인 레벨을 저장하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
    If strTmp = vbOK Then
        '저장
        If objSQL.Set_COM007(True, strDept, Trim(txtForm.Text), Trim(txtNm.Text), Trim(txtDesc.Text), chkRead, chkWrite, chkPrint) = True Then
        '리스트 뷰에 등록
            ListView_Insert
            MsgBox "저장성공하였습니다.", vbInformation, Me.Caption
        End If
    End If
    Set Rs = Nothing
    Set objSQL = Nothing
End Sub

Private Sub Update_Save()
    Dim strTmp As VbMsgBoxResult
    Dim strDept As String
    
    If optDept(0).Value = True Then
        strDept = "L"
    ElseIf optDept(1).Value = True Then
        strDept = "A"
    Else
        strDept = "B"
    End If
    
    Set objSQL = New clsDSMSqlStmt
    '수정여부 확인...
    strTmp = MsgBox("폼 명이 " & txtForm & "인 레벨을 수정하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
    If strTmp = vbOK Then
        '수정
        If objSQL.Set_COM007(False, strDept, Trim(txtForm.Text), Trim(txtNm.Text), Trim(txtDesc.Text), chkRead, chkWrite, chkPrint) = True Then
            '리스트 뷰 수정
            Listview_Update
            MsgBox "수정하였습니다.", vbInformation
        End If
    End If
    Set objSQL = Nothing
End Sub

Private Sub ListView_Insert()
    '리스트 뷰에 추가
    Dim itmX As Object
    Dim strDeptFg As String
    Dim strRead As String
    Dim strWrite As String
    Dim strPrint As String
    
    If optDept(0).Value = True Then
        strDeptFg = "LIS"
    ElseIf optDept(1).Value = True Then
        strDeptFg = "APS"
    Else
        strDeptFg = "BBS"
    End If
    
    strRead = IIf(chkRead.Value = "0", "없음", "있음")
    strWrite = IIf(chkWrite.Value = "0", "없음", "있음")
    strPrint = IIf(chkPrint.Value = "0", "없음", "있음")
    '리스트 뷰에 추가 시키자.
    With frmDSM001.lvwForm
        Set itmX = .ListItems.Add(, , strDeptFg)
        itmX.SubItems(1) = txtForm.Text
        itmX.SubItems(2) = txtNm.Text
        itmX.SubItems(3) = txtDesc.Text
        itmX.SubItems(4) = strRead
        itmX.SubItems(5) = strWrite
        itmX.SubItems(6) = strPrint
    End With
End Sub

Private Sub Listview_Update()
    '리스트 뷰 수정
    Dim itmX As Object
    Dim strRead As String
    Dim strWrite As String
    Dim strPrint As String
    
    strRead = IIf(chkRead.Value = "0", "없음", "있음")
    strWrite = IIf(chkWrite.Value = "0", "없음", "있음")
    strPrint = IIf(chkPrint.Value = "0", "없음", "있음")
    '리스트 뷰 수정하자....
    With frmDSM001.lvwForm
        Set itmX = .ListItems(.SelectedItem.Index)
        itmX.SubItems(2) = txtNm.Text
        itmX.SubItems(3) = txtDesc.Text
        itmX.SubItems(4) = strRead
        itmX.SubItems(5) = strWrite
        itmX.SubItems(6) = strPrint
    End With
End Sub

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm342Common2 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10935
   ControlBox      =   0   'False
   Icon            =   "Lis344.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport crtReport 
      Left            =   4680
      Top             =   3885
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EBF3ED&
      Height          =   690
      Left            =   3105
      ScaleHeight     =   630
      ScaleWidth      =   7320
      TabIndex        =   25
      Top             =   7800
      Width           =   7380
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00DBE6E6&
         Caption         =   "출력(&P)"
         Height          =   510
         Left            =   675
         Style           =   1  '그래픽
         TabIndex        =   30
         Tag             =   "25612"
         Top             =   45
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00DBE6E6&
         Caption         =   "종료(&X)"
         Height          =   510
         Left            =   6000
         Style           =   1  '그래픽
         TabIndex        =   29
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00DBE6E6&
         Caption         =   "삭제(&D)"
         Height          =   510
         Left            =   3345
         Style           =   1  '그래픽
         TabIndex        =   28
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00DBE6E6&
         Caption         =   "저장(&S)"
         Height          =   510
         Left            =   2010
         Style           =   1  '그래픽
         TabIndex        =   27
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00DBE6E6&
         Caption         =   "화면지움(&C)"
         Height          =   510
         Left            =   4665
         Style           =   1  '그래픽
         TabIndex        =   26
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
   End
   Begin VB.ListBox lstSubKey 
      BackColor       =   &H00FEF5F3&
      Height          =   6720
      Left            =   240
      TabIndex        =   0
      Top             =   1185
      Width           =   2835
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   7050
      Left            =   3105
      TabIndex        =   10
      Top             =   720
      Width           =   7380
      Begin VB.CommandButton cmdPopup 
         BackColor       =   &H00DEDBDD&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3780
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "Lis344.frx":08CA
         Style           =   1  '그래픽
         TabIndex        =   35
         Top             =   2055
         Width           =   300
      End
      Begin VB.CommandButton cmdPopup 
         BackColor       =   &H00DEDBDD&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2475
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "Lis344.frx":0E54
         Style           =   1  '그래픽
         TabIndex        =   33
         Top             =   1455
         Width           =   300
      End
      Begin VB.CommandButton cmdPopup 
         BackColor       =   &H00DEDBDD&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2490
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "Lis344.frx":13DE
         Style           =   1  '그래픽
         TabIndex        =   31
         Top             =   585
         Width           =   300
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Index           =   4
         Left            =   270
         TabIndex        =   6
         Top             =   3675
         Width           =   6960
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Index           =   1
         Left            =   270
         TabIndex        =   3
         Top             =   2070
         Width           =   3525
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Index           =   3
         Left            =   270
         TabIndex        =   5
         Top             =   3135
         Width           =   6960
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Index           =   2
         Left            =   270
         TabIndex        =   4
         Top             =   2610
         Width           =   6960
      End
      Begin VB.CheckBox chkKeyLock 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Index Key Lock Mode"
         Height          =   240
         Left            =   5160
         TabIndex        =   9
         Top             =   645
         Value           =   1  '확인
         Width           =   2160
      End
      Begin VB.TextBox txtSubKey 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   270
         TabIndex        =   1
         Top             =   585
         Width           =   2220
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   1065
         Index           =   6
         Left            =   270
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   5910
         Width           =   6960
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   1440
         Index           =   5
         Left            =   270
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   4230
         Width           =   6960
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   1470
         Width           =   2205
      End
      Begin MedControls1.LisLabel lblCaption 
         Height          =   300
         Index           =   0
         Left            =   2820
         TabIndex        =   32
         Top             =   600
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   529
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCaption 
         Height          =   300
         Index           =   1
         Left            =   2790
         TabIndex        =   34
         Top             =   1470
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCaption 
         Height          =   300
         Index           =   2
         Left            =   4095
         TabIndex        =   36
         Top             =   2070
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Field5"
         Height          =   180
         Index           =   5
         Left            =   300
         TabIndex        =   18
         Top             =   3495
         Width           =   6915
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Field2"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   17
         Top             =   1845
         Width           =   3435
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Field4"
         Height          =   180
         Index           =   4
         Left            =   300
         TabIndex        =   16
         Top             =   2940
         Width           =   6915
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Field3"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   15
         Top             =   2415
         Width           =   6915
      End
      Begin VB.Label lblCap 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Index"
         Height          =   225
         Index           =   0
         Left            =   285
         TabIndex        =   14
         Top             =   300
         Width           =   3765
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000011&
         X1              =   105
         X2              =   7320
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   105
         X2              =   7320
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Text2"
         Height          =   180
         Index           =   7
         Left            =   300
         TabIndex        =   13
         Top             =   5715
         Width           =   6945
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Text1"
         Height          =   180
         Index           =   6
         Left            =   300
         TabIndex        =   12
         Top             =   4035
         Width           =   6945
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Field1"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   11
         Top             =   1230
         Width           =   3315
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "건이 조회되었습니다."
      Height          =   255
      Left            =   1260
      TabIndex        =   24
      Top             =   8100
      Width           =   1755
   End
   Begin VB.Label lblSubKeyCnt 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   8100
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "☞ 총 "
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   8100
      Width           =   435
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00EBEBEB&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   420
      Left            =   240
      Shape           =   4  '둥근 사각형
      Top             =   7980
      Width           =   2835
   End
   Begin VB.Label lblTable 
      BackColor       =   &H00DBE6E6&
      Height          =   195
      Left            =   3180
      TabIndex        =   21
      Top             =   8100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblRName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Laboratory Information System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00613636&
      Height          =   495
      Left            =   255
      TabIndex        =   19
      Top             =   285
      Width           =   4935
   End
   Begin VB.Label lblSubName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "세 부 항 목  코  드"
      Height          =   180
      Left            =   915
      TabIndex        =   20
      Top             =   870
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Index           =   0
      Left            =   240
      Shape           =   4  '둥근 사각형
      Top             =   780
      Width           =   2820
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00D191A2&
      BorderWidth     =   3
      FillColor       =   &H00F1F5F4&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   180
      Shape           =   4  '둥근 사각형
      Top             =   180
      Width           =   5115
   End
End
Attribute VB_Name = "frm342Common2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명 : frm342Common2.frm
'   작성자 :
'   내  용 : 공통코드마스터2폼
'   작성일 :
'   버  전 :
'       1. 5.0.4: 이상대(2005-01-04)
'          - LC3_WBCCode, LC3_NRBCCode추가
'-----------------------------------------------------------------------------'

Option Explicit

Const cRKCount = 13
Const cLableCount = 5

'Dim mvarRKey(cRKCount - 1) As String
Dim blnFirst                    As Boolean
Private ChangeFlag              As Boolean
Dim mvarRName(cRKCount - 1)     As String
Dim mvarLabel(cRKCount - 1)     As String
Dim mvatTableDiv                As String
Private mvarRKey                As String

Private objProbar               As clsProgress
Private objSql                  As clsLISSqlCodeMaster
Private WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1

Public Property Get Rkey() As String
    Rkey = mvarRKey
End Property

Public Property Let Rkey(ByVal vNewValue As String)
    mvarRKey = vNewValue
    
    Dim ii As Integer
    
    For ii = 0 To 2
        cmdPopup(ii).Visible = False: lblCaption(ii).Visible = False: lblCaption(ii).Caption = ""
    Next
    
    '## 5.0.4: 이상대(2005-01-04)
    '   - LC3_WBCCode, LC3_NRBCCode 추가
    Select Case mvarRKey
        Case LC3_HighItem, LC3_ByPass
            cmdPopup(0).Visible = True: cmdPopup(1).Visible = True
            lblCaption(0).Visible = True: lblCaption(1).Visible = True
        Case LC3_ICUTestCd, LC3_WBCDiffCode, LC3_WBCCode, LC3_NRBCCode, LC3_RESULTREADTEST
            cmdPopup(0).Visible = True
            lblCaption(0).Visible = True
        Case LC3_ReportTesctCd, LC3_POCTestCd
            cmdPopup(0).Visible = True
            lblCaption(0).Visible = True
            cmdPopup(2).Visible = True
            lblCaption(2).Visible = True
    End Select
End Property

Public Property Get RName() As String
    RName = lblRName
End Property

Public Property Let RName(ByVal vNewValue As String)
    
    lblTable = medGetP(vNewValue, 2, ":")
    lblRName = medGetP(vNewValue, 1, ":")
    
End Property

Private Sub LoadSubKey()

    Dim i As Integer, SSQL As String
    Dim dsSKey As Recordset
    
    Set objSql = New clsLISSqlCodeMaster
    With objSql
        Select Case lblTable
            Case T_COM003
                SSQL = .GetComCdMST2(Rkey)
            Case T_COM004
                SSQL = .GetComCdTemp(Rkey)
        End Select
    End With
        
    Set objProbar = New clsProgress
    
    objProbar.Container = MainFrm.stsbar
    objProbar.Message = "자료를 읽기 위해 준비중입니다..."
'    objProbar.Value = 1
    Set dsSKey = New Recordset
    dsSKey.Open SSQL, DBConn
    
    lstSubKey.Clear
    Call ClearScreen
    txtSubKey.Locked = False
    
    If dsSKey.EOF Then GoTo NoData
    
    objProbar.Max = dsSKey.RecordCount
    objProbar.Message = ""
    
    While (Not dsSKey.EOF)
        i = i + 1
        Select Case Rkey
            Case LC3_Specimen:
                    lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value & vbTab & dsSKey.Fields("field3").Value & ""
            Case lc3_workarea, LC3_StoreCd, LC3_PartCd, LC3_Buildings, LC3_DiffKeyMap, LC3_TUBERCLE, _
                 LC3_OutLab, LC3_TUBERCLE, LC3_HospCd, LC3_Section, _
                 LC3_Media, LC3_POCTimes, LC3_RoundTime, LC3_ColTeam, _
                 LC3_Species, LC3_MWSKinds, LC3_RefLab, LC3_Method, LC3_BldQcRst, LC3_PtDiv, LC3_WorkTime:
                    lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value & vbTab & dsSKey.Fields("field1").Value & ""
            Case LC3_ElectronicSign, _
                    LC3_Volume, LC3_StaticGroup, _
                    "01":
                    lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value & vbTab & dsSKey.Fields("field2").Value & ""
            Case LC3_Microbe, LC3_Vander, _
                 LC3_AntiBiotic, _
                 "02":
                    lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value & vbTab & dsSKey.Fields("text1").Value & ""
            Case LC3_HighItem, LC3_ByPass, LC3_ICUTestCd, LC3_POCTestCd, LC3_ReportTesctCd, _
                 LC3_WBCDiffCode, LC3_WBCCode, LC3_NRBCCode
                    lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value & vbTab & dsSKey.Fields("testnm").Value & ""
            Case LC3_BatchColDept
'                    ObjLISComCode.DeptCd.KeyChange dsSKey.Fields("cdval1").Value & ""
'                    lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value & "" & vbTab & ObjLISComCode.DeptCd.Fields("deptnm")
'                    Dim objDept As clsBasisData
                    
'                    Set objDept = New clsBasisData
                    
                    lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value & "" & vbTab & GetDeptNm(dsSKey.Fields("cdval1").Value & "")
'                    Set objDept = Nothing
            Case Else
                lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value & ""
            End Select
        dsSKey.MoveNext
        objProbar.Value = i
    Wend
    
    lblSubKeyCnt = lstSubKey.ListCount
    
NoData:
   Set dsSKey = Nothing
   Set objSql = Nothing
   Set objProbar = Nothing
   
End Sub

Private Sub chkKeyLock_Click()
    txtSubKey.Locked = chkKeyLock.Value
    If chkKeyLock.Value = 0 Then
      Call cmdClear_Click
      If blnFirst = True Then
         txtSubKey.SetFocus
      End If
    End If
End Sub



Private Sub cmdClear_Click()
    If Not ConfirmExit Then Exit Sub
    lstSubKey.ListIndex = -1
    ClearScreen
    txtSubKey.Locked = False
    If blnFirst = True Then
        Me.Visible = True
        DoEvents
        txtSubKey.SetFocus
    End If
    ChangeFlag = False
End Sub

Private Sub ClearScreen()
    
    Dim i As Integer
    
    lblSubKeyCnt = "0"
    If lstSubKey.ListCount > 0 Then
        lblSubKeyCnt = lstSubKey.ListCount
    End If

    txtSubKey = ""
    For i = 1 To txtVal.Count
        txtVal(i - 1) = ""
    Next i
    lblCaption(0).Caption = "": lblCaption(1).Caption = "": lblCaption(2).Caption = ""
End Sub

Private Sub cmdDelete_Click()
    
    Dim sMsg As String
    Dim sRes As Integer, sStyle As Integer

    If Trim(txtSubKey) = "" Then
        MsgBox "Index가 지정이 잘못 되었습니다. 확인 후 처리 바랍니다."
        Exit Sub
    End If

    sMsg = lblRName & " Table에서 (" & txtSubKey.Text & ") Key와 Data를 삭제합니다" & Chr$(13) & Chr$(10) & _
        "정말 삭제해도 좋습니까?"
    sStyle = vbYesNo + vbCritical + vbDefaultButton2

    sRes = MsgBox(sMsg, sStyle, "삭제 확인")
    If sRes = vbYes Then
        Call DeleteData
        ChangeFlag = False
    Else
        Exit Sub
    End If
    
    cmdClear_Click

End Sub

Private Sub DeleteData()
    
    Dim SSQL(1) As String

    Set objSql = New clsLISSqlCodeMaster
    
    Select Case lblTable
        Case T_COM003
            SSQL(0) = objSql.DelComCdMST2(Rkey, Trim(txtSubKey))
        Case T_COM004
            SSQL(0) = objSql.DelComCdTemp(Rkey, Trim(txtSubKey))
    End Select
    
    DBConn.BeginTrans
    
    If InsertData(SSQL, False) Then
        DBConn.CommitTrans
        
        Call LoadSubKey
    Else
        DBConn.RollbackTrans
        MsgBox Err.Description, vbExclamation
    End If
    
    Set objSql = Nothing
    
End Sub

Private Sub cmdExit_Click()
   
   Unload Me
   
End Sub

Private Sub cmdPrint_Click()
    Dim strTable As String
    Dim strCdIndex As String

    If lstSubKey.ListCount = 0 Then
        MsgBox "출력할 내역이 없습니다.", vbInformation, "정보확인"
        Exit Sub
    End If

    Select Case Mid(Trim(Me.Caption), 1, 2)
        Case "A3"
            strTable = "LC3"
        Case "A4"
            strTable = "LC4"
    End Select
    
    strCdIndex = Rkey
    
    'strCdIndex = Trim(Me.Caption)

    If GetPrintOpt(strTable, strCdIndex) = False Then
        MsgBox "출력하는 도중 오류가 발생하였습니다.", vbExclamation, "정보확인"
        Exit Sub
    End If
End Sub

Private Function GetPrintOpt(ByVal pCdindex As String, ByVal pCdVal As String) As Boolean

'pCdinDex  ex) LC3,LC4
'pCdVal ex) A301, A401

    Dim RS As Recordset
    Dim aryTmp() As String
    Dim strTmp As String
    Dim strTitle As String  '타이틀
    Dim strTemp As String
    Dim lngCnt As Long
    Dim lngHeadCnt As Long
    Dim strHeadPos As String
    Dim strHeadTmp As String
    Dim strPrtOpt As String
    Dim strHeadOpt As String    '헤더 옵션 ex) c1, f1
    Dim i As Long

    Set objSql = New clsLISSqlCodeMaster
    Set objProbar = New clsProgress
    With objProbar
'        Set .StatusBar = medMain.stsBar
'        .HAlign = hCenter
        .Message = "자료를 출력하기 위하여 준비중입니다..."
        .Value = 1
    End With

    With objSql
'        .setDbConn DbConn
'        Set rs = OpenRecordSet(.GetComCdIndex(pCdinDex, pCdVal))
        Set RS = New Recordset
        RS.Open .GetComCdIdxComCd123(pCdindex, pCdVal), DBConn
    End With

    If RS.EOF Then
        GetPrintOpt = True
        MsgBox "출력할 내역이 없습니다.", vbInformation, "정보확인"
        Set RS = Nothing
        Set objSql = Nothing
        Exit Function
    Else
'        strTitle = "" & rs.Fields("field1")   '타이틀
'        AryTmp = Split(Trim(Rs.Fields("text1")), ";")   '헤더갯수 및 헤더

        strTitle = "" & RS.Fields("title").Value   '타이틀
        aryTmp = Split("" & RS.Fields("header").Value, ";") '헤더
        lngCnt = UBound(aryTmp)

        For i = 0 To lngCnt
            strTmp = aryTmp(i)
            If strTmp <> "" Then
                Select Case pCdindex
                    Case "LC3"
                        Select Case i + 1
                            Case "1"
                                strHeadPos = "C1"
                            Case "2"
                                strHeadPos = "F1"
                            Case "3"
                                strHeadPos = "F2"
                            Case "4"
                                strHeadPos = "F3"
                            Case "5"
                                strHeadPos = "F4"
                            Case "6"
                                strHeadPos = "F5"
                            Case "7"
                                strHeadPos = "T1"
                            Case "8"
                                strHeadPos = "T2"
                        End Select
                    Case "LC4"
                        Select Case i + 1
                            Case "1"
                                strHeadPos = "C1"
                            Case "2"
                                strHeadPos = "F1"
                            Case "3"
                                strHeadPos = "F2"
                            Case "4"
                                strHeadPos = "T1"
                            Case "5"
                                strHeadPos = "T2"
                        End Select
                End Select

                strHeadTmp = strHeadTmp & strHeadPos
                lngHeadCnt = lngHeadCnt + 1                             '헤더갯수
                strTemp = strTemp & strTmp & COL_DIV
            End If
        Next i
        '타이틀, 헤더,헤더옵션, 헤더갯수
        strPrtOpt = strTitle & LINE_DIV & strTemp & LINE_DIV & strHeadTmp & LINE_DIV & lngHeadCnt & LINE_DIV

        If PrintOut(strPrtOpt, pCdindex, pCdVal) Then
            GetPrintOpt = True
        Else
            GetPrintOpt = False
        End If
    End If

    Set RS = Nothing
    Set objSql = Nothing
End Function

Private Function PrintOut(ByVal pPrtOpt As String, ByVal pTable As String, _
                          ByVal pCdindex As String) As Boolean

'pPrtOpt 타이틀, 헤더, 헤더위치, 헤더갯수
'pTable LC2, LC3, LC4
'pCdiDex Cdindex
'헤더의 위치번호로 field1, Text1 구분

    Dim i As Long
    Dim j As Long
    Dim strTmp As String
    Dim strFileNm As String
    Dim strRptNm As String
    Dim lngFNum As Long
    Dim strMyFile As String
    Dim strTitle As String
    Dim aryHead() As String
    Dim strHeadCnt As String
    Dim lngCnt As Long
    Dim strHeadOpt As String
    Dim strOpt As String
    Dim strHeadTmp As String

'파리미터로 넘기는거 타이들, 헤더, 병원명

    strTitle = medGetP(pPrtOpt, 1, LINE_DIV)    '파라미터로 넘길거

    strTmp = medGetP(pPrtOpt, 2, LINE_DIV)
    strHeadTmp = medGetP(pPrtOpt, 3, LINE_DIV)
    strHeadCnt = medGetP(pPrtOpt, 4, LINE_DIV)

    aryHead = Split(strTmp, COL_DIV)

    Set objSql = New clsLISSqlCodeMaster

    Select Case strHeadCnt
        Case "2"    'C1,F1,, C1,T1
            Select Case strHeadTmp
                Case "C1F1" 'LC3
                    strOpt = "1"

                    strTmp = GetRecord(objSql.GetComCdMST2(pCdindex), strOpt)
                Case "C1T1" 'LC4
                    strOpt = "2"

                    strTmp = GetRecord(objSql.GetComCdTemp(pCdindex), strOpt)
            End Select

'            strMyFile = Dir(App.Path & "\rpt\rptCOM004Others.rpt")
            strMyFile = Dir(App.Path & "\rpt\rptAPS014.rpt")

            If strMyFile = "" Then
                PrintOut = True
'                MsgBox "rptCOM004Others.rpt 파일이 없습니다.", vbCritical, "정보확인"
                MsgBox "rptAPS014.rpt 파일이 없습니다.", vbCritical, "정보확인"
                Exit Function
            End If

'            strRptNm = App.Path & "\rpt\rptCOM004Others.rpt"
            strRptNm = App.Path & "\rpt\rptAPS014.rpt"
        Case "3"    'C1,F1,T1,, C1,F1,F2
            Select Case strHeadTmp
                Case "C1F1T1"   'LC4
                    strOpt = "3"

                    strTmp = GetRecord(objSql.GetComCdTemp(pCdindex), strOpt)

        '            strMyFile = Dir(App.Path & "\rpt\rptCOM004CytoDx.rpt")
                    strMyFile = Dir(App.Path & "\rpt\rptAPS020.rpt")

                    If strMyFile = "" Then
                        PrintOut = True
        '                MsgBox "rptCOM004CytoDx.rpt 파일이 없습니다.", vbCritical, "정보확인"
                        MsgBox "rptAPS020.rpt 파일이 없습니다.", vbCritical, "정보확인"
                        Exit Function
                    End If

        '            strRptNm = App.Path & "\rpt\rptCOM004CytoDx.rpt"
                    strRptNm = App.Path & "\rpt\rptAPS020.rpt"
                Case "C1F1F2"   'LC3
                    strOpt = "4"

                    strTmp = GetRecord(objSql.GetComCdMST2(pCdindex), strOpt)

        '            strMyFile = Dir(App.Path & "\rpt\rptCOM004three.rpt")
                    strMyFile = Dir(App.Path & "\rpt\rptAPS015.rpt")

                    If strMyFile = "" Then
                        PrintOut = True
        '                MsgBox "rptCOM004three.rpt 파일이 없습니다.", vbCritical, "정보확인"
                        MsgBox "rptAPS015.rpt 파일이 없습니다.", vbCritical, "정보확인"
                        Exit Function
                    End If

        '            strRptNm = App.Path & "\rpt\rptCOM004three.rpt"
                    strRptNm = App.Path & "\rpt\rptAPS015.rpt"
            End Select

        Case "4"    'C1,F1,F2,F3
            Select Case strHeadTmp
                Case "C1F1F2F3" 'LC3
                    strOpt = "5"

                    strTmp = GetRecord(objSql.GetComCdMST2(pCdindex), strOpt)
            End Select

'            strMyFile = Dir(App.Path & "\rpt\rptCOM004four.rpt")
            strMyFile = Dir(App.Path & "\rpt\rptAPS016.rpt")

            If strMyFile = "" Then
                PrintOut = True
'                MsgBox "rptCOM004four.rpt 파일이 없습니다.", vbCritical, "정보확인"
                MsgBox "rptAPS016.rpt 파일이 없습니다.", vbCritical, "정보확인"
                Exit Function
            End If

'            strRptNm = App.Path & "\rpt\rptCOM004four.rpt"
            strRptNm = App.Path & "\rpt\rptAPS016.rpt"
        Case "5"    'C1,F1,F2,F3,F4,, C1,F1,F2,T1,T2
            Select Case strHeadTmp
                Case "C1F1F2F3F4"   'LC3
                    strOpt = "6"

                    strTmp = GetRecord(objSql.GetComCdMST2(pCdindex), strOpt)

'                    strMyFile = Dir(App.Path & "\rpt\rptCOM004five.rpt")
                    strMyFile = Dir(App.Path & "\rpt\rptAPS017.rpt")

                    If strMyFile = "" Then
                        PrintOut = True
'                        MsgBox "rptCOM004five.rpt 파일이 없습니다.", vbCritical, "정보확인"
                        MsgBox "rptAPS017.rpt 파일이 없습니다.", vbCritical, "정보확인"
                        Exit Function
                    End If

'                    strRptNm = App.Path & "\rpt\rptCOM004five.rpt"
                    strRptNm = App.Path & "\rpt\rptAPS017.rpt"
                Case "C1F1F2T1T2"   'LC4
                    strOpt = "7"

                    strTmp = GetRecord(objSql.GetComCdTemp(pCdindex), strOpt)

'                    strMyFile = Dir(App.Path & "\rpt\rptCOM004SugicalDx.rpt")
                    strMyFile = Dir(App.Path & "\rpt\rptAPS018.rpt")

                    If strMyFile = "" Then
                        PrintOut = True
'                        MsgBox "rptCOM004SugicalDx.rpt 파일이 없습니다.", vbCritical, "정보확인"
                        MsgBox "rptAPS018.rpt 파일이 없습니다.", vbCritical, "정보확인"
                        Exit Function
                    End If

'                    strRptNm = App.Path & "\rpt\rptCOM004SugicalDx.rpt"
                    strRptNm = App.Path & "\rpt\rptAPS018.rpt"
            End Select
        Case "6"    'C1,F1,F2,F3,F4,T1
            Select Case strHeadTmp
                Case "C1F1F2F3F4T1" 'LC3
                    strOpt = "8"

                    strTmp = GetRecord(objSql.GetComCdMST2(pCdindex), strOpt)
            End Select

'            strMyFile = Dir(App.Path & "\rpt\rptCOM004six.rpt")
            strMyFile = Dir(App.Path & "\rpt\rptAPS019.rpt")

            If strMyFile = "" Then
                PrintOut = True
'                MsgBox "rptCOM004six.rpt 파일이 없습니다.", vbCritical, "정보확인"
                MsgBox "rptAPS019.rpt 파일이 없습니다.", vbCritical, "정보확인"
                Exit Function
            End If

'            strRptNm = App.Path & "\rpt\rptCOM004six.rpt"
            strRptNm = App.Path & "\rpt\rptAPS019.rpt"
        Case "7"    'C1,F1,F2,F3,F4,T1,T2
            Select Case strHeadTmp
                Case "C1F1F2F3F4T1T2"   'LC3
                    strOpt = "9"

                    strTmp = GetRecord(objSql.GetComCdMST2(pCdindex), strOpt)
            End Select

'            strMyFile = Dir(App.Path & "\rpt\rptAPSBethe.rpt")
            strMyFile = Dir(App.Path & "\rpt\rptAPS013.rpt")

            If strMyFile = "" Then
                PrintOut = True
'                MsgBox "rptAPSBethe.rpt 파일이 없습니다.", vbCritical, "정보확인"
                MsgBox "rptAPS013.rpt 파일이 없습니다.", vbCritical, "정보확인"
                Exit Function
            End If

'            strRptNm = App.Path & "\rpt\rptAPSBethe.rpt"
            strRptNm = App.Path & "\rpt\rptAPS013.rpt"
    End Select

    strMyFile = Dir(App.Path & "\rpt\CrystalReport.txt")
    If strMyFile = "" Then
        PrintOut = True
        MsgBox "CrystalReport.txt 파일이 없습니다.", vbCritical, "정보확인"
        Exit Function
    End If
    strMyFile = ""
    strFileNm = App.Path & "\rpt\CrystalReport.txt"

    If strTmp = "Error" Or strTmp = "NoData" Then
        PrintOut = False
        Exit Function
    End If

    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)

'    Debug.Print strTmp

    lngFNum = FreeFile

On Error GoTo ErrPrint

    Open strFileNm For Output As #lngFNum
    Print #lngFNum, strTmp
    Close #lngFNum
    With crtReport
        .ReportFileName = strRptNm
        .ParameterFields(0) = "title;" & strTitle & ";true"

        Select Case strHeadCnt
            Case "2"
                .ParameterFields(1) = "head1;" & aryHead(0) & ";true"
                .ParameterFields(2) = "head2;" & aryHead(1) & ";true"
            Case "3"
                .ParameterFields(1) = "head1;" & aryHead(0) & ";true"
                .ParameterFields(2) = "head2;" & aryHead(1) & ";true"
                .ParameterFields(3) = "head3;" & aryHead(2) & ";true"
            Case "4"
                .ParameterFields(1) = "head1;" & aryHead(0) & ";true"
                .ParameterFields(2) = "head2;" & aryHead(1) & ";true"
                .ParameterFields(3) = "head3;" & aryHead(2) & ";true"
                .ParameterFields(4) = "head4;" & aryHead(3) & ";true"
            Case "5"
                .ParameterFields(1) = "head1;" & aryHead(0) & ";true"
                .ParameterFields(2) = "head2;" & aryHead(1) & ";true"
                .ParameterFields(3) = "head3;" & aryHead(2) & ";true"
                .ParameterFields(4) = "head4;" & aryHead(3) & ";true"
                .ParameterFields(5) = "head5;" & aryHead(4) & ";true"
            Case "6"
                .ParameterFields(1) = "head1;" & aryHead(0) & ";true"
                .ParameterFields(2) = "head2;" & aryHead(1) & ";true"
                .ParameterFields(3) = "head3;" & aryHead(2) & ";true"
                .ParameterFields(4) = "head4;" & aryHead(3) & ";true"
                .ParameterFields(5) = "head5;" & aryHead(4) & ";true"
                .ParameterFields(6) = "head6;" & aryHead(5) & ";true"
            Case "7"
                .ParameterFields(1) = "head1;" & aryHead(0) & ";true"
                .ParameterFields(2) = "head2;" & aryHead(1) & ";true"
                .ParameterFields(3) = "head3;" & aryHead(2) & ";true"
                .ParameterFields(4) = "head4;" & aryHead(3) & ";true"
                .ParameterFields(5) = "head5;" & aryHead(4) & ";true"
                .ParameterFields(6) = "head6;" & aryHead(5) & ";true"
                .ParameterFields(7) = "head7;" & aryHead(6) & ";true"
        End Select

        .ParameterFields(8) = "hostnm;" & ObjSysInfo.Hospital & ";true"

        .RetrieveDataFiles
        .WindowState = 2 ' crptMaximized
        .Destination = crptToWindow
        .Action = 1
        .Reset
    End With
    PrintOut = True
    Exit Function

ErrPrint:
    PrintOut = False
End Function

    

Private Function GetRecord(ByVal pSql As String, ByVal pOpt As String) As String
'pOpt "1" : C1,F1 "2" : C1,T1 "3" : C1,F1,T1 "4" : C1,F1,F2
'     "5" : C1,F1,F2,F3 "6" : C1,F1,F2,F3,F4 "7" : C1,F1,F2,T1,T2
'     "8" : C1,F1,F2,F3,F4,T1 "9" : C1,F1,F2,F3,F4,F5,T1,T2


    Dim RS As Recordset
    Dim i As Long
    Dim strTmp As String
    Dim strtmp1 As String
    Dim strTmp2 As String
    Dim strTmp3 As String
    Dim strTmp4 As String
    Dim strTmp5 As String
    Dim strTmp6 As String
    Dim strTmp7 As String
    Dim lngCnt As Long

    With objSql
        Set RS = New Recordset
        RS.Open pSql, DBConn
    End With

On Error GoTo ErrTrap

    If RS.EOF Then
        GetRecord = "NoData"
    Else
        objProbar.Message = ""
        objProbar.Max = RS.RecordCount
        Select Case pOpt
            Case "1"
                Do Until RS.EOF
                    i = i + 1
                    strtmp1 = "" & RS.Fields("cdval1").Value
                    strTmp2 = "" & RS.Fields("field1").Value

                    strTmp = strTmp & strtmp1 & vbTab & strTmp2 & vbTab
                    strTmp = strTmp & vbCr
                    RS.MoveNext

                    objProbar.Value = i
                Loop
            Case "2"
                Do Until RS.EOF
                    i = i + 1
                    strtmp1 = "" & RS.Fields("cdval1").Value
                    strTmp2 = Replace("" & RS.Fields("text1").Value, vbCr, Space(1))
                    strTmp2 = Replace(strTmp2, vbLf, Space(1))

                    strTmp = strTmp & strtmp1 & vbTab & strTmp2 & vbTab
                    strTmp = strTmp & vbCr
                    RS.MoveNext
                    objProbar.Value = i
                Loop
            Case "3"
                Do Until RS.EOF
                    i = i + 1
                    strtmp1 = "" & RS.Fields("cdval1").Value
                    strTmp2 = "" & RS.Fields("field1").Value
                    strTmp3 = Replace("" & RS.Fields("text1").Value, vbCr, Space(1))
                    strTmp3 = Replace(strTmp3, vbLf, Space(1))

                    strTmp = strTmp & strtmp1 & vbTab & strTmp2 & vbTab & strTmp3 & vbTab
                    strTmp = strTmp & vbCr
                    RS.MoveNext
                    objProbar.Value = i
                Loop
            Case "4"
                Do Until RS.EOF
                    i = i + 1
                    strtmp1 = "" & RS.Fields("cdval1").Value
                    strTmp2 = "" & RS.Fields("field1").Value
                    strTmp3 = "" & RS.Fields("field2").Value

                    strTmp = strTmp & strtmp1 & vbTab & strTmp2 & vbTab & strTmp3 & vbTab
                    strTmp = strTmp & vbCr
                    RS.MoveNext
                    objProbar.Value = i
                Loop
            Case "5"
                Do Until RS.EOF
                    i = i + 1
                    strtmp1 = "" & RS.Fields("cdval1").Value
                    strTmp2 = "" & RS.Fields("field1").Value
                    strTmp3 = "" & RS.Fields("field2").Value
                    strTmp4 = "" & RS.Fields("field3").Value

                    strTmp = strTmp & strtmp1 & vbTab & strTmp2 & vbTab & strTmp3 & vbTab & _
                             strTmp4 & vbTab
                    strTmp = strTmp & vbCr
                    RS.MoveNext
                    objProbar.Value = i
                Loop
            Case "6"
                Do Until RS.EOF
                    i = i + 1
                    strtmp1 = "" & RS.Fields("cdval1").Value
                    strTmp2 = "" & RS.Fields("field1").Value
                    strTmp3 = "" & RS.Fields("field2").Value
                    strTmp4 = "" & RS.Fields("field3").Value
                    strTmp5 = "" & RS.Fields("field4").Value

                    strTmp = strTmp & strtmp1 & vbTab & strTmp2 & vbTab & strTmp3 & vbTab & _
                             strTmp4 & vbTab & strTmp5 & vbTab
                    strTmp = strTmp & vbCr
                    RS.MoveNext
                    objProbar.Value = i
                Loop
            Case "7"
                Do Until RS.EOF
                    i = i + 1
                    strtmp1 = "" & RS.Fields("cdval1").Value
                    strTmp2 = "" & RS.Fields("field1").Value
                    strTmp3 = "" & RS.Fields("field2").Value
                    strTmp4 = Replace("" & RS.Fields("text1").Value, vbCr, Space(1))
                    strTmp4 = Replace(strTmp4, vbLf, Space(1))
                    strTmp5 = Replace("" & RS.Fields("text2").Value, vbCr, Space(1))
                    strTmp5 = Replace(strTmp5, vbLf, Space(1))

                    strTmp = strTmp & strtmp1 & vbTab & strTmp2 & vbTab & strTmp3 & vbTab & _
                             strTmp4 & vbTab & strTmp5 & vbTab
                    strTmp = strTmp & vbCr
                    RS.MoveNext
                    objProbar.Value = i
                Loop
            Case "8"
                Do Until RS.EOF
                    i = i + 1
                    strtmp1 = "" & RS.Fields("cdval1").Value
                    strTmp2 = "" & RS.Fields("field1").Value
                    strTmp3 = "" & RS.Fields("field2").Value
                    strTmp4 = "" & RS.Fields("field3").Value
                    strTmp5 = "" & RS.Fields("field4").Value
                    strTmp6 = Replace("" & RS.Fields("text1").Value, vbCr, Space(1))
                    strTmp6 = Replace(strTmp6, vbLf, Space(1))

                    strTmp = strTmp & strtmp1 & vbTab & strTmp2 & vbTab & strTmp3 & vbTab & _
                             strTmp4 & vbTab & strTmp5 & vbTab & strTmp6 & vbTab
                    strTmp = strTmp & vbCr
                    RS.MoveNext
                    objProbar.Value = i
                Loop
            Case "9"
                Do Until RS.EOF
                    i = i + 1
                    strtmp1 = "" & RS.Fields("cdval1").Value
                    strTmp2 = "" & RS.Fields("field1").Value
                    strTmp3 = "" & RS.Fields("field2").Value
                    strTmp4 = "" & RS.Fields("field3").Value
                    strTmp5 = "" & RS.Fields("field4").Value
                    strTmp6 = "" & RS.Fields("text1").Value
                    strTmp7 = "" & RS.Fields("text2").Value

                    strTmp = strTmp & strtmp1 & vbTab & strTmp2 & vbTab & strTmp3 & vbTab & _
                             strTmp4 & vbTab & strTmp5 & vbTab & strTmp6 & vbTab & strTmp7 & vbTab
                    strTmp = strTmp & vbCr
                    RS.MoveNext
                    objProbar.Value = i
                Loop
        End Select
    End If

    GetRecord = strTmp
    Set RS = Nothing
    Set objSql = Nothing
    Set objProbar = Nothing
    Exit Function

ErrTrap:
    GetRecord = "Error"
    Set RS = Nothing
    Set objSql = Nothing
    Set objProbar = Nothing
End Function

Private Sub cmdSave_Click()
    
    Dim i As Integer, sIndexKey As String, sFlag As String
    Dim SSQL As String, dsChk As Recordset
    Dim strTmp As String

    If Trim(txtSubKey) = "" Then
        MsgBox "Index가 지정이 잘못 되었습니다. 확인 후 처리 바랍니다."
        Exit Sub
    End If

    strTmp = txtVal(5).Text
    For i = Len(txtVal(5).Text) To 1 Step -1
        If Mid(strTmp, i, 1) <> vbCr And Mid(strTmp, i, 1) <> vbLf Then
            Exit For
        Else
            strTmp = Mid(strTmp, 1, i - 1)
        End If
    Next
    txtVal(5).Text = strTmp

    strTmp = txtVal(6).Text
    For i = Len(txtVal(6).Text) To 1 Step -1
        If Mid(strTmp, i, 1) <> vbCr And Mid(strTmp, i, 1) <> vbLf Then
            Exit For
        Else
            strTmp = Mid(strTmp, 1, i - 1)
        End If
    Next
    txtVal(6).Text = strTmp

    sFlag = "I"

    Set objSql = New clsLISSqlCodeMaster
    
    Select Case lblTable
        Case T_COM003
            SSQL = objSql.GetComCdMST2(Rkey, Trim(txtSubKey))
        Case T_COM004
            SSQL = objSql.GetComCdTemp(Rkey, Trim(txtSubKey))
    End Select
    
    With objSql
        Set dsChk = New Recordset
        dsChk.Open SSQL, DBConn
    End With
    
    If Not dsChk.EOF Then sFlag = "U"
    
    Set dsChk = Nothing
    Set objSql = Nothing

    Select Case sFlag
        Case "I": Call CommonInsert
        Case "U": Call CommonUpdate
        Case Else: MsgBox "시스템에 오류가 있습니다."
    End Select
    
    ChangeFlag = False
    cmdClear_Click

End Sub


Private Sub CommonInsert()
    
    Dim sData As String
    Dim SSQL(1) As String
    Dim ii As Long

    Set objSql = New clsLISSqlCodeMaster
    
    Select Case lblTable
        Case T_COM003
            SSQL(0) = objSql.SetComCdMST2(False, Rkey, txtSubKey, txtVal(0), txtVal(1), txtVal(2), _
                                           txtVal(3), txtVal(4), txtVal(5), txtVal(6))
        Case T_COM004
            SSQL(0) = objSql.SetComCdTemp(False, Rkey, txtSubKey, txtVal(0), txtVal(1), txtVal(5), _
                                           txtVal(6))
    End Select
    
    DBConn.BeginTrans
    If InsertData(SSQL, False) Then
        DBConn.CommitTrans
        
        Call LoadSubKey
    Else
        DBConn.RollbackTrans
        MsgBox Err.Description, vbExclamation
    End If
    
    Set objSql = Nothing

End Sub

Private Sub CommonUpdate()
    
    Dim sData As String
    Dim SSQL(1) As String

    Set objSql = New clsLISSqlCodeMaster
    
    Select Case lblTable
        Case T_COM003
            SSQL(0) = objSql.SetComCdMST2(True, Rkey, Trim(txtSubKey), Trim(txtVal(0)), Trim(txtVal(1)), _
                                          Trim(txtVal(2)), Trim(txtVal(3)), Trim(txtVal(4)), Trim(txtVal(5)), _
                                          Trim(txtVal(6)))
        Case T_COM004
            SSQL(0) = objSql.SetComCdTemp(True, Rkey, Trim(txtSubKey), Trim(txtVal(0)), Trim(txtVal(1)), _
                                          Trim(txtVal(5)), Trim(txtVal(6)))
    End Select
    
    DBConn.BeginTrans
    If InsertData(SSQL, False) Then
        DBConn.CommitTrans
    Else
        DBConn.RollbackTrans
        MsgBox Err.Description, vbExclamation
    End If
    
    Set objSql = Nothing
    
End Sub

Private Sub Form_Load()
    
    Me.WindowState = 2
    ChangeFlag = False
    chkKeyLock.Value = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objProbar = Nothing
    Set objSql = Nothing
    Set objCodeList = Nothing
End Sub

Private Sub lblRName_Change()
    LoadFieldNm
    LoadSubKey
    ChangeFlag = False
End Sub

Private Sub LoadFieldNm()

    Dim SSQL As String
    Dim dsSKey As Recordset
    Dim strFields As String
    Dim strTmp As String
    Dim strCOM001_CdIndex As String
    Dim i As Integer
    Dim aryTmp() As String
    
    If Mid(Rkey, 1, 2) = "C2" Then
        strCOM001_CdIndex = "LC2"
    Else
        strCOM001_CdIndex = "LC3"
    End If
    
    Set objSql = New clsLISSqlCodeMaster
    With objSql
        Set dsSKey = New Recordset
        dsSKey.Open .GetComCdIndex(strCOM001_CdIndex, Rkey), DBConn
    End With
    
    If Not dsSKey.EOF Then
        strFields = "" & dsSKey.Fields("text1").Value
        If strCOM001_CdIndex = "LC3" Then
            aryTmp = Split(strFields, ";")
            ReDim Preserve aryTmp(7)
            aryTmp(6) = aryTmp(3)
            aryTmp(7) = aryTmp(4)
            aryTmp(3) = "": aryTmp(4) = ""
            strFields = Join(aryTmp, ";")
        End If
    Else
        strFields = ""
    End If
    
    Set dsSKey = Nothing
    Set objSql = Nothing

    For i = 0 To lblCap.Count - 1
        strTmp = medShift(strFields, ";")
        lblCap(i).Caption = medShift(strTmp, ":")
        If i > 0 Then
            If Trim(lblCap(i)) = "" Then
                txtVal(i - 1).Visible = False
            Else
                txtVal(i - 1).Visible = True
                txtVal(i - 1).MaxLength = Val(strTmp)
            End If
        End If
    Next
End Sub

Private Sub LoadData()
    
    Dim SSQL As String
    Dim dsInfo As Recordset

    If lstSubKey.ListIndex < 0 Then Exit Sub
    
    
    Set objSql = New clsLISSqlCodeMaster
    
    Select Case lblTable
        Case T_COM003
            SSQL = objSql.GetComCdMST2(Rkey, medGetP(lstSubKey.List(lstSubKey.ListIndex), 1, Chr$(9)))
        Case T_COM004
            SSQL = objSql.GetComCdTemp(Rkey, medGetP(lstSubKey.List(lstSubKey.ListIndex), 1, Chr$(9)))
    End Select
    
    With objSql
        Set dsInfo = New Recordset
        dsInfo.Open SSQL, DBConn
    End With
    
    Call ClearScreen

    If dsInfo.EOF Then
        MsgBox "등록되어 있지 않은 ID 입니다."
        GoTo NoData
    End If
    
    If lblTable = T_COM003 Then
        txtSubKey.Text = "" & dsInfo.Fields("cdval1").Value: lblCaption(0).Caption = medGetP(lstSubKey.List(lstSubKey.ListIndex), 2, Chr$(9))
        txtVal(0).Text = "" & dsInfo.Fields("field1").Value: lblCaption(1).Caption = GetSpcName(txtVal(0).Text)
        txtVal(1).Text = "" & dsInfo.Fields("field2").Value: lblCaption(2).Caption = GetSpcName(txtVal(1).Text)
        txtVal(2).Text = "" & dsInfo.Fields("field3").Value
        txtVal(3).Text = "" & dsInfo.Fields("field4").Value
        txtVal(4).Text = "" & dsInfo.Fields("field5").Value
        txtVal(5).Text = "" & dsInfo.Fields("text1").Value
        txtVal(6).Text = "" & dsInfo.Fields("text2").Value
    Else
        txtSubKey.Text = "" & dsInfo.Fields("cdval1").Value
        txtVal(0).Text = "" & dsInfo.Fields("field1").Value
        txtVal(1).Text = "" & dsInfo.Fields("field2").Value
        txtVal(5).Text = "" & dsInfo.Fields("text1").Value
        txtVal(6).Text = "" & dsInfo.Fields("text2").Value
    End If

    ' 데이타 읽고 나서 키를 바꿀수 없게..
    txtSubKey.Locked = chkKeyLock.Value

NoData:
    Set dsInfo = Nothing
    Set objSql = Nothing

End Sub


Private Sub lstSubKey_Click()
    Call LoadData
End Sub


Private Function ConfirmExit() As Boolean

    Dim intResp As VbMsgBoxResult

    ConfirmExit = True
    If ChangeFlag Then
        intResp = MsgBox("변경된 내용을 저장하지 않고 진행하시겠습니까 ? ", vbYesNo)
        If intResp = vbNo Then
            ConfirmExit = False
            Exit Function
        End If
    End If
    ChangeFlag = False

End Function


Private Sub txtVal_Change(Index As Integer)
    ChangeFlag = True
End Sub

Private Sub txtVal_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Index < 5 Then SendKeys "{TAB}"

End Sub

Private Sub txtSubKey_KeyPress(KeyAscii As Integer)
    If Not ConfirmExit Then
       KeyAscii = 0
       Exit Sub
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtSubKey_LostFocus()
    Dim lngIdx As Long
    lngIdx = medListFind(lstSubKey, txtSubKey.Text)
    If Trim(txtSubKey.Text) <> medGetP(lstSubKey.List(lngIdx), 1, vbTab) Then
        Dim SSQL As String
        Dim RS   As Recordset
        
        lblCaption(0).Caption = ""
        lstSubKey.ListIndex = -1
        SSQL = GetTestItemSQL(Trim(txtSubKey.Text))
        If SSQL = "" Then Exit Sub
        Set RS = New Recordset
        RS.Open SSQL, DBConn
        If Not RS.EOF Then
            lblCaption(0).Caption = RS.Fields("testnm").Value & ""
        Else
            txtSubKey.Text = ""
        End If
        Set RS = Nothing
        Exit Sub
    Else
        lstSubKey.ListIndex = lngIdx
    End If
    Call LoadData
    ChangeFlag = False
End Sub
'검사항목 찾기
Private Function GetTestItemSQL(Optional ByVal sTestCd As String = "") As String
    Dim SSQL As String

    Select Case Rkey
        Case LC3_ByPass, LC3_PocItem, LC3_HighItem, LC3_ICUTestCd, LC3_POCTestCd, _
             LC3_ReportTesctCd, LC3_WBCDiffCode, LC3_WBCCode, LC3_NRBCCode, LC3_RESULTREADTEST
             
            SSQL = " select a.testcd,a.testnm from " & T_LAB001 & " a" & _
                   " where  a.applydt = ( select max(applydt) from " & T_LAB001 & _
                   "                     where testcd = a.testcd )"
            If sTestCd <> "" Then SSQL = SSQL & " and " & DBW("a.testcd=", sTestCd)
    End Select
    GetTestItemSQL = SSQL
End Function
'검체 찾기
Private Function GetSpcItemSQL(Optional ByVal SpcCd As String = "") As String
    Dim SSQL As String

    Select Case Rkey
        Case LC3_ByPass, LC3_PocItem, LC3_HighItem, LC3_ICUTestCd, LC3_POCTestCd, _
             LC3_ReportTesctCd, LC3_WBCDiffCode, LC3_WBCCode, LC3_NRBCCode
            SSQL = " Select a.spccd as spccd, b.field3 as spcnm " & _
                   " From  " & T_LAB004 & " a, " & T_LAB032 & " b " & _
                   " Where " & DBW("a.testcd", Trim(UCase(txtSubKey.Text)), 2) & _
                   " and   " & DBW("b.cdindex", LC3_Specimen, 2) & _
                   " and     b.cdval1 = a.spccd "
            If SpcCd <> "" Then SSQL = SSQL & " and " & DBW("b.cdval1=", SpcCd)
    End Select
    GetSpcItemSQL = SSQL
End Function
'검체명으로 전환
Private Function GetSpcName(ByVal sCode As String)
    Dim RS      As Recordset
    Dim SSQL    As String
    Select Case Rkey
        Case LC3_ByPass, LC3_PocItem, LC3_HighItem
            SSQL = " select field3 as spcnm from " & T_LAB032 & " a" & _
                   " where  " & DBW("cdindex=", LC3_Specimen) & " and " & DBW("cdval1=", sCode)
            Set RS = New Recordset
            RS.Open SSQL, DBConn
            
            If Not RS.EOF Then
                GetSpcName = RS.Fields("spcnm").Value & ""
            Else
                GetSpcName = sCode
            End If
            Set RS = Nothing
        
        Case Else
            GetSpcName = sCode
    End Select
End Function

Private Sub cmdPopup_Click(Index As Integer)
    Dim tmpSql  As String
    Dim lngTop  As Long
    Dim lngLeft As Long
    
    Set objCodeList = New clsPopUpList
    With objCodeList
            .FormCaption = "코드 리스트"
            .ColumnHeaderText = "코드;코드명"
            .Connection = DBConn
        Select Case Index
            Case 0: tmpSql = GetTestItemSQL:
                    lngTop = txtSubKey.Top + 2350
                    lngLeft = Me.Left + Frame1.Left + txtSubKey.Left + 50
                    txtSubKey.Text = "": txtVal(0).Text = ""
                    lblCaption(0).Caption = "": lblCaption(1).Caption = ""
                    Call .LoadPopUp(tmpSql) ', lngTop, lngLeft)
                    txtSubKey.Text = medGetP(.SelectedString, 1, ";")
                    lblCaption(0).Caption = medGetP(.SelectedString, 2, ";")
                    If LC3_ReportTesctCd = Rkey Then
                        txtVal(1).Text = "": lblCaption(2).Caption = ""
                        txtVal(0).Text = medGetP(.SelectedString, 2, ";")
                    End If
            Case 1: tmpSql = GetSpcItemSQL
                    lngTop = txtVal(0).Top + 2350
                    lngLeft = Me.Left + Frame1.Left + txtVal(0).Left + 50
                    lblCaption(1).Caption = "": txtVal(0).Text = ""
                    Call .LoadPopUp(tmpSql) ', lngTop, lngLeft)
                    txtVal(0).Text = medGetP(.SelectedString, 1, ";")
                    lblCaption(1).Caption = medGetP(.SelectedString, 2, ";")
            Case 2: tmpSql = GetSpcItemSQL
                    lngTop = txtVal(1).Top + 2350
                    lngLeft = Me.Left + Frame1.Left + txtVal(1).Left + 50
                    lblCaption(2).Caption = "": txtVal(1).Text = ""
                    Call .LoadPopUp(tmpSql) ', lngTop, lngLeft)
                    txtVal(1).Text = medGetP(.SelectedString, 1, ";")
                    lblCaption(2).Caption = medGetP(.SelectedString, 2, ";")
        End Select
    End With
    Set objCodeList = Nothing
End Sub

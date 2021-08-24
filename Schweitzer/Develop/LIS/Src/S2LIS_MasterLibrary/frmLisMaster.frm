VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm366EDefine 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "특수검사 Template 설정"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H00DBC8D2&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   120
      ScaleHeight     =   540
      ScaleWidth      =   10680
      TabIndex        =   29
      Top             =   195
      Width           =   10680
      Begin MSComctlLib.TabStrip TabItem 
         Height          =   390
         Left            =   135
         TabIndex        =   30
         Top             =   75
         Width           =   10470
         _ExtentX        =   18468
         _ExtentY        =   688
         MultiRow        =   -1  'True
         Style           =   2
         TabFixedWidth   =   1408
         TabFixedHeight  =   616
         Separators      =   -1  'True
         TabMinWidth     =   0
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H009CA3C5&
         BorderWidth     =   3
         Height          =   450
         Left            =   60
         Shape           =   4  '둥근 사각형
         Top             =   45
         Width           =   10575
      End
   End
   Begin VB.CommandButton cmdDelGrp 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Delete(&D)"
      Height          =   495
      Left            =   5490
      Style           =   1  '그래픽
      TabIndex        =   28
      Top             =   8070
      Width           =   1245
   End
   Begin VB.CommandButton cmdFormExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Exit(&X)"
      Height          =   495
      Left            =   9540
      Style           =   1  '그래픽
      TabIndex        =   27
      Top             =   8070
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Save(&S)"
      Height          =   495
      Left            =   8190
      Style           =   1  '그래픽
      TabIndex        =   26
      Top             =   8070
      Width           =   1245
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Clear(&C)"
      Height          =   495
      Left            =   6840
      Style           =   1  '그래픽
      TabIndex        =   25
      Top             =   8070
      Width           =   1245
   End
   Begin VB.Frame fraTmp 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Template 정보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6510
      Left            =   3810
      TabIndex        =   13
      Top             =   1380
      Width           =   6750
      Begin VB.CommandButton cmdSaveTmp 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Save"
         Height          =   345
         Left            =   5820
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   6120
         Width           =   795
      End
      Begin VB.CommandButton cmdClearTmp 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Clear"
         Height          =   345
         Left            =   4980
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   6120
         Width           =   795
      End
      Begin VB.TextBox txtTmpResultNm 
         BackColor       =   &H00F1F5F4&
         Height          =   285
         Left            =   4800
         TabIndex        =   15
         Top             =   270
         Width           =   1575
      End
      Begin VB.TextBox txtTmpCode 
         BackColor       =   &H00F1F5F4&
         Height          =   270
         Left            =   960
         MaxLength       =   4
         TabIndex        =   14
         Top             =   240
         Width           =   2205
      End
      Begin RichTextLib.RichTextBox txtTmpData 
         Height          =   5415
         Left            =   150
         TabIndex        =   18
         Top             =   690
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   9551
         _Version        =   393217
         BackColor       =   15857140
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmLisMaster.frx":0000
      End
      Begin VB.Label Label3 
         BackColor       =   &H00DBE6E6&
         Caption         =   "결과이름"
         Height          =   225
         Left            =   3990
         TabIndex        =   20
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DBE6E6&
         Caption         =   "코드이름"
         Height          =   285
         Left            =   150
         TabIndex        =   19
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame FraForm 
      BackColor       =   &H00DBE6E6&
      Height          =   7350
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   10680
      Begin VB.TextBox txtStcd 
         BackColor       =   &H00F1F5F4&
         Height          =   270
         Left            =   1290
         MaxLength       =   1
         TabIndex        =   7
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtGnm 
         BackColor       =   &H00F1F5F4&
         Height          =   270
         Left            =   3810
         TabIndex        =   6
         Top             =   270
         Width           =   945
      End
      Begin VB.CommandButton cmdTmpHelp 
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
         Left            =   3330
         Picture         =   "frmLisMaster.frx":0294
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   780
         Width           =   315
      End
      Begin VB.CommandButton cmDelSelTmp 
         Appearance      =   0  '평면
         BackColor       =   &H00F4F0F2&
         Caption         =   "제거"
         Height          =   345
         Left            =   3060
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   6930
         Width           =   585
      End
      Begin VB.ListBox lstSTest 
         Appearance      =   0  '평면
         BackColor       =   &H00EEEEEE&
         Height          =   2190
         Left            =   210
         TabIndex        =   3
         Top             =   1080
         Width           =   3435
      End
      Begin VB.CommandButton cmDelSelSTest 
         Appearance      =   0  '평면
         BackColor       =   &H00F4F0F2&
         Caption         =   "제거"
         Height          =   285
         Left            =   3060
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   3300
         Width           =   585
      End
      Begin VB.CommandButton cmAddTmp 
         Appearance      =   0  '평면
         BackColor       =   &H00F4F0F2&
         Caption         =   "추가"
         Height          =   345
         Left            =   2400
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   6930
         Width           =   585
      End
      Begin FPSpread.vaSpread spdTmpResult 
         Height          =   2925
         Left            =   210
         TabIndex        =   8
         Top             =   3960
         Width           =   3435
         _Version        =   196608
         _ExtentX        =   6059
         _ExtentY        =   5159
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   3
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "frmLisMaster.frx":07C6
         UserResize      =   0
      End
      Begin VB.Label Label7 
         BackColor       =   &H00DBE6E6&
         Caption         =   "코드이름 :"
         Height          =   270
         Left            =   330
         TabIndex        =   12
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label6 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사항목그룹이름 : "
         Height          =   330
         Left            =   1950
         TabIndex        =   11
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "특수 검사항목 설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   10
         Top             =   780
         Width           =   2835
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "텍스트결과 Template Code 설정"
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
         Left            =   270
         TabIndex        =   9
         Top             =   3630
         Width           =   3345
      End
   End
   Begin VB.Frame fraToSTest 
      BackColor       =   &H00DBE6E6&
      Height          =   3735
      Left            =   840
      TabIndex        =   21
      Top             =   5070
      Visible         =   0   'False
      Width           =   3585
      Begin VB.CommandButton cmdAddSTest 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Add"
         Height          =   345
         Left            =   60
         Style           =   1  '그래픽
         TabIndex        =   23
         Top             =   3330
         Width           =   645
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Exit"
         Height          =   345
         Left            =   2850
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   3330
         Width           =   675
      End
      Begin FPSpread.vaSpread spdToSTest 
         Height          =   3135
         Left            =   60
         TabIndex        =   24
         Top             =   150
         Width           =   3465
         _Version        =   196608
         _ExtentX        =   6112
         _ExtentY        =   5530
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   10
         OperationMode   =   1
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "frmLisMaster.frx":2697
      End
   End
End
Attribute VB_Name = "frm366EDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iPspdtmpresult As Integer
Private iPspdToSTest As Integer
Private iPspdNumResultItem As Integer
Private iMode As Integer ' 1: 생성  2: Edit
    

Private aTmpvalue() As tTmpValue
Private iTmpArrayCnt As Integer

Private objSql As New clsLISSqlMasters

'************************************************************************************
'                                     Clear Function
'************************************************************************************

Private Sub ClearTmpFrame()
    txtTmpCode.Text = ""
    txtTmpCode.Enabled = False
    txtTmpResultNm.Text = ""
    txtTmpResultNm.Enabled = False
    txtTmpData.Text = ""
    
End Sub

Private Sub ClearspdToSTest()
    Dim I%
    
    With spdToSTest
        For I = 1 To .DataRowCnt
            .Row = I
            .Col = 1
            .Value = 0
        Next I
    End With
End Sub

Private Sub clearPositionVar()
    
    iPspdtmpresult = 0
    iPspdToSTest = 0
    iPspdNumResultItem = 0
  
End Sub

Private Sub cmAddTmp_Click()
    If txtStcd.Text = "" Then
        MsgBox " 코드이름이 존재하지 않습니다."
        Exit Sub
    End If
    
    Call ClearTmpFrame
    txtTmpCode.Enabled = True
    txtTmpResultNm.Enabled = True
    txtTmpCode.SetFocus
    iMode = 1           ' Setting iMode variable
    
End Sub
Public Sub ClearAll()
    txtStcd.Text = ""
    txtGnm.Text = ""
    
    txtStcd.Enabled = True
    txtGnm.Enabled = True
    
    txtStcd.SetFocus
    
    Call ClearlstSTest
    Call ClearspdTmpResult
    Call ClearTmpFrame
End Sub
Public Sub ClearlstSTest()
    lstSTest.Clear
End Sub
Public Sub ClearspdTmpResult()
    If spdTmpResult.MaxRows <> 0 Then
        spdTmpResult.Col = 1: spdTmpResult.Col2 = spdTmpResult.MaxCols
        spdTmpResult.Row = 1: spdTmpResult.Row2 = spdTmpResult.MaxRows
        spdTmpResult.BlockMode = True
        spdTmpResult.Action = ActionClearText
        spdTmpResult.BlockMode = False
        
        spdTmpResult.MaxRows = 0
    End If
End Sub

'************************************************************************************
'************************************************************************************

Private Sub cmdAddSTest_Click()
    Dim I%, sAddData As String
    
    If txtStcd.Text = "" Then
        MsgBox " 코드이름을 먼저입력하십시요"
        fraToSTest.Visible = False
        txtStcd.SetFocus
        Exit Sub
    End If
    
    With spdToSTest
        For I = 1 To .DataRowCnt
            .Col = 1
            .Row = I
            If .Value = 1 Then ' Case checked
                .Col = 2: sAddData = Trim(.Text)
                .Col = 3: sAddData = sAddData & vbTab & Trim(.Text)
                If chkDuplicateSTest(sAddData) = False Then '  Case Not Duplicate
                    lstSTest.AddItem sAddData, lstSTest.ListCount ' insert into List
                    'Call Setrsttype(GetCodeStr(sAddData))    ' lab001의 rsttype 내용 Setting
                    fraToSTest.Visible = False
                Else                                    ' case Duplicate
                    .Col = 1
                    .Value = 0
                End If
            End If
        Next I
    End With
End Sub

Private Sub Setrsttype(sTestCd As String)
    
    Dim sSQL As String
  '  On Error GoTo DBExecError

    sSQL = objSql.SqlUpdateRstType(sTestCd, Trim(txtStcd.Text), "1")
   ' DbConn.BeginTrans
    DBConn.Execute (sSQL)
   ' DbConn.CommitTrans

    Exit Sub
    
'DBExecError:
 '  DbConn.RollbackTrans

End Sub

Private Sub Delrsttype(sStcd As String)
' 현재 기타검사항목이 포함한 모든 testcd의 rsttype 삭제
    
    Dim sSQL As String
'    On Error GoTo DBExecError

    sSQL = objSql.SqlUpdateRstType(sStcd, "null", "2")
 '   DbConn.BeginTrans
    DBConn.Execute (sSQL)
  '  DbConn.CommitTrans

    Exit Sub
    
'DBExecError:
 '  DbConn.RollbackTrans

End Sub

Public Function chkDuplicateSTest(sAddData As String) As Boolean
    Dim I%
    
    For I = 0 To lstSTest.ListCount - 1
        If GetCodeStr(lstSTest.List(I)) = GetCodeStr(sAddData) Then
            MsgBox "duplicate"
            chkDuplicateSTest = True
            Exit Function
        End If
    Next I
    
    chkDuplicateSTest = False
End Function

Public Function chkDuplicateTmp(sTpcd As String) As Boolean
    Dim I%
    Dim sSQL As String
    Dim rsTmpCd As DrRecordSet
    
    If sTpcd = "N0" Then Exit Function
    
    With spdTmpResult                       ' 스프레드 중복 체크
        For I = 1 To .DataRowCnt
            .Row = I
            .Col = 1
            If Trim(.Text) = Trim(txtTmpCode.Text) Then
                MsgBox "duplicate"
                chkDuplicateTmp = True      ' duplicate
                Exit Function
            End If
        Next I
    End With
 
    Set rsTmpCd = OpenRecordSet(objSql.SqlGetLAB036("", sTpcd))
    If rsTmpCd.EOF = False Then
        MsgBox rsTmpCd.Fields("tpcd").Value & "는 " & _
                rsTmpCd.Fields("stcd").Value & "의 Template 입니다. "
        chkDuplicateTmp = True          ' duplicate
        rsTmpCd.RsClose
        Set rsTmpCd = Nothing
        Exit Function
    End If
    
    chkDuplicateTmp = False                 ' Not duplicate
    
    rsTmpCd.RsClose
    Set rsTmpCd = Nothing
    
End Function

Private Function GetCodeStr(sString As String) As String
    
    GetCodeStr = Trim(Mid(sString, 1, InStr(1, sString, vbTab) - 1))

End Function

Private Sub cmdClear_Click()
    Call ClearAll
End Sub

Private Sub cmdClearTmp_Click()
    If txtTmpCode.Enabled = True Then
        Call ClearTmpFrame
    Else
        txtTmpData.Text = ""
    End If
End Sub

Private Sub cmdDelGrp_Click()
    
    Dim sMsg As String
    Dim sRes As Integer, sStyle As Integer
    Dim I%
    
    If Trim(txtStcd.Text) = "" Then Exit Sub

    sMsg = txtStcd.Text & " 에 관한 정보를 모두 삭제합니다" & Chr$(13) & Chr$(10) & _
        "정말 삭제해도 좋습니까?"
    sStyle = vbYesNo + vbCritical + vbDefaultButton2
    
    sRes = MsgBox(sMsg, sStyle, "삭제 확인")
    If sRes = vbYes Then
        Call DeleteGrp
        Call ClearAll
        For I = 1 To tabItem.Tabs.Count - 1
            tabItem.Tabs.Remove 1
        Next I
        Call Initialize
'        medMain.stsBar.Panels(2).Text = "정상적으로 삭제 처리 되었습니다. 다음 작업을 처리하세요"
    Else
        Exit Sub
    End If
    
End Sub

Private Sub DeleteGrp()
    Dim sSqlDel350 As String  ' 기타검사화면 설정Table
    Dim sSqlDel036 As String  ' template Table
    Dim sSqlUpdate001 As String  ' rsttype = Null 설정
    Dim I%
    
    sSqlDel350 = objSql.SqlDeleteLAB350(txtStcd.Text)
    sSqlDel036 = objSql.SqlDeleteLAB036(txtStcd.Text)
    sSqlUpdate001 = objSql.SqlUpdateRstType(txtStcd.Text, "null", 2)
                    
    On Error GoTo DBExecError

    DBConn.BeginTrans
    DBConn.Execute (sSqlDel350)
    DBConn.Execute (sSqlDel036)
    DBConn.Execute (sSqlUpdate001)
    DBConn.CommitTrans
    
    Exit Sub

DBExecError:
   DBConn.RollbackTrans
    
End Sub

Private Sub FillfraTmp(sTmpCode As String)
    
    Dim iTmpArrayindex As Integer
    
    iTmpArrayindex = findTmpArrayIndex(sTmpCode)
    If iTmpArrayindex = 0 Then Exit Sub
    
    txtTmpCode.Text = aTmpvalue(iTmpArrayindex).sTpcd
    txtTmpResultNm.Text = aTmpvalue(iTmpArrayindex).sTpnm
    txtTmpData.Text = aTmpvalue(iTmpArrayindex).sTpdata
    
    
    txtTmpCode.Enabled = False
    txtTmpResultNm.Enabled = False
    
End Sub
Public Function findTmpArrayIndex(sTmpCode As String) As Integer
    Dim I%
    For I = 1 To iTmpArrayCnt
        If Trim(sTmpCode) = Trim(aTmpvalue(I).sTpcd) Then
            findTmpArrayIndex = I
            Exit Function
        End If
    Next I
    findTmpArrayIndex = 0
End Function
'Private Sub cmDelSelNRst_Click()
'    Dim sTpcd As String
'
'    If txtStcd.Text = "" Then
'        MsgBox " 코드이름이 존재하지 않습니다."
'        Exit Sub
'    End If
'
'    If iPspdNumResultItem < 1 Then ' case header selected
'        Exit Sub
'    End If
'
'    With spdNumResultItem
'        .Row = iPspdNumResultItem
'        .Action = ActionDeleteRow
'    End With
'
'
'End Sub

Private Sub cmDelSelSTest_Click()
    
    If txtStcd.Text = "" Then
        MsgBox " 코드이름을 먼저입력하십시요"
        Exit Sub
    End If
    
    If lstSTest.SelCount = 0 Then    ' Not exist Selected Item
        Exit Sub
    End If
    
    lstSTest.RemoveItem (lstSTest.ListIndex)  ' 선택된 특수검사항목을 리스트에서 제거

End Sub

Private Sub cmDelSelTmp_Click()
    Dim sTpcd As String
    Dim iTmpArrayindex As Integer
    
    If txtStcd.Text = "" Then
        MsgBox " 코드이름이 존재하지 않습니다."
        Exit Sub
    End If
    
    If iPspdtmpresult < 1 Then ' case header selected
        Exit Sub
    End If
        
        
    With spdTmpResult
        .Row = iPspdtmpresult
        .Col = 1
        sTpcd = .Text
        
'        If Trim(sTpcd) = "N0" Then
'            MsgBox "worksheet template입니다."
'            Exit Sub
'        End If
        
        .Action = ActionDeleteRow
        .MaxRows = .MaxRows - 1
    End With
    
    iTmpArrayindex = findTmpArrayIndex(Trim(sTpcd))
    If iTmpArrayindex = 0 Then Exit Sub
    
    aTmpvalue(iTmpArrayindex).sSaveDecision = 0     ' Unsaved
    
    Call ClearTmpFrame
End Sub

Private Sub cmdExit_Click()
    fraToSTest.Visible = False
End Sub

'
Private Sub cmdFormExit_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
    Dim I%
    
    If txtStcd.Text = "" Then
        
        MsgBox " 코드이름을 입력하시오"
        txtStcd.Enabled = True
        txtStcd.SetFocus
        Exit Sub
    End If
    
On Error GoTo DBExecError
    
    DBConn.BeginTrans
   
    Call SaveRstType ' ok
    Call SaveTmp 'ok
    Call SaveSTCD
    
    DBConn.CommitTrans
    
    Call ClearAll
    For I = 1 To tabItem.Tabs.Count - 1
        tabItem.Tabs.Remove 1
    Next I
    Call ClearTmpFrame
    Call Initialize
    
    Exit Sub
    
DBExecError:
   DBConn.RollbackTrans

End Sub

Private Sub SaveSTCD()

    Dim sDelSQL As String
    Dim sInsSql2 As String
    Dim objWsSql As New clsLISSqlMasters
    
    sDelSQL = objWsSql.SqlDeleteLAB350(Trim(txtStcd.Text))

'On Error GoTo DBExecError
    
'    DBConn.BeginTrans
    DBConn.Execute (sDelSQL)
    
    sInsSql2 = objWsSql.SqlInsertLAB350(Trim(txtStcd.Text), "0", Trim(txtGnm.Text))
    DBConn.Execute (sInsSql2)
'    DBConn.CommitTrans
    Set objWsSql = Nothing
    Exit Sub

'DBExecError:
'    DBConn.RollbackTrans
'    Set objWsSql = Nothing
End Sub

Private Sub SaveRstType()
    Dim I%
    Dim sTestCd As String
    
    Call Delrsttype(txtStcd.Text)
    
    For I = 0 To lstSTest.ListCount - 1
        sTestCd = GetCodeStr(lstSTest.List(I))
        Call Setrsttype(sTestCd)
    Next I
End Sub

Private Sub SaveTmp()

    Dim sDelSQL As String
    Dim sInsSql As String
    Dim I%
    Dim objWsSql As New clsLISSqlMasters
    
    sDelSQL = objWsSql.SqlDeleteLAB036(Trim(txtStcd.Text))

'On Error GoTo DBExecError

 '   DbConn.BeginTrans
    DBConn.Execute (sDelSQL)
    
    Call SaveWStmpcheck
    For I = 1 To iTmpArrayCnt
        If aTmpvalue(I).sSaveDecision = 1 Then  ' case Save
            sInsSql = objSql.SqlInsertLAB036(Trim(txtStcd.Text), aTmpvalue(I).sTpcd, aTmpvalue(I).swsfg, _
                                             aTmpvalue(I).sTpnm, aTmpvalue(I).sTpdata)
            DBConn.Execute (sInsSql)
        End If
    Next I
    
    Set objWsSql = Nothing
  '  DbConn.CommitTrans
    
'DBExecError:
 '  DbConn.RollbackTrans
    
End Sub

Private Sub cmdSaveTmp_Click()
    Dim iTmpArrayindex As Integer
    Dim I%
    
    If txtTmpCode.Text = "" Then
        MsgBox "템플릿 코드값을 입력하시오"
        Exit Sub
    End If
    If iMode = 1 Then       ' 추가 Mode일 경우
        If chkDuplicateTmp(txtTmpCode.Text) = False Then     ' 중복이 아닌경우
            With spdTmpResult       ' 수정된 새로운 내용을 스프레드에 기입한다.
                .MaxRows = .MaxRows + 1
                
                
                
                iTmpArrayCnt = .MaxRows
                ReDim Preserve aTmpvalue(1 To iTmpArrayCnt)
                .Row = .MaxRows
                .Col = 1
                .Text = txtTmpCode.Text
                aTmpvalue(iTmpArrayCnt).sTpcd = txtTmpCode.Text
                
                .Col = 2
                .Text = txtTmpResultNm.Text
                aTmpvalue(iTmpArrayCnt).sTpnm = txtTmpResultNm.Text
                
                
                .Col = 3
                .CellType = CellTypeCheckBox
                .TypeCheckCenter = True
               
                Call SaveWStmpcheck
                aTmpvalue(iTmpArrayCnt).sTpdata = txtTmpData.Text
                aTmpvalue(iTmpArrayCnt).sSaveDecision = 1 ' save
                
                .SortBy = SortByRow
                .SortKey(1) = 1  '검사코드
        
                .SortKeyOrder(1) = SortKeyOrderAscending
        
                .Col = 1:   .Col2 = .MaxRows
                .Row = 2:  .Row2 = .MaxRows
                .Action = ActionSort
            End With
            Call ClearTmpFrame
            iMode = 0               ' Clear iMode variable

        Else                            ' 중복인 경우
            txtTmpCode.SetFocus
        End If
        
    ElseIf iMode = 2 Then   ' Edit Mode일 경우
        With spdTmpResult
            .Row = iPspdtmpresult
    
    
    
            iTmpArrayindex = findTmpArrayIndex(txtTmpCode.Text)
            If iTmpArrayindex = 0 Then Exit Sub
    
            
            .Col = 1                    ' 수정된 새로운 내용을 스프레드에 기입한다.
            .Text = txtTmpCode.Text
            aTmpvalue(iTmpArrayindex).sTpcd = txtTmpCode.Text
            
            .Col = 2
            .Text = txtTmpResultNm.Text
            aTmpvalue(iTmpArrayindex).sTpnm = txtTmpResultNm.Text
    
'            .Col = 3
'            .CellType = CellTypeCheckBox
'            .TypeCheckCenter = True
'            If .Value = True Then
'                aTmpvalue(iTmpArrayCnt).swsfg = "1"     ' checked
'            Else
'                aTmpvalue(iTmpArrayCnt).swsfg = "0"     ' unchecked
'            End If
'
            Call SaveWStmpcheck
            aTmpvalue(iTmpArrayindex).sTpdata = txtTmpData.Text
        
        End With
        Call ClearTmpFrame
        iMode = 0               ' Clear iMode variable
    End If
       
End Sub
Private Sub SaveWStmpcheck()
    Dim I%
    With spdTmpResult
        .Col = 3
        For I = 1 To .MaxRows           ' Clear wsfg
            
            aTmpvalue(I).swsfg = "0"
            .Row = I
            If .Value = True Then
                aTmpvalue(I).swsfg = "1"
            End If
        Next I
    End With
End Sub
Private Sub cmdTmpHelp_Click()
    
    If LoadlstToSTest = True Then    ' 아직 그룹화되지않은 검사항목이 있을경우
        fraToSTest.Visible = True
        Call SetfraToSTestPosition
        Call ClearspdToSTest
        spdToSTest.SetFocus
        spdToSTest.CursorStyle = CursorStyleArrow
    End If
    
End Sub
Private Sub Form_Load()
'    SetPosition 2, Me ' set size of window
    Me.Show
    Call Initialize
    
End Sub

Public Sub Initialize()

    Dim rschkST As DrRecordSet
    Dim sSQL1 As String
    Dim I%
    
    sSQL1 = objSql.SqlGetLAB350("", " where " & DBW("stseq = ", "0"))
    Set rschkST = OpenRecordSet(sSQL1)
    
    If rschkST.EOF = True Then ' 기타검사코드가 존재하지 않을 경우
        tabItem.Tabs(1).Caption = "New"
        txtStcd.SetFocus
    Else                        ' 기타검사코드가 존재할 경우
        rschkST.MoveFirst
        tabItem.Tabs(1).Caption = "New"
        For I = 1 To rschkST.RecordCount
            tabItem.Tabs.Add I + 1, Trim(rschkST.Fields("stcd").Value), Trim(rschkST.Fields("stitem").Value)
            rschkST.MoveNext
        Next I
    End If
    tabItem.Tabs(1).Selected = True
    txtStcd.SetFocus
    
    rschkST.RsClose
    Set rschkST = Nothing
End Sub

Private Function LoadlstToSTest() As Boolean

    Dim rsToSTest As DrRecordSet
    Dim sSQL As String
    Dim I%
    
    Set rsToSTest = OpenRecordSet(objSql.SqlLoadSpecialTest)
    If rsToSTest.EOF = True Then
        LoadlstToSTest = False
        Exit Function
    End If
    
    With spdToSTest
        .MaxRows = rsToSTest.RecordCount
        For I = 1 To rsToSTest.RecordCount
            .Row = I
            .Col = 2
            .Text = rsToSTest.Fields("testcd").Value
            
            .Col = 3
            .Text = rsToSTest.Fields("testnm").Value
            
            rsToSTest.MoveNext
        Next I
    End With
    
    rsToSTest.RsClose
    Set rsToSTest = Nothing
    
    LoadlstToSTest = True
End Function




Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
End Sub

'Private Sub spdNumResultItem_Click(ByVal Col As Long, ByVal Row As Long)
'    Call DspSelRow(spdNumResultItem, iPspdNumResultItem, Row)
'End Sub

Private Sub spdTmpResult_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row < 1 Then Exit Sub
    Call DspSelRow(spdTmpResult, iPspdtmpresult, Row)
    Call DspTmpInfo
    
    If Col = 3 Then         ' worksheet template check 일 경우
        Call DspWSTmpCheck(Row)
    End If
    
End Sub
Private Sub DspWSTmpCheck(selrow As Long)
    Dim I%
    
    With spdTmpResult
        .Col = 3
        For I = 1 To .MaxRows
            .Row = I
            If I <> selrow Then
               .Value = False
            End If
        Next I
    End With
End Sub
Private Sub DspTmpInfo()
    Dim sTpcd As String
    Dim sTmpCode As String, sTmpResultNm As String, sTmpData As String
    
    If txtStcd.Text = "" Then
        MsgBox " 코드이름이 존재하지 않습니다."
        Exit Sub
    End If
    
    If iPspdtmpresult < 1 Then ' case header selected
        Exit Sub
    End If
    
    iMode = 2                   ' Setting iMode variable


    
    With spdTmpResult
        .Row = iPspdtmpresult
        .Col = 1
        sTmpCode = .Text
    End With

    Call FillfraTmp(sTmpCode)

    txtTmpCode.Enabled = False
    txtTmpResultNm.Enabled = False
'    txtTmpCode.SetFocus
    
    

End Sub
Private Sub DspSelRow(spread As vaSpread, iPspread, ByVal LRow As Long)
    With spread
        If iPspread >= 1 Then     ' 이전 선택항목 색깔 복귀
            .Col = -1
            .Row = iPspread
            .ForeColor = RGB(0, 0, 0)
        End If
        .Col = -1                       ' 선택항목 색깔 표시
        .Row = LRow
        .ForeColor = RGB(0, 0, 255)
    End With
    
    iPspread = LRow

End Sub
Private Sub spdToSTest_Click(ByVal Col As Long, ByVal Row As Long)
        
    If Row < 1 Then Exit Sub     ' case header Click!!!
    
    With spdToSTest
        .Col = 1
        .Row = Row
        
        If .Value = 1 Then
            .Value = 0
        Else
            .Value = 1
        End If
    End With
    
End Sub

Private Sub tabItem_Click()
    
    Dim iTabindex As Integer

    txtStcd.Enabled = True: txtGnm.Enabled = True

    'set the spread text color back  to black
    spdToSTest.ForeColor = &H0&
    spdToSTest.ForeColor = &H0&
    
    Call ClearSelColor
    
    If tabItem.SelectedItem.Caption = "New" Then
        cmdDelGrp.Enabled = False
        Call InitTmpArray
        Call ClearAll
        cmdClear.Enabled = True
        txtStcd.SetFocus
    Else
        iTabindex = tabItem.SelectedItem.Index
        cmdDelGrp.Enabled = True
        DspForm (iTabindex)
        cmdClear.Enabled = False
        txtStcd.Enabled = False: txtGnm.Enabled = False
    End If

End Sub
Private Sub AddWSTmp()
    Call cmAddTmp_Click
    txtTmpCode.Text = "N0"
    txtTmpResultNm.Text = "WorkSheet Template"
    Call cmdSaveTmp_Click
    
End Sub
Private Sub ClearSelColor()
    
    With spdTmpResult
        .Col = -1
        .Row = iPspdtmpresult
        .ForeColor = RGB(0, 0, 0)
    End With
    

End Sub
Private Sub DspForm(iTabindex As Integer)
        
    Call clearPositionVar
    Call InitTmpArray
    Call ClearTmpFrame
    Call DspStcd(iTabindex)
    Call DsplstItem
    Call DspspdTmpResult
    
End Sub
Private Sub InitTmpArray()
    
    Erase aTmpvalue
    iTmpArrayCnt = 0
End Sub


Private Sub lstToSTest_Click()

End Sub

Private Sub txtGnm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdTmpHelp.SetFocus
    End If
End Sub

Private Sub txtStcd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub txtStcd_LostFocus()
    
    Dim sSQL As String
    Dim rsToStcd As DrRecordSet
    Dim I%
    
    If Trim(txtStcd.Text) = "" Then Exit Sub
    If tabItem.Tabs(1).Selected = True Then ' "new" tab을 선택한경우
        
        Set rsToStcd = OpenRecordSet(objSql.SqlGetStCd("2", Trim(txtStcd.Text)))
        If rsToStcd.EOF = True Then
'            Call AddWSTmp
            txtGnm.SetFocus
            Exit Sub
        End If
        
        For I = 1 To rsToStcd.RecordCount
            If Trim(txtStcd.Text) = Trim(rsToStcd.Fields("stcd").Value) Then
                MsgBox " 이미 존재하는 코드이름 입니다. "
                txtStcd.Text = ""
                txtStcd.SetFocus
                Exit Sub
            End If
            rsToStcd.MoveNext
        Next I
        
        rsToStcd.RsClose
        Set rsToStcd = Nothing
    
'        Call AddWSTmp
        txtGnm.SetFocus
    End If
    
    
End Sub

Private Sub txtTmpCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub txtTmpCode_LostFocus()
    If Trim(txtTmpCode.Text) = "" Then Exit Sub
    If chkDuplicateTmp(txtTmpCode.Text) = True Then     ' 중복이 아닌경우
        txtTmpCode.Text = ""
        txtTmpCode.SetFocus
    End If
End Sub

Private Sub txtTmpResultNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTmpData.SetFocus
    End If
End Sub

Private Sub txtStcd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtGnm.SetFocus
    End If
End Sub

Public Sub SetfraToSTestPosition()
    fraToSTest.Top = lstSTest.Top + 250
    fraToSTest.Left = lstSTest.Left + lstSTest.Width + 120
End Sub

Public Sub DspStcd(iTabindex As Integer)
    Dim rsStcd As DrRecordSet
    Dim sSQLGetstcd As String
    
    txtStcd.Text = ""
    txtGnm.Text = ""
    sSQLGetstcd = objSql.SqlGetStCd("2", "0", tabItem.Tabs(iTabindex).Key, tabItem.Tabs(iTabindex).Caption)
                    
    Set rsStcd = OpenRecordSet(sSQLGetstcd)
    
    If rsStcd.EOF = True Then Exit Sub
    
    txtStcd.Text = rsStcd.Fields("stcd").Value
    txtGnm.Text = tabItem.Tabs(iTabindex).Caption
    'disable textbox when DspForm
    txtStcd.Enabled = False: txtGnm.Enabled = False
    
    rsStcd.RsClose
    Set rsStcd = Nothing
End Sub




Public Sub DsplstItem()
    Dim rsSTItem As DrRecordSet
    Dim sSQL As String
    Dim I%
    
    Call ClearlstSTest
    
    sSQL = objSql.SqlLoadSpecialTest(Trim(txtStcd.Text))
            
    Set rsSTItem = OpenRecordSet(sSQL)
    
    If rsSTItem.EOF = True Then Exit Sub
    
    rsSTItem.MoveFirst
    
    For I = 0 To rsSTItem.RecordCount - 1
        lstSTest.AddItem rsSTItem.Fields("TestCd").Value & vbTab & _
                         rsSTItem.Fields("TestNm").Value, I
        rsSTItem.MoveNext
    Next I
    
    rsSTItem.RsClose
    Set rsSTItem = Nothing
    
End Sub


Public Sub DspspdTmpResult()
    
    Dim rsTmpResult As DrRecordSet
    Dim sSQL As String
    Dim I%
    
    Call ClearspdTmpResult
    
    sSQL = objSql.SqlGetLAB036(Trim(txtStcd.Text))
    Set rsTmpResult = OpenRecordSet(sSQL)
    
    If rsTmpResult.EOF = True Then Exit Sub
    
    With spdTmpResult
        .MaxRows = rsTmpResult.RecordCount
        iTmpArrayCnt = rsTmpResult.RecordCount
        ReDim aTmpvalue(1 To iTmpArrayCnt)
        For I = 1 To rsTmpResult.RecordCount
            .Row = I
            .Col = 1
            .Text = rsTmpResult.Fields("tpcd").Value
            
            aTmpvalue(I).sTpcd = rsTmpResult.Fields("tpcd").Value
            
            .Col = 2
            .Text = rsTmpResult.Fields("tpnm").Value
            aTmpvalue(I).sTpnm = rsTmpResult.Fields("tpnm").Value
            
            .Col = 3
            .CellType = CellTypeCheckBox
            .TypeCheckCenter = True
            If rsTmpResult.Fields("wsfg").Value = "1" Then      ' checked
                .Value = True
            End If
            
            aTmpvalue(I).sTpdata = rsTmpResult.Fields("tpdata").Value
            aTmpvalue(I).sSaveDecision = 1 ' save
            
            rsTmpResult.MoveNext
        Next I
        .RowHeight(-1) = 14
        
    End With
    
    rsTmpResult.RsClose
    Set rsTmpResult = Nothing
    
End Sub


Private Sub txtTmpCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtTmpResultNm.SetFocus
    End If
End Sub



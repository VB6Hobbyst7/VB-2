VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm364AppAnti 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   10905
   Begin VB.CommandButton cmdUP 
      BackColor       =   &H00FFF2EE&
      Caption         =   "△"
      Height          =   915
      Index           =   1
      Left            =   7290
      Style           =   1  '그래픽
      TabIndex        =   23
      Top             =   5175
      Width           =   375
   End
   Begin VB.CommandButton cmdDown 
      BackColor       =   &H00FFF2EE&
      Caption         =   "▽"
      Height          =   900
      Index           =   1
      Left            =   7290
      Style           =   1  '그래픽
      TabIndex        =   22
      Top             =   7005
      Width           =   375
   End
   Begin VB.CommandButton cmdUP 
      BackColor       =   &H00FFF2EE&
      Caption         =   "△"
      Height          =   915
      Index           =   0
      Left            =   7275
      Style           =   1  '그래픽
      TabIndex        =   21
      Top             =   1905
      Width           =   375
   End
   Begin VB.CommandButton cmdDown 
      BackColor       =   &H00FFF2EE&
      Caption         =   "▽"
      Height          =   900
      Index           =   0
      Left            =   7275
      Style           =   1  '그래픽
      TabIndex        =   20
      Top             =   3735
      Width           =   375
   End
   Begin VB.CommandButton cmdDellstMs 
      BackColor       =   &H00CDE7FA&
      Caption         =   ">>"
      Height          =   420
      Left            =   7290
      Style           =   1  '그래픽
      TabIndex        =   19
      Top             =   6570
      Width           =   375
   End
   Begin VB.CommandButton cmdAddMsAnti 
      BackColor       =   &H00CDE7FA&
      Caption         =   "<<"
      Height          =   420
      Left            =   7290
      Style           =   1  '그래픽
      TabIndex        =   18
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton cmdDellstGs 
      BackColor       =   &H00CDE7FA&
      Caption         =   ">>"
      Height          =   405
      Left            =   7275
      Style           =   1  '그래픽
      TabIndex        =   16
      Top             =   3300
      Width           =   375
   End
   Begin VB.CommandButton cmdAddGsAnti 
      BackColor       =   &H00CDE7FA&
      Caption         =   "<<"
      Height          =   420
      Left            =   7260
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   2850
      Width           =   390
   End
   Begin VB.ListBox lstTotalAnti 
      BackColor       =   &H00EFFEFE&
      Height          =   6780
      Left            =   7785
      Sorted          =   -1  'True
      Style           =   1  '확인란
      TabIndex        =   14
      Top             =   1335
      Width           =   3075
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   7305
      Left            =   3000
      TabIndex        =   8
      Top             =   810
      Width           =   4755
      Begin VB.ListBox LstMsAnti 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   2730
         Left            =   75
         TabIndex        =   11
         Top             =   4365
         Width           =   4185
      End
      Begin VB.ListBox LstGsAnti 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   2730
         Left            =   90
         TabIndex        =   9
         Top             =   1095
         Width           =   4170
      End
      Begin VB.Line Line1 
         X1              =   135
         X2              =   4605
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label lblSpe 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   225
         Width           =   4140
      End
      Begin VB.Label Label3 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Mic 감수성 적용 항생제"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   225
         TabIndex        =   12
         Top             =   4050
         Width           =   3120
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "일반 감수성 적용 항생제"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   10
         Top             =   750
         Width           =   2655
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   405
      Left            =   60
      TabIndex        =   7
      Top             =   915
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   714
      BackColor       =   13752531
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "균종 선택"
   End
   Begin VB.ListBox lstTotalSpe 
      BackColor       =   &H00FBEDEA&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1380
      Width           =   2805
   End
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   5535
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9525
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton CmdToDel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "삭제(&D)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8190
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   720
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   10770
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "균종별 항생제 등록"
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
         Left            =   3870
         TabIndex        =   1
         Top             =   240
         Width           =   3150
      End
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   405
      Left            =   7785
      TabIndex        =   17
      Top             =   900
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   714
      BackColor       =   13752531
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "항생제 리스트"
   End
End
Attribute VB_Name = "frm364AppAnti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objSql As New clsLISSqlStatement
Private dicAntiList As New clsDictionary


Private Sub cmdClear_Click()
    Dim i%
    
    lblSpe.Caption = ""
    LstGsAnti.Clear
    LstMsAnti.Clear
  '  lblAntiNm.Caption = ""
    
    Call ClearLstTotalAnti
    Call ClearLstTotalSpe
            
End Sub

Private Sub ClearLstTotalSpe()
    Dim i%
    
    For i = 0 To lstTotalSpe.ListCount - 1
        lstTotalSpe.Selected(i) = False
    Next i
End Sub

Private Sub cmdClearChk_Click()
    
    Call ClearLstTotalAnti
End Sub

Private Sub ClearLstTotalAnti()
    Dim i%
    For i = 0 To lstTotalAnti.ListCount - 1
        lstTotalAnti.Selected(i) = False
    Next i

End Sub

Private Sub cmdDellstGs_Click()
    Dim i%
    Dim iDelcount%, iselcount%
    
    iselcount = LstGsAnti.SelCount
    If LstGsAnti.ListCount <> 0 And LstGsAnti.SelCount <> 0 Then
        For i = 0 To LstGsAnti.ListCount - 1
            If LstGsAnti.Selected(i) = True Then    ' Case Selected
                LstGsAnti.RemoveItem i
                i = i - 1
                 iDelcount = iDelcount + 1
                If iDelcount = iselcount Then Exit Sub
            End If
            
        Next i
    End If
 
End Sub

Private Sub cmdDellstMs_Click()
    Dim i%
    Dim iDelcount%, iselcount%
    
    iselcount = LstMsAnti.SelCount
    If LstMsAnti.ListCount <> 0 And LstMsAnti.SelCount <> 0 Then
        For i = 0 To LstMsAnti.ListCount - 1
            If LstMsAnti.Selected(i) = True Then    ' Case Selected
                LstMsAnti.RemoveItem i
                i = i - 1
                 iDelcount = iDelcount + 1
                If iDelcount = iselcount Then Exit Sub
            End If
            
        Next i
    End If
End Sub

Private Sub cmdAddMsAnti_Click()
    Dim i%
    For i = 0 To lstTotalAnti.ListCount - 1
        If lstTotalAnti.Selected(i) Then
            If chkExistlstMs(lstTotalAnti.List(i)) = False Then      ' 이전에 존재하지 않는 항생제일 경우
                LstMsAnti.AddItem lstTotalAnti.List(i)
            End If
        End If
    Next i

    Call ClearLstTotalAnti
End Sub

Private Sub cmdAddGsAnti_Click()
    Dim i%
    For i = 0 To lstTotalAnti.ListCount - 1
        If lstTotalAnti.Selected(i) Then
            If chkExistlstGs(lstTotalAnti.List(i)) = False Then      ' 이전에 존재하지 않는 항생제일 경우
                LstGsAnti.AddItem lstTotalAnti.List(i)
            End If
        End If
    Next i
    
    Call ClearLstTotalAnti
    
End Sub

Private Function chkExistlstGs(sInsertAntiCd As String) As Boolean
    Dim i%
    For i = 0 To LstGsAnti.ListCount - 1
        If sInsertAntiCd = LstGsAnti.List(i) Then ' 이전에 값이 존재하면
            chkExistlstGs = True
            Exit Function
        End If
    Next i
    chkExistlstGs = False
End Function

Private Function chkExistlstMs(sInsertAntiCd As String) As Boolean
    Dim i%
    For i = 0 To LstMsAnti.ListCount - 1
        If sInsertAntiCd = LstMsAnti.List(i) Then ' 이전에 값이 존재하면
            chkExistlstMs = True
            Exit Function
        End If
    Next i
    chkExistlstMs = False
End Function

Private Sub cmdDown_Click(Index As Integer)
    
    Dim lngIndex As Long
    Dim strList As String
    
    Select Case Index
    Case 0
        If LstGsAnti.ListIndex < 0 Then Exit Sub
        lngIndex = LstGsAnti.ListIndex
        If lngIndex = LstGsAnti.ListCount - 1 Then Exit Sub
        strList = LstGsAnti.Text
        LstGsAnti.RemoveItem lngIndex
        LstGsAnti.AddItem strList, lngIndex + 1
        LstGsAnti.ListIndex = lngIndex + 1
        LstGsAnti.SetFocus
    Case 1
        If LstMsAnti.ListIndex < 0 Then Exit Sub
        lngIndex = LstMsAnti.ListIndex
        If lngIndex = LstMsAnti.ListCount - 1 Then Exit Sub
        strList = LstMsAnti.Text
        LstMsAnti.RemoveItem lngIndex
        LstMsAnti.AddItem strList, lngIndex + 1
        LstMsAnti.ListIndex = lngIndex + 1
        LstMsAnti.SetFocus
    End Select
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sSqlInC108Ms As String
    Dim sSqlInC108Gs As String
    Dim sMsText1 As String
    Dim sGsText1 As String
    Dim sAntiCd As String
    Dim sSpeCd As String
    Dim i%
    
    sMsText1 = CStr(LstMsAnti.ListCount)
    sGsText1 = CStr(LstGsAnti.ListCount)
    For i = 0 To LstMsAnti.ListCount - 1
        sAntiCd = medGetP(LstMsAnti.List(i), 1, vbTab)
        sMsText1 = sMsText1 & ";" & sAntiCd
'        sMsText1 = sMsText1 & ";" & LstMsAnti.List(I)
    Next i
    For i = 0 To LstGsAnti.ListCount - 1
        sAntiCd = medGetP(LstGsAnti.List(i), 1, vbTab)
        sGsText1 = sGsText1 & ";" & sAntiCd
        
        'sGsText1 = sGsText1 & ";" & LstGsAnti.List(I)
    Next i
    sSpeCd = Trim(medGetP(lstTotalSpe.Text, 1, vbTab))
    
    sSqlInC108Ms = objSql.SqlSaveLAB031(LC2_MicroAnti, sSpeCd, "MS", "", "", "", "", "", sMsText1, "", 1)
    sSqlInC108Gs = objSql.SqlSaveLAB031(LC2_MicroAnti, sSpeCd, "GS", "", "", "", "", "", sGsText1, "", 1)
                    
On Error GoTo DBExecError
    dbconn.BeginTrans
    
    Call DeleteSpeAnti(sSpeCd)
    dbconn.Execute (sSqlInC108Ms)
    dbconn.Execute (sSqlInC108Gs)
    
    dbconn.CommitTrans
    Call ReLoadForm
    Exit Sub

DBExecError:
    dbconn.RollbackTrans
    MsgBox Err.Description, vbExclamation
     
End Sub

Private Sub ReLoadForm()
    Call LoadLstTotalSpe
    Call LoadLstTotalAnti
    LstGsAnti.Clear
    LstMsAnti.Clear
    lblSpe.Caption = ""
   ' lblAntiNm.Caption = ""

End Sub

Private Sub CmdToDel_Click()
    Dim sMsg As String
    Dim sRes As Integer, sStyle As Integer
    Dim sSpeCd As String
    
    sSpeCd = Trim(medGetP(lstTotalSpe.Text, 1, vbTab))
    If sSpeCd = "" Then Exit Sub

    sMsg = lblSpe.Caption & " 에 관한 정보를 모두 삭제합니다" & Chr$(13) & Chr$(10) & _
        "정말 삭제해도 좋습니까?"
    sStyle = vbYesNo + vbCritical + vbDefaultButton2
    
    sRes = MsgBox(sMsg, sStyle, "삭제 확인")
    If sRes = vbYes Then
        Call DeleteSpeAnti(sSpeCd)
        Call ReLoadForm
'        medMain.stsBar.Panels(2).Text = "정상적으로 삭제 처리 되었습니다. 다음 작업을 처리하세요"
    Else
        Exit Sub
    End If
End Sub

Private Sub DeleteSpeAnti(sSpeCd As String)
    
    Dim sSqlDelC108 As String
    
    sSqlDelC108 = objSql.SqlDeleteLAB031(LC2_MicroAnti, sSpeCd)
                  
On Error GoTo Err_Trap
   
'    DBConn.BeginTrans
    dbconn.Execute (sSqlDelC108)
'    DBConn.CommitTrans
    
    Call ReLoadForm
    Exit Sub
   
Err_Trap:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdUP_Click(Index As Integer)
    
    Dim lngIndex As Long
    Dim strList As String
    
    Select Case Index
    Case 0
        If LstGsAnti.ListIndex < 0 Then Exit Sub
        lngIndex = LstGsAnti.ListIndex
        If lngIndex = 0 Then Exit Sub
        strList = LstGsAnti.Text
        LstGsAnti.RemoveItem lngIndex
        LstGsAnti.AddItem strList, lngIndex - 1
        LstGsAnti.ListIndex = lngIndex - 1
        LstGsAnti.SetFocus
    Case 1
        If LstMsAnti.ListIndex < 0 Then Exit Sub
        lngIndex = LstMsAnti.ListIndex
        If lngIndex = 0 Then Exit Sub
        strList = LstMsAnti.Text
        LstMsAnti.RemoveItem lngIndex
        LstMsAnti.AddItem strList, lngIndex - 1
        LstMsAnti.ListIndex = lngIndex - 1
        LstMsAnti.SetFocus
    End Select
    
End Sub

Private Sub Form_Load()
'    SetPosition 2, Me
    Call LoadLstTotalSpe
    Call LoadLstTotalAnti
    Call GetAntiNmDic
    
End Sub

'    sMicCd = lstMic.List(lstMic.ListIndex)
'
'    sqlMic = "SELECT text1 micnm FROM " & T_LAB032 & _
'             " WHERE cdindex='" & LC3_Microbe & "' AND cdval1='" & sMicCd & "'"
'    iMicCol = dsMic.OpenCursor(DbConn, sqlMic)

Private Sub GetAntiNmDic()

    Dim sSqlGetAntiNm As String
    Dim rsGetAntiNm As New Recordset
    Dim i As Long
    
    '항생제명 딕셔너리에 담기...
    dicAntiList.Clear
    dicAntiList.DeleteAll
    dicAntiList.FieldInialize "anticd", "antinm"
    
    sSqlGetAntiNm = objSql.SqlLAB032CodeList(LC3_AntiBiotic, "cdval1 anti, text1 antinm")
    rsGetAntiNm.Open sSqlGetAntiNm, dbconn
    
    With rsGetAntiNm
        For i = 1 To .RecordCount
            dicAntiList.AddNew "" & .Fields("anti").Value, "" & .Fields("antinm").Value
            'Debug.Print .Fields("anti").Value, "" & .Fields("antinm").Value
            .MoveNext
        Next
    End With
    
    Set rsGetAntiNm = Nothing
End Sub

Private Sub LoadLstTotalAnti()
    Dim sSqlGetTotalAnti As String
    Dim rsGetTotalAnti As Recordset
    Dim i%
    
    sSqlGetTotalAnti = objSql.SqlLAB032CodeList(LC3_AntiBiotic, "cdval1 as AntiCd , text1 as AntiNm")
    Set rsGetTotalAnti = New Recordset
    rsGetTotalAnti.Open sSqlGetTotalAnti, dbconn
    
    If rsGetTotalAnti.EOF = True Then Exit Sub
    
    lstTotalAnti.Clear
    
    For i = 1 To rsGetTotalAnti.RecordCount
        lstTotalAnti.AddItem "" & rsGetTotalAnti.Fields("AntiCd").Value & vbTab & _
                             "" & rsGetTotalAnti.Fields("AntiNm").Value
        rsGetTotalAnti.MoveNext
    Next i
    
    Set rsGetTotalAnti = Nothing
    
End Sub

Private Sub LoadLstTotalSpe()
    
    Dim sSqlGetTotalSpe As String
    Dim rsGetTotalSpe As Recordset
    Dim i%
    
    sSqlGetTotalSpe = objSql.SqlLAB032CodeList(LC3_Species, "cdval1 as spcCd, field1 as spcNm ")
    Set rsGetTotalSpe = New Recordset
    rsGetTotalSpe.Open sSqlGetTotalSpe, dbconn
    
    If rsGetTotalSpe.EOF = True Then Exit Sub
    
    lstTotalSpe.Clear
    For i = 1 To rsGetTotalSpe.RecordCount

        lstTotalSpe.AddItem "" & rsGetTotalSpe.Fields("spcCd").Value & vbTab & _
                            "" & rsGetTotalSpe.Fields("spcNm").Value
                            
        rsGetTotalSpe.MoveNext
    Next i
    
    Set rsGetTotalSpe = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
    Set dicAntiList = Nothing
End Sub



Private Sub lstTotalSpe_Click()
    Dim sSqlGetAnti As String
    Dim rsGetAnti As Recordset
    Dim i%, j%
    
    Dim sSpeCd As String
    Dim sSpeNm As String
    Dim sAntiCd As String
    Dim sAntiNm As String
    
    If lstTotalSpe.Text = "" Then Exit Sub
    
    sSpeCd = Trim(Mid(lstTotalSpe.Text, 1, _
                 InStr(1, lstTotalSpe.Text, vbTab) - 1))
    
    sSpeNm = Trim(Mid(lstTotalSpe.Text, InStr(1, lstTotalSpe.Text, vbTab) + 1, _
                  Len(lstTotalSpe.Text)))
    lblSpe.Caption = sSpeCd & "  ( " & sSpeNm & " )"
    
    sSqlGetAnti = objSql.SqlLAB031CodeList(LC2_MicroAnti, "cdval2 as MsGs , text1 as Anti", sSpeCd)
    Set rsGetAnti = New Recordset
    rsGetAnti.Open sSqlGetAnti, dbconn
    
    If rsGetAnti.EOF = True Then
        LstMsAnti.Clear
        LstGsAnti.Clear
        Exit Sub
    End If
    
    For i = 1 To rsGetAnti.RecordCount
        If "" & rsGetAnti.Fields("MsGs").Value = "MS" Then
            LstMsAnti.Clear
            For j = 1 To Val(medGetP("" & rsGetAnti.Fields("Anti").Value, 1, ";"))
                sAntiCd = Trim(medGetP("" & rsGetAnti.Fields("Anti").Value, j + 1, ";"))
                If dicAntiList.Exists(sAntiCd) Then
                    dicAntiList.KeyChange (sAntiCd)
                    sAntiNm = dicAntiList.Fields("antinm")
                    LstMsAnti.AddItem sAntiCd & vbTab & sAntiNm
                End If
'                sSqlGetAntiNm = objSql.SqlLAB032CodeList(LC3_AntiBiotic, "text1 as AntiNm ", sAntiCd)
'                Set rsGetAntiNm = OpenRecordSet(sSqlGetAntiNm)
'                If Not rsGetAntiNm.EOF Then
'                    sAntiNm = Trim("" & rsGetAntiNm.Fields("antinm").Value)
'                    LstMsAnti.AddItem sAntiCd & vbTab & sAntiNm
'                    rsGetAntiNm.RsClose
'                End If
                                        
           Next j
        ElseIf "" & rsGetAnti.Fields("MsGs").Value = "GS" Then
            LstGsAnti.Clear
            For j = 1 To Val(medGetP("" & rsGetAnti.Fields("Anti").Value, 1, ";"))
                sAntiCd = Trim(medGetP("" & rsGetAnti.Fields("Anti").Value, j + 1, ";"))
                If dicAntiList.Exists(sAntiCd) Then
                    dicAntiList.KeyChange (sAntiCd)
                    sAntiNm = dicAntiList.Fields("antinm")
                    LstGsAnti.AddItem sAntiCd & vbTab & sAntiNm
                End If
'                sSqlGetAntiNm = objSql.SqlLAB032CodeList(LC3_AntiBiotic, "text1 as AntiNm ", sAntiCd)
'                Set rsGetAntiNm = OpenRecordSet(sSqlGetAntiNm)
'                If Not rsGetAntiNm.EOF Then
'                    sAntiNm = Trim("" & rsGetAntiNm.Fields("Antinm").Value)
'                    LstGsAnti.AddItem sAntiCd & vbTab & sAntiNm
'                    rsGetAntiNm.RsClose
'                End If
            Next j
        Else
            LstMsAnti.Clear
            LstGsAnti.Clear
        End If
        rsGetAnti.MoveNext
    Next i
    
    Set rsGetAnti = Nothing
End Sub

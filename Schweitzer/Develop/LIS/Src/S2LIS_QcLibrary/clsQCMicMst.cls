VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQCMicrobe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objQcSql As New clsLISSqlQc

'----------------------------------------------------------------------------------
'1. 미생물 QC Master 관련 메쏘드
'----------------------------------------------------------------------------------

Public Sub LoadMicData(ByRef lstMicList As Object, ByRef lstMicIndex As Object)
    
    Dim iMicCol As Integer
    Dim objMicSql As New clsLISSqlMicRst
    Dim dsMic As New Recordset

    dsMic.Open objMicSql.SQLLoadMicrobe, dbconn

    lstMicList.Clear
    lstMicIndex.Clear

    While Not dsMic.EOF
        lstMicList.addItem Format(Trim("" & dsMic.Fields("cdval1").Value), "!@@@@@@") & Trim("" & dsMic.Fields("text1").Value)
        lstMicIndex.addItem Trim("" & dsMic.Fields("text1").Value)
        dsMic.MoveNext
    Wend
    
    Set dsMic = Nothing
    Set objMicSql = Nothing

End Sub


Public Sub LoadMicQcRstCode(ByRef tvwRstCode As Object, ByRef sComboString As String)
    
    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim strTmp As String
    
    Set Rs = New Recordset
    Rs.Open objQcSql.SqlMicQcRstCode, dbconn
    
    strTmp = "": sComboString = ""
    With tvwRstCode
       .Nodes.Clear
        
        While (Not Rs.EOF)
           
            If Rs.Fields("cdval1").Value & "" <> strTmp Then
                If strTmp = "" Then
                    Call .Nodes.Add(, , Rs.Fields("cdval1").Value & "", Rs.Fields("cdval1").Value & "")
                Else
                    Call .Nodes.Add(strTmp, 2, Rs.Fields("cdval1").Value & "", Rs.Fields("cdval1").Value & "")
                End If
                .Nodes(.Nodes.Count).Tag = "P"
                strTmp = Rs.Fields("cdval1").Value & ""
                sComboString = sComboString & strTmp & vbTab
            End If
            Call .Nodes.Add(strTmp, 4, strTmp & Rs.Fields("cdval2").Value & "", Rs.Fields("field1").Value & "")
            .Nodes(.Nodes.Count).Tag = "C"
            .Nodes(.Nodes.Count).EnsureVisible
            Rs.MoveNext
           
        Wend
    End With
    
    Set Rs = Nothing
    
End Sub


Public Sub LoadGroupCd(ByRef tabGroup As Object)

    Dim Rs As Recordset
    Dim SqlStmt As String
    
    Set Rs = New Recordset
    Rs.Open objQcSql.SqlMicQcRstGroup, dbconn
    
    tabGroup.Tabs.Clear
    While (Not Rs.EOF)
        tabGroup.Tabs.Add , Rs.Fields("CdVal1").Value & "", Rs.Fields("Field1").Value & ""
        tabGroup.Tabs(tabGroup.Tabs.Count).Tag = "OLD"
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    tabGroup.Tabs.Add , "MQC" & (tabGroup.Tabs.Count + 1), "New"
    tabGroup.Tabs(tabGroup.Tabs.Count).Tag = "NEW"
    
End Sub

Public Function SaveMicQcMst(ByVal tmpGroupKey As String, ByVal tmpGroupNm As String, _
                             ByVal tblResultSet As Object, ByVal pExpFg As String) As Boolean
    
    Dim i As Integer
    Dim sCdVal1 As String, sCdVal2 As String
    Dim SqlStmt() As String
    
    '그룹코드 : MQC1, MQC2, MQC3...
    With tblResultSet
    
        ReDim SqlStmt(.MaxRows + 1)
        
        sCdVal1 = tmpGroupKey
        sCdVal2 = "*"
        
        '원래 있던 데이타 삭제..
        SqlStmt(0) = objQcSql.SqlDeleteMicQcRstGroup(sCdVal1)
                               
        
        'Column Header 부분 (그룹명, 사용여부, 시약명들..) : cdval2 = '*' 인 데이타.
        .Col = 2: .Col2 = .MaxCols: .Row = 0: .Row2 = 0
        .BlockMode = True
        SqlStmt(1) = objQcSql.SqlInsertMicQcRstGroup(sCdVal1, sCdVal2, tmpGroupNm, _
                                                     pExpFg, "0", CStr(.MaxCols - 1), .Clip)
        .BlockMode = False
        
        For i = 1 To .MaxRows
            .Row = i: .Col = 1: sCdVal2 = .Value

            '각 표준균주별 시약에 대한 결과코드...
            .Col = 2: .Col2 = .MaxCols: .Row = i: .Row2 = i
            .BlockMode = True
            SqlStmt(i + 1) = objQcSql.SqlInsertMicQcRstGroup(sCdVal1, sCdVal2, "", "", CStr(i), "", .Clip)
            .BlockMode = False
        Next
        
    End With

On Error GoTo Err_Trap

    dbconn.BeginTrans
    For i = 0 To UBound(SqlStmt)
        dbconn.Execute SqlStmt(i)
    Next
    dbconn.CommitTrans
    
    SaveMicQcMst = True
    Exit Function

Err_Trap:
    SaveMicQcMst = False
    MsgBox Err.Description, vbExclamation
End Function


'----------------------------------------------------------------------------------
'2. 미생물 QC 결과등록 관련 메쏘드
'----------------------------------------------------------------------------------

Public Function FindMicQcDate(ByVal strGroupCd As String, ByRef lstRstDate As Object) As Boolean
    
    Dim Rs As Recordset
    
    Set Rs = New Recordset
    Rs.Open objQcSql.SqlFindMicQcDate(strGroupCd), dbconn
    
    lstRstDate.Clear
    While (Not Rs.EOF)
        lstRstDate.addItem Format(Rs.Fields("RstDt").Value & "", CS_DateMask)
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
End Function


Public Function LastDataExist(ByVal strGroupCd As String, ByRef strRstDt As String, _
                              ByVal strQcKind As String) As Boolean

    Dim Rs As Recordset
    
    Set Rs = New Recordset
    Rs.Open objQcSql.SqlLastDataExist(strGroupCd, strRstDt, strQcKind), dbconn
    
    If Rs.Fields("RstDate").Value & "" = "" Then
        LastDataExist = False
    Else
        LastDataExist = True
    End If

    Set Rs = Nothing
    
End Function


Public Function ShowLastResult(ByVal pRstDate As String, ByVal pTempKey As String, _
                               ByRef tblResult As Object, ByRef pRstNm As String, _
                               ByRef pRemark As String) As Boolean

    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim iRow As Integer
    Dim i As Integer
    Dim blnRemark As Boolean
    
    Set Rs = New Recordset
    Rs.Open objQcSql.SqlLastDataShow(pTempKey, pRstDate), dbconn
    
    If Not Rs.EOF Then
        pRstNm = Rs.Fields("EmpNm").Value & ""
        With tblResult
            .ReDraw = False
            .MaxRows = 0
            .MaxRows = (Rs.RecordCount - 1) * 2
            iRow = 0
            While (Not Rs.EOF)
                If Rs.Fields("MicCd").Value & "" = "HEADER" Then    'Column Header(시약)
                    .MaxCols = Val(medGetP(Rs.Fields("RstText").Value & "", 1, ":")) + 1
                    .Row = 0: .Row2 = 0: .Col = 2: .Col2 = .MaxCols
                    .BlockMode = True
                    .Clip = medGetP(Rs.Fields("RstText").Value & "", 2, ":")
                    .BlockMode = False
                ElseIf Rs.Fields("MicCd").Value & "" = "REMARK" Then    'Remark
                    pRemark = Rs.Fields("RstText").Value & ""
                Else    'Row Header(균주명)
                    .Row = Val(Rs.Fields("Seq").Value & "") * 2 - 1: .Row2 = .Row
                    .Col = 0: .Value = Rs.Fields("MicNm").Value & "": .TypeHAlign = 0    'TypeHAlignLeft
                    .Col = 1: .Value = Rs.Fields("MicCd").Value & ""
                    '결과
                    .Col = 2: .Col2 = .MaxCols
                    .BlockMode = True
                    .Clip = Rs.Fields("RstText").Value & ""
                    .BlockMode = False
                    '결과Template코드
                    .Row = Val(Rs.Fields("Seq").Value & "") * 2: .Row2 = .Row
                    .Col = 2: .Col2 = .MaxCols
                    .BlockMode = True
                    .Clip = Rs.Fields("RstTemp").Value & ""
                    .BlockMode = False
                    .RowHidden = True
                End If
                Rs.MoveNext
            Wend
            .ReDraw = True
        End With
        ShowLastResult = True
    Else
        ShowLastResult = False
    End If
    Set Rs = Nothing
    
End Function



Public Sub LoadRstTemp(ByRef mMicRst As Collection)

    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim MyRst As clsQCMicRst
    Dim sGrpCd As String
    Dim blnFirst As Boolean
    
    Set mMicRst = New Collection
    
    Set Rs = New Recordset
    Rs.Open objQcSql.SqlMicQcRstCode, dbconn
    
    sGrpCd = "": blnFirst = True
    While (Not Rs.EOF)
        If sGrpCd <> Trim(Rs.Fields("CdVal1").Value & "") Then
            
            If Not blnFirst Then Call mMicRst.Add(MyRst, sGrpCd)
                
            Set MyRst = New clsQCMicRst
            MyRst.RstCd = Rs.Fields("CdVal1").Value & ""
            MyRst.RstString = " "
            MyRst.RstCnt = 0
            sGrpCd = MyRst.RstCd
        End If
        MyRst.RstString = MyRst.RstString & vbTab & Rs.Fields("Field1").Value & ""
        MyRst.RstCnt = MyRst.RstCnt + 1
        blnFirst = False
        Rs.MoveNext
    Wend
    Call mMicRst.Add(MyRst, sGrpCd)
    
    Set Rs = Nothing
    
End Sub

Public Sub ShowTemplate(ByVal pTempKey As String, _
                        ByRef tblResult As Object)

    Dim Rs As Recordset
    Dim iRow As Integer
    Dim i As Integer
    
    Set Rs = New Recordset
    Rs.Open objQcSql.SqlTemplateShow(pTempKey), dbconn
    
    If Not Rs.EOF Then
        
        With tblResult
            .ReDraw = False
            .MaxRows = 0
            .MaxRows = (Rs.RecordCount - 1) * 2
            iRow = 0
            While (Not Rs.EOF)
                If Rs.Fields("CdVal2").Value & "" = "*" Then    'Column Header(시약)
                    .MaxCols = Val(Rs.Fields("Field4").Value & "") + 1
                    .Row = 0: .Row2 = 0: .Col = 2: .Col2 = .MaxCols
                    .BlockMode = True
                    .Clip = Trim(Rs.Fields("Text1").Value & "")
                    .BlockMode = False
                Else    'Row Header(균주명)
                    .Row = Val(Rs.Fields("Field3").Value & "") * 2 - 1
                    .Col = 0: .Value = Rs.Fields("MicNm").Value & "": .TypeHAlign = 0    'TypeHAlignLeft
                    .Col = 1: .Value = Rs.Fields("CdVal2").Value & ""
                    '결과Template코드
                    .Row = Val(Rs.Fields("Field3").Value & "") * 2: .Row2 = .Row
                    .Col = 2: .Col2 = .MaxCols
                    .BlockMode = True
                    .Clip = Rs.Fields("Text1").Value & ""
                    .RowHidden = True
                    .BlockMode = False
                End If
                Rs.MoveNext
            Wend
            .ReDraw = True
        End With
        
    End If
    
    Set Rs = Nothing
        
End Sub


Public Function SaveMicQcResult(ByVal pGroupKey As String, ByVal pRstDate As String, _
                                ByVal pEmpId As String, ByVal pRstText As String, _
                                ByVal tblResult As Object) As Boolean

    Dim i As Integer, iSeq As Integer
    Dim sKey1 As String, sKey2 As String
    Dim SqlStmt() As String
    
    '그룹코드 : MQC1, MQC2, MQC3...
    With tblResult
    
        ReDim SqlStmt((.MaxRows / 2) + 2)
        
        sKey1 = pGroupKey
        
        '원래 있던 데이타 삭제..
        SqlStmt(0) = objQcSql.SqlDeleteMicQcResult(pGroupKey, pRstDate)
        
        'Column Header 부분 (그룹명, 사용여부, 시약명들..) : miccd = 'HEADER' 인 데이타.
        sKey2 = "HEADER"
        .Col = 2: .Col2 = .MaxCols: .Row = 0: .Row2 = 0
        .BlockMode = True
        SqlStmt(1) = objQcSql.SqlInsertMicQcResult(pGroupKey, pRstDate, 0, sKey2, pEmpId, CStr(.MaxCols - 1) & ":" & .Clip)
        .BlockMode = False
        
        'Remark : miccd = 'REMARK' 인 데이타.
        If Trim(pRstText) <> "" Then
            sKey2 = "REMARK"
            SqlStmt(2) = objQcSql.SqlInsertMicQcResult(pGroupKey, pRstDate, 0, sKey2, pEmpId, pRstText)
        Else
            SqlStmt(2) = ""
        End If
        
        iSeq = 0
        For i = 1 To .MaxRows Step 2
            .Row = i: .Col = 1: sKey2 = .Value
            iSeq = iSeq + 1

            '각 표준균주별 시약에 대한 결과...
            .Col = 2: .Col2 = .MaxCols: .Row = i: .Row2 = i
            .BlockMode = True
            SqlStmt(iSeq + 2) = objQcSql.SqlInsertMicQcResult(pGroupKey, pRstDate, iSeq, sKey2, pEmpId, .Clip)
            .BlockMode = False
        Next
        
    End With
    
On Error GoTo Err_Trap

    dbconn.BeginTrans
    For i = 0 To UBound(SqlStmt)
'        Debug.Print SqlStmt(i)
        If SqlStmt(i) <> "" Then dbconn.Execute SqlStmt(i)
    Next
    dbconn.CommitTrans
    SaveMicQcResult = True
    Exit Function
    
Err_Trap:
    SaveMicQcResult = False
    MsgBox Err.Description, vbExclamation

End Function


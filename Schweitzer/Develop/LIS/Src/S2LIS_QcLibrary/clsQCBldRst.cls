VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQCBldRst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objBldSql As New clsLISSqlQc

Public Function FindBldQcDate(ByRef lstRstDate As Object) As Boolean
    
    Dim Rs As Recordset
    
    Set Rs = New Recordset
    Rs.Open objBldSql.SqlFindBldQcDate, DBConn
    
    lstRstDate.Clear
    While (Not Rs.EOF)
        lstRstDate.addItem Format(Rs.Fields("RstDt").Value, CS_DateMask)
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
End Function

Public Function ShowResult(ByVal pRstDate As String, ByRef pRstNm As String, _
                           ByVal tblResult As Object, _
                           ByVal intComboCnt As Long, _
                           ByVal strComboString As String, _
                           Optional ByVal pTempFg As Boolean = False) As Boolean

    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim iRow As Integer
    Dim i As Integer
    
    Set Rs = New Recordset
    Rs.Open objBldSql.SqlGetBldQcResult(pRstDate), DBConn
    
    If Not Rs.EOF Then
        pRstNm = Rs.Fields("EmpNm").Value & ""
        With tblResult
            .ReDraw = False
            .MaxRows = 0
            .MaxRows = Rs.RecordCount
            iRow = 0
            While (Not Rs.EOF)
                .Row = Val(Rs.Fields("Seq").Value & ""): .Row2 = .Row
                '결과
                .Col = 1: .Col2 = 6
                .BlockMode = True
                .CellType = 5   'CellTypeStaticText
                .TypeHAlign = 2 'TypeHAlignCenter
                .Clip = Rs.Fields("Title").Value & ""
                .BlockMode = False
                If pTempFg Then
                    .Col = 7: .Value = ""
                    .Col = 8: .Value = ""
                Else
                    .Col = 7: .Value = Rs.Fields("RstCd").Value & ""
                    .Col = 8: .Value = Rs.Fields("Remark").Value & ""
                End If
                Rs.MoveNext
            Wend
            .Row = 1: .Row2 = .MaxRows
            .Col = 7: .Col2 = 7
            .BlockMode = True
            If Trim(strComboString) = "" Then
                .CellType = 1   'CellTypeEdit
            Else
                .CellType = 8   'CellTypeComboBox
                .TypeComboBoxList = strComboString
                .TypeComboBoxMaxDrop = intComboCnt
            End If
            .ForeColor = &HB9602F
            .BlockMode = False
            
            .ReDraw = True
        End With
        ShowResult = True
    Else
        ShowResult = False
    End If
    Set Rs = Nothing
    
End Function


Public Function SaveBldQcResult(ByVal pRstDate As String, ByVal tblResult As Object, ByVal pEmpId As String) As Boolean

    Dim i As Integer
    Dim SqlStmt() As String
    Dim sRstCd As String
    Dim sRemark As String
    
    With tblResult
    
        ReDim SqlStmt(.MaxRows + 1)
        
        '원래 있던 데이타 삭제..
        SqlStmt(0) = objBldSql.SqlDeleteBldQcResult(pRstDate)
        
        For i = 1 To .MaxRows
            .Row = i
            '각 시약에 대한 결과...
            .Col = 7: sRstCd = .Text
            .Col = 8: sRemark = .Text
            .Col = 1: .Col2 = 6: .Row = i: .Row2 = i
            .BlockMode = True
            SqlStmt(i) = objBldSql.SqlInsertBldQcResult(pRstDate, i, pEmpId, sRstCd, .Clip, sRemark)
            SqlStmt(i) = SqlStmt(i) & DBV("title", .Clip, 1)
            .BlockMode = False
        Next
        
    End With
    
    DBConn.BeginTrans
    For i = 0 To UBound(SqlStmt)
        If SqlStmt(i) <> "" Then DBConn.Execute SqlStmt(i)
    Next
    DBConn.CommitTrans
    
    SaveBldQcResult = True
    Exit Function

Err_Trap:
    MsgBox Err.Description, vbExclamation
    SaveBldQcResult = False

End Function


Public Function ShowTemplateDate(ByVal pRstDt As String) As String

    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim sRstNm As String
    
    Set Rs = New Recordset
    Rs.Open objBldSql.SqlLastDataExist("", pRstDt, "2"), DBConn
    ShowTemplateDate = Rs.Fields("RstDate").Value & ""
    
    Set Rs = Nothing
    
End Function

Public Sub LoadRstTemp(ByRef intComboCnt As Long, ByRef strComboString As String)

    Dim Rs As Recordset
    
    Set Rs = New Recordset
    Rs.Open objBldSql.SqlGetBldTemplate, DBConn
    
    strComboString = " "
    intComboCnt = 1
    While (Not Rs.EOF)
        strComboString = strComboString & vbTab & Rs.Fields("Field1").Value
        intComboCnt = intComboCnt + 1
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
End Sub




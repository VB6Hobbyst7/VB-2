VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCS0201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private recordFiled01 As String
Private recordFiled02 As String
Private recordFiled03 As String
Private recordFiled04 As String
Private recordFiled05 As String
Private recordFiled06 As String
Private recordFiled07 As String
Private recordFiled08 As String
Private recordFiled09 As String

Private count As Integer        'query 결과 레코드 개수

Dim AdoConn     As New ADODB.Connection

Public Property Let LetrecordFiled01(ByVal data As String)
    recordFiled01 = data
End Property

Public Property Get GetrecordField01() As String
    GetrecordField01 = recordFiled01
End Property

Public Property Let LetrecordFiled02(ByVal data As String)
    recordFiled02 = data
End Property

Public Property Get GetrecordFiled02() As String
    GetrecordFiled02 = recordFiled02
End Property

Public Property Let LetrecordFiled03(ByVal data As String)
    recordFiled03 = data
End Property

Public Property Get GetrecordFiled03() As String
    GetrecordFiled03 = recordFiled03
End Property

Public Property Let LetrecordFiled04(ByVal data As String)
    recordFiled04 = data
End Property

Public Property Get GetrecordFiled04() As String
    GetrecordFiled04 = recordFiled04
End Property

Public Property Let LetrecordFiled05(ByVal data As String)
    recordFiled05 = data
End Property

Public Property Get GetrecordFiled05() As String
    GetrecordFiled05 = recordFiled05
End Property

Public Property Let LetrecordFiled06(ByVal data As String)
    recordFiled06 = data
End Property

Public Property Get GetrecordFiled06() As String
    GetrecordFiled06 = recordFiled06
End Property

Public Property Let LetrecordFiled07(ByVal data As String)
    recordFiled07 = data
End Property

Public Property Get GetrecordFiled07() As String
    GetrecordFiled07 = recordFiled07
End Property

Public Property Let LetrecordFiled08(ByVal data As String)
    recordFiled08 = data
End Property

Public Property Get GetrecordFiled08() As String
    GetrecordFiled08 = recordFiled08
End Property

Public Property Let LetrecordFiled09(ByVal data As String)
    recordFiled09 = data
End Property

Public Property Get GetrecordFiled09() As String
    GetrecordFiled09 = recordFiled09
End Property

Public Property Let Letcount(ByVal data As Integer)
    count = data
End Property

Public Property Get Getcount() As Integer
    Getcount = count
End Property

Public Sub Get_SlipNm(cd As String)

    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
        
    sSql = "SELECT SLIPNM" _
        & " FROM PART" _
        & " WHERE PARTCD = '" & Left$(cd, 1) & "'" _
        & " AND PARTGBN = '" & Right$(cd, 2) & "'"
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        AdoConn.Close
        count = 0
        Exit Sub
    Else
        
        count = 1
        
        AdoConn.MoveFirst
        
        recordFiled01 = AdoConn("SLIPNM")
        
    End If
    
    AdoConn.Close
    
    Exit Sub
ErrHandler:
    MsgBox "Get_SlipNm(AS02.DLL) - " & Err.Description

End Sub

Public Sub Get_RegNo(regNo As String, flag As Boolean)

    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    
    If flag Then
        sSql = "SELECT NAME, SEX, BORNYEAR" _
            & " FROM PERSON" _
            & " WHERE REGNO = '" & regNo & "'"
    Else
        sSql = "SELECT REGNO, SEX, BORNYEAR" _
            & " FROM PERSON" _
            & " WHERE NAME = '" & regNo & "'"
    End If

    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        AdoConn.Close
        count = 0
        Exit Sub
    Else
        
        AdoConn.MoveFirst
        
        count = 1
        recordFiled01 = AdoConn(0)
        recordFiled02 = AdoConn(1)
        recordFiled03 = AdoConn(2)
    
    End If
    
    AdoConn.Close
    
    Exit Sub
ErrHandler:
    MsgBox "Get_RegNo(AS02.DLL) - " & Err.Description

End Sub

Public Sub Get_Name(name As String)
    
    On Error GoTo ErrHandler

    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim j%
        
    sSql = "SELECT REGNO" _
        & " FROM PERSON" _
        & " WHERE NAME = '" & Trim(name) & "'"
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"

    count = AdoConn.RecordCount
    If count = 0 Then
        AdoConn.Close
        count = 0
        Exit Sub
    End If

    AdoConn.MoveFirst

    For j = 1 To count
        recordFiled01 = recordFiled01 & AdoConn(0) & "|"

        AdoConn.MoveNext
    Next

    AdoConn.Close
    
    Exit Sub
ErrHandler:
    MsgBox "Get_Name(AS02.DLL) - " & Err.Description

End Sub
Public Sub Get_Contents(regNo As String, _
                        sDate As String, _
                        eDate As String, _
                        slipcd As String)

    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim j As Integer
    Dim sTestNmCfg As String
    
    sTestNmCfg = fGetCurTestItemNmCfg
    
    If sTestNmCfg = "T" Then
        sSql = "SELECT RESULT.LABDATE, RESULT.LABSEQ, JUBSU.JUBSUGBN, JUBSU.EMERGBN, DEPTNM," _
              & " WARD, DOCTOR, TESTITEMNM, RESULT" _
              & " FROM TESTITEM INNER JOIN ((H_JUBSU AS JUBSU" _
              & " INNER JOIN H_RESULT AS RESULT ON" _
              & " (JUBSU.LABSEQ = RESULT.LABSEQ) AND" _
              & " (RESULT.PARTGBN = JUBSU.PARTGBN) AND" _
              & " (JUBSU.LABDATE = RESULT.LABDATE)) LEFT JOIN DEPT ON" _
              & " JUBSU.DEPTCD = DEPT.DEPTCD) ON" _
              & " (TESTITEM.SUBMCD = RESULT.SUBMCD) AND" _
              & " (TESTITEM.TESTITEMSEQ = RESULT.TESTITEMSEQ) AND" _
              & " (TESTITEM.SPECIMENCD = RESULT.SPECIMENCD) AND" _
              & " (TESTITEM.PARTGBN = RESULT.PARTGBN)" _
              & " WHERE RESULT.REGNO = '" & regNo & "'" _
              & " AND RESULT.LABDATE BETWEEN '" & sDate & "'" _
              & " AND '" & eDate & "'" _
              & " AND TESTITEM.PARTCD = '" & Left(slipcd, 1) & "'" _
              & " ORDER BY RESULT.LABDATE, RESULT.LABSEQ, TESTITEM.DISPLAYORDER"
    ElseIf sTestNmCfg = "P" Then
                sSql = "SELECT RESULT.LABDATE, RESULT.LABSEQ, JUBSU.JUBSUGBN, JUBSU.EMERGBN, DEPTNM," _
              & " WARD, DOCTOR, PRINTNM, RESULT" _
              & " FROM TESTITEM INNER JOIN ((H_JUBSU AS JUBSU" _
              & " INNER JOIN H_RESULT AS RESULT ON" _
              & " (JUBSU.LABSEQ = RESULT.LABSEQ) AND" _
              & " (RESULT.PARTGBN = JUBSU.PARTGBN) AND" _
              & " (JUBSU.LABDATE = RESULT.LABDATE)) LEFT JOIN DEPT ON" _
              & " JUBSU.DEPTCD = DEPT.DEPTCD) ON" _
              & " (TESTITEM.SUBMCD = RESULT.SUBMCD) AND" _
              & " (TESTITEM.TESTITEMSEQ = RESULT.TESTITEMSEQ) AND" _
              & " (TESTITEM.SPECIMENCD = RESULT.SPECIMENCD) AND" _
              & " (TESTITEM.PARTGBN = RESULT.PARTGBN)" _
              & " WHERE RESULT.REGNO = '" & regNo & "'" _
              & " AND RESULT.LABDATE BETWEEN '" & sDate & "'" _
              & " AND '" & eDate & "'" _
              & " AND TESTITEM.PARTCD = '" & Left(slipcd, 1) & "'" _
              & " ORDER BY RESULT.LABDATE, RESULT.LABSEQ, TESTITEM.DISPLAYORDER"
    End If
   
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"

    If AdoConn.RecordCount = 0 Then
        AdoConn.Close
        count = 0
        Exit Sub
    Else

        count = AdoConn.RecordCount

        AdoConn.MoveFirst

        For j = 1 To count
            recordFiled01 = recordFiled01 & AdoConn(0) & "|"
            recordFiled02 = recordFiled02 & AdoConn(1) & "|"
            recordFiled03 = recordFiled03 & AdoConn(2) & "|"
            recordFiled04 = recordFiled04 & AdoConn(3) & "|"
            recordFiled05 = recordFiled05 & AdoConn(4) & "|"
            recordFiled06 = recordFiled06 & AdoConn(5) & "|"
            recordFiled07 = recordFiled07 & AdoConn(6) & "|"
            recordFiled08 = recordFiled08 & AdoConn(7) & "|"
            recordFiled09 = recordFiled09 & AdoConn(8) & "|"

            AdoConn.MoveNext
        Next

    End If

    AdoConn.Close

    Exit Sub
ErrHandler:

    Select Case Err.Number
        Case -2147217865
            ViewMsg "해당 SLIP은 사용할 수 없습니다."
            Exit Sub
        Case Else
            MsgBox "Get_Contents(AS02.DLL) - " & Err.Description
    End Select
End Sub

Public Sub Get_SlipCD()
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim j As Integer

    sSql = "SELECT * FROM PART"

    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"

    If AdoConn.RecordCount = 0 Then
        AdoConn.Close
        count = 0
        Exit Sub
    Else

        count = AdoConn.RecordCount

        AdoConn.MoveFirst

        For j = 1 To count
            recordFiled01 = recordFiled01 & AdoConn("PARTCD") & "|"
            recordFiled02 = recordFiled02 & AdoConn("PARTGBN") & "|"
            recordFiled03 = recordFiled03 & AdoConn("SLIPNM") & "|"

            AdoConn.MoveNext
        Next

    End If

    AdoConn.Close

    Exit Sub
ErrHandler:
    MsgBox "Get_SlipCD(AS02.DLL) - " & Err.Description

End Sub

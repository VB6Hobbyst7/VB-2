VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCS0101"
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
Private recordFiled10 As String
Private recordFiled11 As String
Private recordFiled12 As String

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

Public Property Let LetrecordFiled10(ByVal data As String)
    recordFiled10 = data
End Property

Public Property Get GetrecordFiled10() As String
    GetrecordFiled10 = recordFiled10
End Property

Public Property Let LetrecordFiled11(ByVal data As String)
    recordFiled11 = data
End Property

Public Property Get GetrecordFiled11() As String
    GetrecordFiled11 = recordFiled11
End Property

Public Property Let LetrecordFiled12(ByVal data As String)
    recordFiled12 = data
End Property

Public Property Get GetrecordFiled12() As String
    GetrecordFiled12 = recordFiled12
End Property

Public Property Let Letcount(ByVal data As Integer)
    count = data
End Property

Public Property Get Getcount() As Integer
    Getcount = count
End Property

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
    MsgBox "Get_SlipCD(AS01.DLL) - " & Err.Description

End Sub

Public Sub Get_SlipName(cd As String)

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
    MsgBox "Get_SlipName(AS01.DLL) - " & Err.Description

End Sub

Public Sub Get_Contents(slipcd As String, _
                        sDate As String, _
                        eDate As String, _
                        sTime As String, _
                        eTime As String)

    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim j As Integer

    sSql = "SELECT LABDATE, LABSEQ, JUBSU.REGNO," _
          & " JUBSU.NAME, AGE, JUBSU.SEX, JUBSUGBN, EMERGBN," _
          & " DEPTNM, WARD, DOCTOR, MEDICALREMARK" _
          & " FROM (H_JUBSU AS JUBSU" _
          & " LEFT JOIN DEPT ON JUBSU.DEPTCD = DEPT.DEPTCD)" _
          & " INNER JOIN PERSON ON JUBSU.REGNO = PERSON.REGNO" _
          & " WHERE JUBSU.PARTGBN = '" & Right(slipcd, 2) & "'" _
          & " AND JUBSU.LABDATE BETWEEN '" & sDate & "'" _
          & " AND '" & eDate & "'" _
          & " AND JUBSU.LABTIME BETWEEN '" & sTime & "00'" _
          & " AND '" & eTime & "59'" _
          & " ORDER BY JUBSU.LABDATE, JUBSU.LABSEQ"

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
            recordFiled10 = recordFiled10 & AdoConn(9) & "|"
            recordFiled11 = recordFiled11 & AdoConn(10) & "|"
            recordFiled12 = recordFiled12 & AdoConn(11) & "|"

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
            MsgBox "Get_Contents(AS01.DLL) - " & Err.Description
    End Select
End Sub

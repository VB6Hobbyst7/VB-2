VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCB0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarTotField01 As String '지역 복사
Private mvarTotField02 As String '지역 복사
Private mvarTotField03 As String '지역 복사
Private mvarTotField04 As String '지역 복사
Private mvarTotField05 As String '지역 복사
Private mvarTotField06 As String '지역 복사
Private mvarTotField07 As String '지역 복사
Private mvarTotField08 As String '지역 복사
Private mvarTotField09 As String '지역 복사
Private mvarTotField10 As String '지역 복사
Private mvarTotField11 As String '지역 복사
Private mvarTotField12 As String '지역 복사
Private mvarTotField13 As String '지역 복사
Private mvarTotField14 As String '지역 복사
Private mvarTotField15 As String '지역 복사
Private mvarTotField16 As String '지역 복사
Private mvarTotField17 As String '지역 복사
Private mvarTotField18 As String '지역 복사
Private mvarTotField19 As String '지역 복사
Private mvarTotField20 As String '지역 복사
Private mvarTotField21 As String '지역 복사
Private mvarTotField22 As String '지역 복사
Private mvarTotField23 As String '지역 복사
Private mvarTotField24 As String '지역 복사
Private mvarTotField25 As String '지역 복사
Private mvarTotField26 As String '지역 복사
Private mvarTotField27 As String '지역 복사
Private mvarTotField28 As String '지역 복사
Private mvarTotField29 As String '지역 복사
Private mvarTotField30 As String '지역 복사
Private mvarTotField31 As String '지역 복사
Private mvarTotField32 As String '지역 복사
Private mvarTotField33 As String '지역 복사
Private mvarTotField34 As String '지역 복사
Private mvarTotField35 As String '지역 복사

Private mvarErrNum  As Long
Private mCurItemCnt As Integer

Private Function GetByOne(ByVal tStr As String, sOriginal As String) As String
    Dim Pos%
    
    Pos = InStr(tStr, "|")
    
    If Pos = 0 Then
    Else
        GetByOne = Trim$(Mid$(tStr, 1, Pos - 1))
        sOriginal = Trim$(Mid$(sOriginal, Pos + 1, Len(sOriginal) - Pos))
    End If
End Function

Public Property Let TotField35(ByVal vData As String)
    mvarTotField35 = vData
End Property

Public Property Get TotField35() As String
    TotField35 = mvarTotField35
End Property

Public Property Let TotField34(ByVal vData As String)
    mvarTotField34 = vData
End Property

Public Property Get TotField34() As String
    TotField34 = mvarTotField34
End Property

Public Property Let TotField33(ByVal vData As String)
    mvarTotField33 = vData
End Property

Public Property Get TotField33() As String
    TotField33 = mvarTotField33
End Property

Public Property Let TotField32(ByVal vData As String)
    mvarTotField32 = vData
End Property

Public Property Get TotField32() As String
    TotField32 = mvarTotField32
End Property

Public Property Let TotField31(ByVal vData As String)
    mvarTotField31 = vData
End Property

Public Property Get TotField31() As String
    TotField31 = mvarTotField31
End Property

Public Property Let TotField30(ByVal vData As String)
    mvarTotField30 = vData
End Property

Public Property Get TotField30() As String
    TotField30 = mvarTotField30
End Property

Public Property Let TotField29(ByVal vData As String)
    mvarTotField29 = vData
End Property

Public Property Get TotField29() As String
    TotField29 = mvarTotField29
End Property

Public Property Let TotField28(ByVal vData As String)
    mvarTotField28 = vData
End Property

Public Property Get TotField28() As String
    TotField28 = mvarTotField28
End Property

Public Property Let TotField27(ByVal vData As String)
    mvarTotField27 = vData
End Property

Public Property Get TotField27() As String
    TotField27 = mvarTotField27
End Property

Public Property Let TotField26(ByVal vData As String)
    mvarTotField26 = vData
End Property

Public Property Get TotField26() As String
    TotField26 = mvarTotField26
End Property

Public Property Let TotField25(ByVal vData As String)
    mvarTotField25 = vData
End Property

Public Property Get TotField25() As String
    TotField25 = mvarTotField25
End Property

Public Property Let TotField24(ByVal vData As String)
    mvarTotField24 = vData
End Property

Public Property Get TotField24() As String
    TotField24 = mvarTotField24
End Property

Public Property Let TotField23(ByVal vData As String)
    mvarTotField23 = vData
End Property

Public Property Get TotField23() As String
    TotField23 = mvarTotField23
End Property

Public Property Let TotField22(ByVal vData As String)
    mvarTotField22 = vData
End Property

Public Property Get TotField22() As String
    TotField22 = mvarTotField22
End Property

Public Property Let TotField21(ByVal vData As String)
    mvarTotField21 = vData
End Property

Public Property Get TotField21() As String
    TotField21 = mvarTotField21
End Property

Public Property Let TotField20(ByVal vData As String)
    mvarTotField20 = vData
End Property

Public Property Get TotField20() As String
    TotField20 = mvarTotField20
End Property

Public Property Let TotField19(ByVal vData As String)
    mvarTotField19 = vData
End Property

Public Property Get TotField19() As String
    TotField19 = mvarTotField19
End Property

Public Property Let TotField18(ByVal vData As String)
    mvarTotField18 = vData
End Property

Public Property Get TotField18() As String
    TotField18 = mvarTotField18
End Property

Public Property Let TotField17(ByVal vData As String)
    mvarTotField17 = vData
End Property

Public Property Get TotField17() As String
    TotField17 = mvarTotField17
End Property

Public Property Let TotField16(ByVal vData As String)
    mvarTotField16 = vData
End Property

Public Property Get TotField16() As String
    TotField16 = mvarTotField16
End Property

Public Property Let TotField15(ByVal vData As String)
    mvarTotField15 = vData
End Property

Public Property Get TotField15() As String
    TotField15 = mvarTotField15
End Property

Public Property Let TotField14(ByVal vData As String)
    mvarTotField14 = vData
End Property

Public Property Get TotField14() As String
    TotField14 = mvarTotField14
End Property

Public Property Let TotField13(ByVal vData As String)
    mvarTotField13 = vData
End Property

Public Property Get TotField13() As String
    TotField13 = mvarTotField13
End Property

Public Property Let TotField12(ByVal vData As String)
    mvarTotField12 = vData
End Property

Public Property Get TotField12() As String
    TotField12 = mvarTotField12
End Property

Public Property Let TotField11(ByVal vData As String)
    mvarTotField11 = vData
End Property

Public Property Get TotField11() As String
    TotField11 = mvarTotField11
End Property

Public Property Let TotField10(ByVal vData As String)
    mvarTotField10 = vData
End Property

Public Property Get TotField10() As String
    TotField10 = mvarTotField10
End Property

Public Property Let TotField09(ByVal vData As String)
    mvarTotField09 = vData
End Property

Public Property Get TotField09() As String
    TotField09 = mvarTotField09
End Property

Public Property Let TotField08(ByVal vData As String)
    mvarTotField08 = vData
End Property

Public Property Get TotField08() As String
    TotField08 = mvarTotField08
End Property

Public Property Let TotField07(ByVal vData As String)
    mvarTotField07 = vData
End Property

Public Property Get TotField07() As String
    TotField07 = mvarTotField07
End Property

Public Property Let TotField06(ByVal vData As String)
    mvarTotField06 = vData
End Property

Public Property Get TotField06() As String
    TotField06 = mvarTotField06
End Property

Public Property Let TotField05(ByVal vData As String)
    mvarTotField05 = vData
End Property

Public Property Get TotField05() As String
    TotField05 = mvarTotField05
End Property

Public Property Let TotField04(ByVal vData As String)
    mvarTotField04 = vData
End Property

Public Property Get TotField04() As String
    TotField04 = mvarTotField04
End Property

Public Property Let TotField03(ByVal vData As String)
    mvarTotField03 = vData
End Property

Public Property Get TotField03() As String
    TotField03 = mvarTotField03
End Property

Public Property Let TotField02(ByVal vData As String)
    mvarTotField02 = vData
End Property

Public Property Get TotField02() As String
    TotField02 = mvarTotField02
End Property

Public Property Let TotField01(ByVal vData As String)
    mvarTotField01 = vData
End Property

Public Property Get TotField01() As String
    TotField01 = mvarTotField01
End Property

Public Property Get CurItemCnt() As Integer
    CurItemCnt = mCurItemCnt
End Property

Public Property Get AdoErrNum() As Long
    AdoErrNum = mvarErrNum
End Property

Public Sub Add_PART(ByVal sPart As String, ByVal sPartGbn As String, ByVal sSlipNm As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sPart = Trim$(sPart)
    sPartGbn = Trim$(sPartGbn)
    sSlipNm = Trim$(sSlipNm)
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "INSERT INTO PART VALUES (" & _
           "'" & sPart & "'," & " " & "'" & sPartGbn & "'," & " " & "'" & sSlipNm & "')"

    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Add_PART(AB01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Edit_PART(ByVal sPart As String, ByVal sPartGbn As String, ByVal sSlipNm As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sPart = Trim$(sPart)
    sPartGbn = Trim$(sPartGbn)
    sSlipNm = Trim$(sSlipNm)
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    'Slip명 Edit
    sSql = " UPDATE PART SET SLIPNM = " & "'" & sSlipNm & "'" & _
           " WHERE PARTCD = " & "'" & sPart & "'" & " AND PARTGBN = " & _
           "'" & sPartGbn & "' "
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Edit_PART(AB01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Delete_PART(ByVal sPartCd As String, ByVal sPartGbn As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sPartCd = Trim$(sPartCd)
    sPartGbn = Trim$(sPartGbn)
    
    sSql = "DELETE FROM Part WHERE PARTCD = " & _
           "'" & sPartCd & "'" & " AND " & _
           "PARTGBN = " & "'" & sPartGbn & "'" & " "
        
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Delete_PART(AB01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Get_PART(Optional ByVal sField01 As String, Optional ByVal sField02 As String)
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    
    If sField01 <> "" And sField02 <> "" Then
        sSql = "SELECT PARTCD, PARTGBN, SLIPNM FROM PART " & _
               "WHERE PARTCD = " & "'" & sField01 & "' " & _
               "AND PARTGBN = " & "'" & sField02 & "' "
    ElseIf sField01 <> "" Then
        sSql = "SELECT PARTCD, PARTGBN, SLIPNM FROM PART " & _
               "WHERE PARTCD = " & "'" & sField01 & "' "
    Else
        sSql = "SELECT PARTCD, PARTGBN, SLIPNM FROM PART ORDER BY PARTCD, PARTGBN"
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            mvarTotField01 = mvarTotField01 & AdoRecord("PARTCD") & "|"
            mvarTotField02 = mvarTotField02 & AdoRecord("PARTGBN") & "|"
            mvarTotField03 = mvarTotField03 & AdoRecord("SLIPNM") & "|"
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub

ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_PART(AB01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub

Public Sub Get_TESTITEMORD(ByVal sPartCd As String, ByVal sPartGbn As String, ByVal sSpc As String)
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sPartCd = Trim$(sPartCd)
    sPartGbn = Trim$(sPartGbn)
    sSpc = Trim$(sSpc)
    
    sSql = "SELECT TOP 1 DISPLAYORDER FROM TESTITEM " & _
           "WHERE PARTCD = " & "'" & sPartCd & "' " & _
           "AND PARTGBN = " & "'" & sPartGbn & "' " & _
           "AND SPECIMENCD = " & "'" & sSpc & "' " & _
           "ORDER BY DISPLAYORDER DESC"
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            mvarTotField01 = AdoRecord(0) + 1
            
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_TESTITEMORD(AB01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub

Public Sub Add_SPC(ByVal sField01 As String, ByVal sField02 As String, ByVal sField03 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField03)
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "INSERT INTO SPECIMEN VALUES (" & _
           "'" & sField01 & "'," & " " & "'" & sField02 & "'," & " " & "'" & sField03 & "')"

    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Add_SPC(AB01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Edit_SPC(ByVal sField01 As String, ByVal sField02 As String, ByVal sField03 As String, ByVal iEditField As Integer)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField02)
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    'SpecimenBriefNm Or SpecimenFullNm Edit
    sSql = " UPDATE SPECIMEN SET SPECIMENBRIEFNM = " & "'" & sField02 & "'" & _
            " , SPECIMENFULLNM = " & "'" & sField03 & "'" & _
           " WHERE SPECIMENCD = " & "'" & sField01 & "'"
   
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Edit_SPC(AB01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Delete_SPC(ByVal sField01 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    
    sSql = "DELETE FROM SPECIMEN WHERE SPECIMENCD = " & _
           "'" & sField01 & "'"
        
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Delete_SPC(AB01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Get_SPC(Optional ByVal sField01 As String)
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
        
    sField01 = Trim$(sField01)
        
    If sField01 <> "" Then
        sSql = "SELECT SPECIMENCD, SPECIMENBRIEFNM, SPECIMENFULLNM " & _
               "FROM SPECIMEN " & _
               "WHERE SPECIMENCD = " & "'" & sField01 & "' "
    Else
        sSql = "SELECT * FROM SPECIMEN ORDER BY SPECIMENCD"
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            mvarTotField01 = mvarTotField01 & AdoRecord("SPECIMENCD") & "|"
            mvarTotField02 = mvarTotField02 & AdoRecord("SPECIMENBRIEFNM") & "|"
            mvarTotField03 = mvarTotField03 & AdoRecord("SPECIMENFULLNM") & "|"
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_SPC(AB01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub

Public Sub Add_TESTITEM(ByVal iAddMode As Integer)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    If iAddMode = 1 Then
        sSql = "INSERT INTO TESTITEM VALUES (" & _
           "'" & mvarTotField01 & "', " & "'" & mvarTotField02 & "', " & _
           "'" & mvarTotField03 & "', " & "'" & mvarTotField04 & "', " & _
           "'" & mvarTotField05 & "', " & "'" & mvarTotField06 & "', " & _
           "'" & mvarTotField07 & "', " & "'" & mvarTotField08 & "', " & _
           "'" & mvarTotField09 & "', " & "'" & mvarTotField10 & "', " & _
           "'" & mvarTotField11 & "', " & "'" & mvarTotField12 & "', " & _
           "'" & mvarTotField13 & "', " & "'" & mvarTotField14 & "', " & _
           "'" & mvarTotField15 & "', " & "'" & mvarTotField16 & "', " & _
           "'" & mvarTotField17 & "', " & "'" & mvarTotField18 & "', " & _
           "'" & mvarTotField19 & "', " & "'" & mvarTotField20 & "', " & _
           "'" & mvarTotField21 & "', " & "'" & mvarTotField22 & "', " & _
           "'" & mvarTotField23 & "', " & "'" & mvarTotField24 & "', " & _
           "'" & mvarTotField25 & "', " & "'" & mvarTotField26 & "', " & _
           "'" & mvarTotField27 & "', " & "'" & mvarTotField28 & "', " & _
           "'" & mvarTotField29 & "', " & "'" & mvarTotField30 & "', " & _
           "'" & mvarTotField31 & "', " & "'" & mvarTotField32 & "', " & _
           "'" & mvarTotField33 & "', " & "'" & mvarTotField34 & "' ) "
    Else
    
    End If
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0  'Insert시
    
    Exit Sub
    
ErrHandler:

    AdoConn.Close
    
    Select Case Err.Number
        Case -2147217900
        
            mvarErrNum = 1  'Update시
            Call Edit_TESTITEM
            
        Case Else
            MsgBox "Add_TESTITEM(AB01.DLL) - " & Err.Description
            mvarErrNum = Err.Number
    End Select
    
End Sub

Public Sub Edit_TESTITEM()
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "UPDATE TESTITEM SET " & _
           "TESTITEMNM = " & "'" & mvarTotField07 & "', " & _
           "PRINTNM = " & "'" & mvarTotField08 & "', " & _
           "UNIT = " & "'" & mvarTotField09 & "', " & _
           "DISPLAYORDER = " & "'" & mvarTotField10 & "', " & _
           "PRINTORDER = " & "'" & mvarTotField11 & "', " & _
           "TAKEROOMYN = " & "'" & mvarTotField12 & "', " & _
           "REFGBN = " & "'" & mvarTotField13 & "', " & _
           "PANICGBN = " & "'" & mvarTotField14 & "', " & _
           "DELTAGBN = " & "'" & mvarTotField15 & "', " & _
           "DELTAVAL = " & "'" & mvarTotField16 & "', " & _
           "REGSERVERCD = " & "'" & mvarTotField17 & "', " & _
           "PANJUNGGBN = " & "'" & mvarTotField18 & "', " & _
           "PANICLOW = " & "'" & mvarTotField19 & "', " & _
           "PANICHIGH = " & "'" & mvarTotField20 & "', " & _
           "REFLOWM = " & "'" & mvarTotField21 & "', " & _
           "REFHIGHM = " & "'" & mvarTotField22 & "', " & _
           "UPPERLIMITM = " & "'" & mvarTotField23 & "', " & _
           "LOWERLIMITM = " & "'" & mvarTotField24 & "', " & _
           "GRAYUPPERM = " & "'" & mvarTotField25 & "', " & _
           "GRAYLOWERM = " & "'" & mvarTotField26 & "', " & _
           "OTHERFLAG = " & "'" & mvarTotField27 & "', " & _
           "REFLETTER = " & "'" & mvarTotField28 & "', "
    
    sSql = sSql & _
           "REFLOWF = " & "'" & mvarTotField29 & "', " & _
           "REFHIGHF = " & "'" & mvarTotField30 & "', " & _
           "UPPERLIMITF = " & "'" & mvarTotField31 & "', " & _
           "LOWERLIMITF = " & "'" & mvarTotField32 & "', " & _
           "GRAYUPPERF = " & "'" & mvarTotField33 & "', " & _
           "GRAYLOWERF = " & "'" & mvarTotField34 & "' "
           
    sSql = sSql & _
           "WHERE PARTCD = " & "'" & mvarTotField01 & "' " & _
           "AND PARTGBN = " & "'" & mvarTotField02 & "' " & _
           "AND SPECIMENCD = " & "'" & mvarTotField03 & "' " & _
           "AND TESTITEMSEQ = " & "'" & mvarTotField04 & "' " & _
           "AND SUBMCD = " & "'" & mvarTotField05 & "' "

' 이 필드는 나중에 STATUS 테이블에 TESTITEM이 관계를 맺을 때 "FLAGYN = " & "'" & mvarTotField06 & "'
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    Exit Sub
    
ErrHandler:

    AdoConn.Close
    
    MsgBox Err.Number
    MsgBox "Edit_TESTITEM(AB01.DLL) - " & Err.Description
End Sub

Public Sub Delete_TESTITEM(ByVal sField01 As String, ByVal sField02 As String, _
                      ByVal sField03 As String, ByVal sField04 As String, _
                      ByVal sField05 As String)
                      
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField03)
    sField04 = Trim$(sField04)
    sField05 = Trim$(sField05)
    
    sSql = "DELETE FROM TESTITEM " & _
           "WHERE PARTCD = " & "'" & sField01 & "' " & _
           "AND PARTGBN = " & "'" & sField02 & "' " & _
           "AND SPECIMENCD = " & "'" & sField03 & "' " & _
           "AND TESTITEMSEQ = " & "'" & sField04 & "' " & _
           "AND SUBMCD = " & "'" & sField05 & "' "
           
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    Exit Sub
    
ErrHandler:
    MsgBox "Delete_TESTITEM(AB01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Get_TESTITEM(ByVal iGetItemMode As Integer, ByVal sField01 As String, ByVal sField02 As String, ByVal sField03 As String, Optional ByVal sField04 As String, Optional ByVal sField05 As String, Optional ByVal sField06 As String, Optional ByVal sField07 As String, Optional ByVal sField08 As String)
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField03)
    sField04 = Trim$(sField04)
    sField05 = Trim$(sField05)
    sField06 = Trim$(sField06)
    sField07 = Trim$(sField07)
    sField08 = Trim$(sField08)
    
    If iGetItemMode = 1 Then    '화면초기시 스프레드에 뿌릴 내용 가져오기
        sSql = "SELECT * "
        sSql = sSql & "From TESTITEM "
        sSql = sSql & "WHERE PARTCD = " & "'" & sField01 & "' " & _
                      "AND PARTGBN = " & "'" & sField02 & "' " & _
                      "AND SPECIMENCD = " & "'" & sField03 & "' "
        sSql = sSql & "ORDER BY PARTCD, PARTGBN, SPECIMENCD, DISPLAYORDER, TESTITEMSEQ"
        
    ElseIf iGetItemMode = 2 Then    '스프레드 클릭시 화면에 뿌릴 내용 가져오기
        sSql = "SELECT * FROM TESTITEM " & _
               "WHERE PARTCD = " & "'" & sField01 & "' " & _
               "AND PARTGBN = " & "'" & sField02 & "' " & _
               "AND SPECIMENCD = " & "'" & sField03 & "' " & _
               "AND TESTITEMSEQ = " & "'" & sField04 & "' " & _
               "AND SUBMCD = " & "'" & sField05 & "' " & _
               "ORDER BY PARTCD, PARTGBN, SPECIMENCD, DISPLAYORDER, TESTITEMSEQ"
               
    ElseIf iGetItemMode = 3 Then    'Default Cd 찾기
        sSql = "SELECT * "
        sSql = sSql & "From TESTITEM "
        sSql = sSql & "WHERE PARTCD = " & "'" & sField01 & "' " & _
                      "AND PARTGBN = " & "'" & sField02 & "' " & _
                      "AND SPECIMENCD = " & "'" & sField03 & "' " & _
                      "AND TESTITEMSEQ = " & "'" & sField04 & "' " & _
                      "AND ( SUBMCD = 'NNNN' " & _
                      "OR SUBMCD = '00NN' ) " & _
                      "ORDER BY PARTCD, PARTGBN, SPECIMENCD, DISPLAYORDER, TESTITEMSEQ "
                      
    ElseIf iGetItemMode = 4 Then    'With SubCd
        sSql = "SELECT * "
        sSql = sSql & "From TESTITEM "
        sSql = sSql & "WHERE PARTCD = " & "'" & sField01 & "' " & _
                      "AND PARTGBN = " & "'" & sField02 & "' " & _
                      "AND SPECIMENCD = " & "'" & sField03 & "' " & _
                      "AND TESTITEMSEQ = " & "'" & sField04 & "' " & _
                      "AND SUBMCD LIKE " & "'" & sField05 & "%' " & _
                      "ORDER BY PARTCD, PARTGBN, SPECIMENCD, DISPLAYORDER, TESTITEMSEQ"
                      
    ElseIf iGetItemMode = 5 Then    '마지막 Sub 검사 찾기
        sSql = "SELECT TOP 1 * "
        sSql = sSql & "From TESTITEM "
        sSql = sSql & "WHERE PARTCD = " & "'" & sField01 & "' " & _
                      "AND PARTGBN = " & "'" & sField02 & "' " & _
                      "AND SPECIMENCD = " & "'" & sField03 & "' " & _
                      "AND TESTITEMSEQ = " & "'" & sField04 & "' " & _
                      "ORDER BY SUBMCD DESC "
    
    ElseIf iGetItemMode = 6 Then    'WithOut Slip, All Search
        sSql = "SELECT * "
        sSql = sSql & "From TESTITEM "
        sSql = sSql & "ORDER BY PARTCD, PARTGBN, SPECIMENCD, DISPLAYORDER, TESTITEMSEQ"
        
    ElseIf iGetItemMode = 7 Then    'With Slip, Search
        sSql = "SELECT * "
        sSql = sSql & "From TESTITEM "
        sSql = sSql & "WHERE PARTCD = " & "'" & sField01 & "' " & _
                      "AND PARTGBN = " & "'" & sField02 & "' " & _
                      "ORDER BY PARTCD, PARTGBN, SPECIMENCD, DISPLAYORDER, TESTITEMSEQ"
                      
    ElseIf iGetItemMode = 8 Then    'With Slip, With Specimen
        'iGetItemMode = 1 과 같다.
    ElseIf iGetItemMode = 9 Then    'Normal 설정만 조회
        sSql = "SELECT * "
        sSql = sSql & "From TESTITEM "
        sSql = sSql & "WHERE PARTCD = " & "'" & sField01 & "' " & _
                      "AND PARTGBN = " & "'" & sField02 & "' " & _
                      "AND SPECIMENCD = " & "'" & sField03 & "' " & _
                      "AND TESTITEMSEQ = " & "'" & sField04 & "' " & _
                      "AND SUBMCD = " & "'NNNN' "
                      
    ElseIf iGetItemMode = 10 Then   'Sub 설정만 조회
        sSql = "SELECT * "
        sSql = sSql & "From TESTITEM "
        sSql = sSql & "WHERE PARTCD = " & "'" & sField01 & "' " & _
                      "AND PARTGBN = " & "'" & sField02 & "' " & _
                      "AND SPECIMENCD = " & "'" & sField03 & "' " & _
                      "AND TESTITEMSEQ = " & "'" & sField04 & "' " & _
                      "AND SUBMCD <> " & "'NNNN' "
                      
'''    ElseIf iGetItemMode = 11 Then   'Multi 설정만 조회
'''        sSql = "SELECT * "
'''        sSql = sSql & "From TestItem "
'''        sSql = sSql & "WHERE PartCd = " & "'" & sField01 & "' " & _
'''                      "AND PartGbn = " & "'" & sField02 & "' " & _
'''                      "AND SpecimenCd = " & "'" & sField03 & "' " & _
'''                      "AND TestItemSeq = " & "'" & sField04 & "' " & _
'''                      "AND SubCd = " & "'N' " & _
'''                      "AND MultiRefCd <> 'N'"
'''    ElseIf iGetItemMode = 12 Then   'Cal 설정만 조회
'''        sSql = "SELECT * "
'''        sSql = sSql & "From TestItem "
'''        sSql = sSql & "WHERE PartCd = " & "'" & sField01 & "' " & _
'''                      "AND PartGbn = " & "'" & sField02 & "' " & _
'''                      "AND SpecimenCd = " & "'" & sField03 & "' " & _
'''                      "AND TestItemSeq = " & "'" & sField04 & "' " & _
'''                      "AND SubCd = " & "'C' " & _
'''                      "AND MultiRefCd = 'C'"

    ElseIf iGetItemMode = 13 Then
        sSql = "SELECT DISTINCT PARTGBN, SPECIMENCD, TESTITEMSEQ, SUBMCD, TESTITEMNM " & _
               "FROM TESTITEM " & _
               "WHERE PARTCD = " & "'" & sField01 & "' "
        If Not sField02 = "" Then
            sSql = sSql & _
               "AND PARTGBN = " & "'" & sField02 & "' "
        End If
        If Not sField03 = "" Then
            sSql = sSql & _
               "AND SPECIMENCD = " & "'" & sField03 & "' "
        End If
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
        
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            If iGetItemMode = 1 Or iGetItemMode = 2 Or _
               iGetItemMode = 3 Or iGetItemMode = 4 Or _
               iGetItemMode = 5 Or iGetItemMode = 6 Or _
               iGetItemMode = 7 Or iGetItemMode = 8 Or _
               iGetItemMode = 9 Or iGetItemMode = 10 Or _
               iGetItemMode = 11 Or iGetItemMode = 12 Then
               
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
                mvarTotField03 = mvarTotField03 & AdoRecord(2) & "|"
                mvarTotField04 = mvarTotField04 & AdoRecord(3) & "|"
                mvarTotField05 = mvarTotField05 & AdoRecord(4) & "|"
                mvarTotField06 = mvarTotField06 & AdoRecord(5) & "|"
                mvarTotField07 = mvarTotField07 & AdoRecord(6) & "|"
                mvarTotField08 = mvarTotField08 & AdoRecord(7) & "|"
                mvarTotField09 = mvarTotField09 & AdoRecord(8) & "|"
                mvarTotField10 = mvarTotField10 & AdoRecord(9) & "|"
                mvarTotField11 = mvarTotField11 & AdoRecord(10) & "|"
                mvarTotField12 = mvarTotField12 & AdoRecord(11) & "|"
                mvarTotField13 = mvarTotField13 & AdoRecord(12) & "|"
                mvarTotField14 = mvarTotField14 & AdoRecord(13) & "|"
                mvarTotField15 = mvarTotField15 & AdoRecord(14) & "|"
                mvarTotField16 = mvarTotField16 & AdoRecord(15) & "|"
                mvarTotField17 = mvarTotField17 & AdoRecord(16) & "|"
                mvarTotField18 = mvarTotField18 & AdoRecord(17) & "|"
                mvarTotField19 = mvarTotField19 & AdoRecord(18) & "|"
                mvarTotField20 = mvarTotField20 & AdoRecord(19) & "|"
                mvarTotField21 = mvarTotField21 & AdoRecord(20) & "|"
                mvarTotField22 = mvarTotField22 & AdoRecord(21) & "|"
                mvarTotField23 = mvarTotField23 & AdoRecord(22) & "|"
                mvarTotField24 = mvarTotField24 & AdoRecord(23) & "|"
                mvarTotField25 = mvarTotField25 & AdoRecord(24) & "|"
                mvarTotField26 = mvarTotField26 & AdoRecord(25) & "|"
                mvarTotField27 = mvarTotField27 & AdoRecord(26) & "|"
                mvarTotField28 = mvarTotField28 & AdoRecord(27) & "|"
                mvarTotField29 = mvarTotField29 & AdoRecord(28) & "|"
                mvarTotField30 = mvarTotField30 & AdoRecord(29) & "|"
                mvarTotField31 = mvarTotField31 & AdoRecord(30) & "|"
                mvarTotField32 = mvarTotField32 & AdoRecord(31) & "|"
                mvarTotField33 = mvarTotField33 & AdoRecord(32) & "|"
                mvarTotField34 = mvarTotField34 & AdoRecord(33) & "|"
                
            ElseIf iGetItemMode = 13 Then
                mvarTotField01 = mvarTotField01 & AdoRecord("PARTGBN") & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord("SPECIMENCD") & "|"
                mvarTotField03 = mvarTotField03 & AdoRecord("TESTITEMSEQ") & "|"
                mvarTotField04 = mvarTotField04 & AdoRecord("SUBMCD") & "|"
                mvarTotField05 = mvarTotField05 & AdoRecord("TESTITEMNM") & "|"
                
            End If
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_TESTITEM(AB01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub

Public Sub Add_RTN(ByVal sField01 As String, ByVal sField02 As String, _
                   ByVal sField03 As String, ByVal sField04 As String, _
                   ByVal sField05 As String, ByVal sField06 As String, _
                   ByVal sField07 As String, ByVal iRepeatCnt As Integer)
                   
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    Dim i%
        
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField03)
    sField04 = Trim$(sField04)
    sField05 = Trim$(sField05)
    sField06 = Trim$(sField06)
    sField07 = Trim$(sField07)
        
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    For i = 1 To iRepeatCnt
        sSql = "INSERT INTO ROUTINE " & _
               "VALUES ( '" & GetByOne(sField01, sField01) & "', " & "'" & GetByOne(sField02, sField02) & "', " & _
                        "'" & GetByOne(sField03, sField03) & "', " & "'" & GetByOne(sField04, sField04) & "', " & _
                        "'" & GetByOne(sField05, sField05) & "', " & "'" & GetByOne(sField06, sField06) & "', " & _
                        "'" & GetByOne(sField07, sField07) & "' " & _
                        ")"
    
        AdoConn.Execute sSql
    Next
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    
    MsgBox "Add_RTN(AB01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Edit_RTNNM(ByVal sField01 As String, ByVal sField02 As String, ByVal sField03 As String, ByVal iEditField As Integer)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField03)
        
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    If iEditField = 0 Then      'RoutinName 만 수정
        sSql = "UPDATE ROUTINE SET ROUTINENM = " & "'" & sField03 & "' " & _
               "WHERE PARTCD = " & "'" & sField01 & "' " & _
               "AND ROUTINECD = " & "'" & sField02 & "' "
    End If
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    
    MsgBox "Edit_RTNNM(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Delete_RTN(ByVal sField01 As String, ByVal sField02 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    
    sSql = "DELETE FROM ROUTINE " & _
           "WHERE EXISTS " & _
           "(SELECT * FROM ROUTINE " & _
           "WHERE PARTCD = " & "'" & sField01 & "' " & _
           "AND ROUTINECD = " & "'" & sField02 & "') " & _
           "AND PARTCD = " & "'" & sField01 & "' " & _
           "AND ROUTINECD = " & "'" & sField02 & "' "
        
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    
    MsgBox "Delete_RTN(AO1.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Get_RTN(ByVal iMode As Integer, ByVal sField01 As String, ByVal sField02 As String)
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)

    If iMode = 0 Then   '초기화 Default Part
        sSql = "SELECT DISTINCT ROUTINECD, ROUTINENM FROM ROUTINE " & _
               "WHERE PARTCD = " & "'" & sField01 & "' "
               
    ElseIf iMode = 1 Then   'All Search
        sSql = "SELECT DISTINCT PARTCD, ROUTINECD, ROUTINENM FROM ROUTINE "
       
    ElseIf iMode = 2 Then
        sSql = "SELECT TESTITEM.PARTGBN, TESTITEM.SPECIMENCD, TESTITEM.TESTITEMSEQ, " & _
                      "TESTITEM.SUBMCD, MAX(TESTITEM.TESTITEMNM) AS TESTNM " & _
               "FROM TESTITEM INNER JOIN ROUTINE " & _
               "ON (TESTITEM.SUBMCD = ROUTINE.SUBMCD) " & _
               "AND (TESTITEM.TESTITEMSEQ = ROUTINE.TESTITEMSEQ) " & _
               "AND (TESTITEM.PARTCD = ROUTINE.PARTCD) " & _
               "AND (TESTITEM.PARTGBN = ROUTINE.PARTGBN) " & _
               "AND (TESTITEM.SPECIMENCD = ROUTINE.SPECIMENCD) " & _
               "WHERE (ROUTINE.PARTCD) = " & "'" & sField01 & "' " & _
               "AND (ROUTINE.ROUTINECD) = " & "'" & sField02 & "' " & _
               "GROUP BY TESTITEM.PARTGBN, TESTITEM.SPECIMENCD, TESTITEM.TESTITEMSEQ, TESTITEM.SUBMCD "
               
    ElseIf iMode = 3 Then
        sSql = "SELECT DISTINCT ROUTINENM FROM ROUTINE " & _
               "WHERE PARTCD = " & "'" & sField01 & "' " & _
               "AND ROUTINECD = " & "'" & sField02 & "' "
    ElseIf iMode = 4 Then
        sSql = "SELECT DISTINCT ROUTINECD, ROUTINENM FROM ROUTINE " & _
               "WHERE PARTCD = " & "'" & sField01 & "' " & _
               "  AND SPECIMENCD = '" & sField02 & "' "
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            If iMode = 0 Then
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
            End If
            
            If iMode = 1 Then
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
                mvarTotField03 = mvarTotField03 & AdoRecord(2) & "|"
            End If
            
            If iMode = 2 Then
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
                mvarTotField03 = mvarTotField03 & AdoRecord(2) & "|"
                mvarTotField04 = mvarTotField04 & AdoRecord(3) & "|"
                mvarTotField05 = mvarTotField05 & AdoRecord(4) & "|"
                
            End If
            
            If iMode = 3 Then
                mvarTotField01 = mvarTotField01 & AdoRecord(0)
            End If
            
            If iMode = 4 Then
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
            End If
            
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_RTN(A01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub

Public Sub Add_DEPT(ByVal sField01 As String, ByVal sField02 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
   
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "INSERT INTO DEPT VALUES " & _
           "( " & "'" & sField01 & "', " & "'" & sField02 & "' )"

    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Add_DEPT(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Edit_DEPT(ByVal sField01 As String, ByVal sField02 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
   
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "UPDATE DEPT SET DEPTNM = " & "'" & sField02 & "' " & _
           "WHERE DEPTCD = " & "'" & sField01 & "' "
   
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Edit_DEPT(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Delete_DEPT(ByVal sField01 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
        
    sField01 = Trim$(sField01)
   
    sSql = "DELETE FROM DEPT WHERE DEPTCD = " & _
           "'" & sField01 & "' "
        
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Delete_DEPT(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Get_DEPT(Optional ByVal sField01 As String)
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
        
    sField01 = Trim$(sField01)
    
    If sField01 <> "" Then
        sSql = "SELECT DEPTNM FROM DEPT " & _
               "WHERE DEPTCD = " & "'" & sField01 & "' " & _
               "ORDER BY DEPTCD"
    Else
        sSql = "SELECT DEPTCD, DEPTNM FROM DEPT ORDER BY DEPTCD"
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            If sField01 <> "" Then
                mvarTotField01 = mvarTotField01 & AdoRecord(0)
            Else
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
            End If
            
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_DEPT(A01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub

Public Sub Add_USER(ByVal sField01 As String, ByVal sField02 As String, _
                    ByVal sField03 As String, ByVal sField04 As String, _
                    ByVal sField05 As String, ByVal sField06 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
        
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField03)
    sField04 = Trim$(sField04)
    sField05 = Trim$(sField05)
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "INSERT INTO USER VALUES " & _
           "( " & _
           "'" & sField01 & "', " & "'" & sField02 & "', " & _
           "'" & sField03 & "', " & "'" & sField04 & "', " & _
           "'" & sField05 & "', " & "'" & sField06 & "' " & _
           ")"

    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Add_USER(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Edit_USER(ByVal sField01 As String, ByVal sField02 As String, _
                    ByVal sField03 As String, ByVal sField04 As String, _
                    ByVal sField05 As String, ByVal sField06 As String)
                    
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
        
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField03)
    sField04 = Trim$(sField04)
    sField05 = Trim$(sField05)
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "UPDATE USER SET " & _
           "PASSWORD = " & "'" & sField02 & "', " & _
           "USERNM = " & "'" & sField03 & "', " & _
           "SLIPCD = " & "'" & sField04 & "', " & _
           "SPECIMENCD = " & "'" & sField05 & "', " & _
           "SCHOPT = " & "'" & sField06 & "' " & _
           "WHERE USERCD = " & "'" & sField01 & "' "
   
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Edit_USER(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Delete_USER(ByVal sField01 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
        
    sField01 = Trim$(sField01)
    
    sSql = "DELETE FROM USER WHERE USERCD = " & _
           "'" & sField01 & "' "
        
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Delete_USER(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub


Public Sub Get_USER(Optional ByVal sUser As String)
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i%, j%
    
    sUser = Trim$(sUser)
    
    If sUser <> "" Then
        sSql = "SELECT PASSWORD, USERNM, SLIPCD, SPECIMENCD, SCHOPT " & _
               "FROM USER " & _
               "WHERE USERCD = " & "'" & sUser & "' "
    Else
        sSql = "SELECT * " & _
               "FROM USER "
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        AdoRecord.Close
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            If sUser <> "" Then
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
                mvarTotField03 = mvarTotField03 & AdoRecord(2) & "|"
                mvarTotField04 = mvarTotField04 & AdoRecord(3) & "|"
                mvarTotField05 = mvarTotField05 & AdoRecord(4) & "|"
            Else
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
                mvarTotField03 = mvarTotField03 & AdoRecord(2) & "|"
                mvarTotField04 = mvarTotField04 & AdoRecord(3) & "|"
                mvarTotField05 = mvarTotField05 & AdoRecord(4) & "|"
                mvarTotField06 = mvarTotField06 & AdoRecord(5) & "|"
            End If
                 
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_USER(AB01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub


Public Sub Add_FLAG(ByVal sField01 As String, ByVal sField02 As String, _
                    ByVal sField03 As String, ByVal sField04 As String, _
                    ByVal sField05 As String, ByVal sField06 As String, _
                    ByVal sField07 As String, ByVal sField08 As String, _
                    ByVal sField09 As String, ByVal sField10 As String)
                    
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField03)
    sField04 = Trim$(sField04)
    sField05 = Trim$(sField05)
    sField06 = Trim$(sField06)
    sField07 = Trim$(sField07)
    sField08 = Trim$(sField08)
    sField09 = Trim$(sField09)
    sField10 = Trim$(sField10)

    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "INSERT INTO STATUS VALUES " & _
           "( " & _
           "'" & sField01 & "', " & "'" & sField02 & "', " & _
           "'" & sField03 & "', " & "'" & sField04 & "', " & _
           "'" & sField05 & "', " & "'" & sField06 & "', " & _
           "'" & sField07 & "', " & "'" & sField08 & "', " & _
           "'" & sField09 & "', " & "'" & sField10 & "' " & _
           ")"

    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    AdoConn.Close
    
    Select Case Err.Number
        Case -2147217900
        
            mvarErrNum = 1  'Update시
            Call Edit_FLAG(sField01, sField02, sField03, sField04, sField05, _
                        sField06, sField07, sField08, sField09, sField10)
            
        Case Else
            MsgBox "Add_FLAG(AB01.DLL) - " & Err.Description
            mvarErrNum = Err.Number
    End Select
End Sub

Public Sub Edit_FLAG(ByVal sField01 As String, ByVal sField02 As String, _
                    ByVal sField03 As String, ByVal sField04 As String, _
                    ByVal sField05 As String, ByVal sField06 As String, _
                    ByVal sField07 As String, ByVal sField08 As String, _
                    ByVal sField09 As String, ByVal sField10 As String)
                    
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField03)
    sField04 = Trim$(sField04)
    sField05 = Trim$(sField05)
    sField06 = Trim$(sField06)
    sField07 = Trim$(sField07)
    sField08 = Trim$(sField08)
    sField09 = Trim$(sField09)
    sField10 = Trim$(sField10)

    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "UPDATE STATUS SET " & _
           "PARTCD = " & "'" & sField02 & "', " & _
           "PARTGBN = " & "'" & sField03 & "', " & _
           "SPECIMENCD = " & "'" & sField04 & "', " & _
           "TESTITEMSEQ = " & "'" & sField05 & "', " & _
           "SUBMCD = " & "'" & sField06 & "', " & _
           "CRITICALVAL = " & "'" & sField07 & "', " & _
           "COMPOPERATOR = " & "'" & sField08 & "', " & _
           "BACKCOLOR = " & "'" & sField09 & "', " & _
           "FORECOLOR = " & "'" & sField10 & "' " & _
           "WHERE STATUSCD = " & "'" & sField01 & "' "
   
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Edit_FLAG(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Delete_FLAG(ByVal sField01 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    
    sSql = "DELETE FROM STATUS " & _
           "WHERE EXISTS " & _
           "(SELECT * FROM STATUS " & _
           "WHERE STATUSCD = " & "'" & sField01 & "') " & _
           "AND STATUSCD = " & "'" & sField01 & "' "
        
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Delete_FLAG(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Get_FLAG(Optional ByVal sFlag As String)
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i%, j%
    
    sFlag = Trim$(sFlag)
    
    If sFlag <> "" Then
    Else
        sSql = "SELECT S.STATUSCD, S.PARTCD, S.PARTGBN, " & _
               "S.SPECIMENCD, S.TESTITEMSEQ, S.SUBMCD, T.PRINTNM, " & _
               "S.CRITICALVAL, S.COMPOPERATOR, S.BACKCOLOR, S.FORECOLOR " & _
               "FROM STATUS S, TESTITEM T " & _
               "WHERE S.PARTCD = T.PARTCD " & _
               "AND S.PARTGBN = T.PARTGBN " & _
               "AND S.SPECIMENCD = T.SPECIMENCD " & _
               "AND S.TESTITEMSEQ = T.TESTITEMSEQ " & _
               "AND S.SUBMCD = T.SUBMCD "
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        AdoRecord.Close
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            If sFlag <> "" Then
            Else
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
                mvarTotField03 = mvarTotField03 & AdoRecord(2) & "|"
                mvarTotField04 = mvarTotField04 & AdoRecord(3) & "|"
                mvarTotField05 = mvarTotField05 & AdoRecord(4) & "|"
                mvarTotField06 = mvarTotField06 & AdoRecord(5) & "|"
                mvarTotField07 = mvarTotField07 & AdoRecord(6) & "|"
                mvarTotField08 = mvarTotField08 & AdoRecord(7) & "|"
                mvarTotField09 = mvarTotField09 & AdoRecord(8) & "|"
                mvarTotField10 = mvarTotField10 & AdoRecord(9) & "|"
                mvarTotField11 = mvarTotField11 & AdoRecord(10) & "|"
                
            End If
                 
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_FLAG(AB01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub

Public Sub Add_COMCD(ByVal sField01 As String, ByVal sField02 As String, ByVal sField03 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField03)

    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "INSERT INTO COMMENTCD VALUES " & _
           "( " & "'" & sField01 & "', " & "'" & sField02 & "', " & "'" & sField03 & "' )"

    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Add_COMCD(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Edit_COMCD(ByVal sField01 As String, ByVal sField02 As String, ByVal sField03 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    sField03 = Trim$(sField03)
        
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "UPDATE COMMENTCD SET COMMENTNM = " & "'" & sField03 & "' " & _
           "WHERE PARTCD = " & "'" & sField01 & "' " & _
           "AND COMMENTCD = " & "'" & sField02 & "' "
   
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Edit_COMCD(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Delete_COMCD(ByVal sField01 As String, ByVal sField02 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)

    sSql = "DELETE FROM COMMENTCD " & _
           "WHERE PARTCD = " & "'" & sField01 & "' " & _
           "AND COMMENTCD = " & "'" & sField02 & "' "
        
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Delete_COMCD(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Get_COMCD(Optional ByVal sField01 As String, Optional ByVal sField02 As String)
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i%, j%
        
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    
    If sField01 = "" Then
        sSql = "SELECT * " & _
               "FROM COMMENTCD ORDER BY PARTCD, COMMENTCD "
    Else
        If sField01 <> "" And sField02 = "" Then    'PART
            sSql = "SELECT COMMENTCD, COMMENTNM " & _
               "FROM COMMENTCD " & _
               "WHERE PARTCD = " & "'" & sField01 & "' "
               
        ElseIf sField01 <> "" And sField02 <> "" Then
            If Len(sField02) = 1 Then  'Interface Only OR 'Exam Only OR 'Sunsu Comm Only
                sSql = "SELECT COMMENTCD, COMMENTNM " & _
                    "FROM COMMENTCD " & _
                    "WHERE PARTCD = " & "'" & sField01 & "' " & _
                    "AND COMMENTCD LIKE " & "'" & sField02 & "%' "
            
            Else    'Specific Commentcd
                sSql = "SELECT COMMENTCD, COMMENTNM " & _
                    "FROM COMMENTCD " & _
                    "WHERE PARTCD = " & "'" & sField01 & "' " & _
                    "AND COMMENTCD =" & "'" & sField02 & "' "
            End If
        End If
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        AdoRecord.Close
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            If sField01 = "" Then
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
                mvarTotField03 = mvarTotField03 & AdoRecord(2) & "|"
                
            Else
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
                
            End If
                 
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_COMCD(AB01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub

Public Sub Add_MACH(ByVal sField01 As String, ByVal sField02 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
        
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "INSERT INTO MACHINE VALUES " & _
           "( " & "'" & sField01 & "', " & "'" & sField02 & "', '0', '0' )"

    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Add_MACH(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Edit_MACH(ByVal sField01 As String, ByVal sField02 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
        
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = "UPDATE MACHINE SET MACHINENM = " & "'" & sField02 & "' " & _
           "WHERE MACHINECD = " & "'" & sField01 & "' "
   
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Edit_MACH(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Delete_MACH(ByVal sField01 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    
    sField01 = Trim$(sField01)
    
    sSql = "DELETE FROM MACHINE WHERE MACHINECD = " & _
           "'" & sField01 & "' "
        
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.Execute sSql
    
    AdoConn.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Delete_MACH(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Get_MACH(Optional ByVal sField01 As String)
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sField01 = Trim$(sField01)
    
    If sField01 <> "" Then
        sSql = "SELECT MACHINENM FROM MACHINE " & _
               "WHERE MACHINECD = " & "'" & sField01 & "' " & _
               "ORDER BY MACHINECD"
    Else
        sSql = "SELECT MACHINECD, MACHINENM FROM MACHINE ORDER BY MACHINECD"
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            If sField01 <> "" Then
                mvarTotField01 = mvarTotField01 & AdoRecord(0)
            Else
                mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
                mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
            End If
            
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_MACH(A01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub

Public Sub Save_MorphCd(ByVal sField01 As String, ByVal sField02 As String, ByVal iCnt As Integer)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    Dim i As Integer
        
    sSql = "DELETE FROM MORCD "
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.Execute sSql
    
    mvarErrNum = 0
    
    For i = 1 To iCnt
        sSql = "INSERT INTO STATUS VALUES " & _
               "( " & _
               "'" & GetByOne(sField01, sField01) & "', " & "'" & GetByOne(sField02, sField02) & "' " & _
               ")"
               
        AdoConn.Execute sSql
    Next
    
    mvarErrNum = 0
    
    AdoConn.Close
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Save_MorphCd(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub
Public Sub Get_MorphCd()
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
        
    sSql = "SELECT * FROM MORCD ORDER BY MORCD"
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
            mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
        
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_MorphCd(A01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub

Public Sub Save_SecondCd(ByVal sField01 As String, ByVal sField02 As String, ByVal iCnt As Integer)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Connection
    Dim sSql As String
    Dim i As Integer
        
    sSql = "DELETE FROM MORCD "
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.Execute sSql
    
    mvarErrNum = 0
    
    For i = 1 To iCnt
        sSql = "INSERT INTO STATUS VALUES " & _
               "( " & _
               "'" & GetByOne(sField01, sField01) & "', " & "'" & GetByOne(sField02, sField02) & "' " & _
               ")"
               
        AdoConn.Execute sSql
    Next
    
    mvarErrNum = 0
    
    AdoConn.Close
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Save_SecondCd(A01.DLL) - " & Err.Description
    AdoConn.Close
End Sub
Public Sub Get_SecondCd()
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
        
    sSql = "SELECT * FROM SECONDCD ORDER BY SECONDCD"
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            mvarTotField01 = mvarTotField01 & AdoRecord(0) & "|"
            mvarTotField02 = mvarTotField02 & AdoRecord(1) & "|"
            mvarTotField02 = mvarTotField02 & AdoRecord(2) & "|"
            
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub
    
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_SecondCd(A01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub

Public Sub Get_Login_Info(ByVal iMode As Integer, ByVal sField01 As String, ByVal sField02 As String)
    On Error GoTo ErrHandler
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sField01 = Trim$(sField01)
    sField02 = Trim$(sField02)
    
    If iMode = 0 Then   '초기화 LogOn 확인시
        sSql = "SELECT USER.PASSWORD, USER.USERNM, USER.SLIPCD, USER.SPECIMENCD, PART.SLIPNM, SPECIMEN.SPECIMENBRIEFNM, USER.SCHOPT " & _
               "FROM PART, SPECIMEN, USER " & _
               "WHERE SPECIMEN.SPECIMENCD = USER.SPECIMENCD " & _
               "AND PART.PARTCD & PART.PARTGBN = USER.SLIPCD " & _
               "AND USER.USERCD = " & "'" & sField01 & "'"
               
    ElseIf iMode = 1 Then   'All Search
       
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        mCurItemCnt = 0
        AdoRecord.Close
        Exit Sub
    Else
        i = AdoRecord.RecordCount
        mCurItemCnt = i
        
        AdoRecord.MoveFirst
        
        For j = 1 To i
            If iMode = 0 Then
                mvarTotField01 = mvarTotField01 & AdoRecord("PASSWORD")
                mvarTotField02 = mvarTotField02 & AdoRecord("USERNM")
                mvarTotField03 = mvarTotField03 & AdoRecord("SLIPCD")
                mvarTotField04 = mvarTotField04 & AdoRecord("SLIPNM")
                mvarTotField05 = mvarTotField05 & AdoRecord("SPECIMENCD")
                mvarTotField06 = mvarTotField06 & AdoRecord("SPECIMENBRIEFNM")
                mvarTotField07 = mvarTotField07 & AdoRecord("SCHOPT")
            End If
            
            If iMode = 1 Then
            
            End If
                 
            AdoRecord.MoveNext
        Next
    End If
    
    AdoRecord.Close
    
    mvarErrNum = 0
    
    Exit Sub
ErrHandler:
    mvarErrNum = Err.Number
    MsgBox "Get_Login_Info(AB01.DLL) - " & Err.Description
    AdoRecord.Close
End Sub


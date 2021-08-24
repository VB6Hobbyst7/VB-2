VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDisease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarPtId    As String
Private mvarBedInDt As String

Public Property Let ptid(ByVal vData As String)
    mvarPtId = vData
End Property

Private Function InOutGubun() As Long
    '1:외래 2:입원
    Dim sSQL As String
    Dim DrRS As Recordset
    
'    sSql = "select ptid,orddt,ordno,bussdiv,bedindt " & _
           "from " & T_LAB101 & " " & _
           "where " & DBW("ptid=", mvarPtId) & " " & _
           "and   " & DBW("orddt=", mvarOrdDt) & " " & _
           "and   " & DBW("ordno=", mvarOrdNo)
    
    sSQL = "select distinct ptid,orddt,ordno,bussdiv,bedindt " & _
           "from " & T_LAB101 & " " & _
           "where " & DBW("ptid=", mvarPtId) & " " & _
           "and  orddt= (select max(orddt) from " & T_LAB101 & " where " & DBW("ptid=", mvarPtId) & ")"
           
    
    Set DrRS = New Recordset
    DrRS.Open sSQL, DBConn
    
'    If DrRS.DBerror Then
'        InOutGubun = -1
'    Else
        If DrRS.EOF Then
            InOutGubun = -1
        Else
            InOutGubun = DrRS.Fields("bussdiv").Value & ""
            mvarBedInDt = DrRS.Fields("bedindt").Value & ""
        End If
'    End If
    Set DrRS = Nothing
End Function

Private Function GetInDiseaseSql() As String
    GetInDiseaseSql = " select a.seq,a.icd diseacd,b." & F_IENM & " diseanm " & _
                      " from " & T_HIS006 & " b," & T_LAB106 & " a " & _
                      " where " & DBW("a.ptid=", mvarPtId) & " " & _
                      " and   " & DBW("a.bedindt=", mvarBedInDt) & " " & _
                      " and   " & DBJ("a.icd*=b." & F_ICD) & _
                      " order by a.seq"
End Function

Private Function GetOutDiseaseSql() As String
    GetOutDiseaseSql = " select a.seq,a.icd diseacd,b." & F_IENM & " diseanm " & _
                       " from " & T_HIS006 & " b," & T_LAB105 & " a " & _
                       " where " & DBW("a.ptid=", mvarPtId) & " " & _
                       " and   " & DBJ("a.icd*=b." & F_ICD) & " " & _
                       " and  a.orddt=(select max(z.orddt) from " & T_LAB105 & " z where " & DBW("z.ptid=", mvarPtId) & ") " & _
                       " and   " & DBJ("a.icd*=b." & F_ICD) & _
                       " order by a.seq"
End Function

Public Function Disease() As String
    Dim DrRS          As Recordset
    Dim sSQL          As String
    Dim lngInOutGubun As Long
    
    On Error GoTo ErrMsg
    
    '외래와 병동환자의 상병테이블이 다르다
    lngInOutGubun = InOutGubun
    If lngInOutGubun < 0 Then: Exit Function
        
    '상병을 불러온다(병동/응급은 입원상병테이블에서)
    If lngInOutGubun = 2 Or lngInOutGubun = 3 Then
        sSQL = GetInDiseaseSql
    Else
        sSQL = GetOutDiseaseSql
    End If

    Set DrRS = New Recordset
    Call DrRS.Open(sSQL, DBConn)
'    If DrRS.DBerror = True Then
'        Disease = ""
'        Exit Function
'    End If
'    If DrRS.RecordCount > 0 Then
        Do Until DrRS.EOF
            If DrRS.Fields("seq").Value & "" = 1 Then
                Disease = DrRS.Fields("diseanm").Value & ""
                Exit Do
            End If
            DrRS.MoveNext
        Loop
'    Else
'        Disease = ""
'    End If
    
ErrMsg:

    Set DrRS = Nothing
    
End Function

Public Function DiseaseCdNM() As String
    Dim DrRS          As Recordset
    Dim sSQL          As String
    Dim lngInOutGubun As Long
    
        
    '외래와 병동환자의 상병테이블이 다르다
    lngInOutGubun = InOutGubun
    If lngInOutGubun < 0 Then: Exit Function
        
    '상병을 불러온다(병동/응급은 입원상병테이블에서)
    If lngInOutGubun = 2 Or lngInOutGubun = 3 Then
        sSQL = GetInDiseaseSql
    Else
        sSQL = GetOutDiseaseSql
    End If

    Set DrRS = New Recordset
    Call DrRS.Open(sSQL, DBConn)
'    If DrRS.DBerror = True Then
'        DiseaseCdNM = ""
'        Exit Function
'    End If
'    If DrRS.RecordCount > 0 Then
        Do Until DrRS.EOF
            If DrRS.Fields("seq").Value & "" = 1 Then
                DiseaseCdNM = DrRS.Fields("diseacd").Value & "" & vbTab & DrRS.Fields("diseanm").Value & ""
                Exit Do
            End If
            DrRS.MoveNext
        Loop
'    Else
'        DiseaseCdNM = ""
'    End If
    

    Set DrRS = Nothing
End Function

'========================================================
' 08.10.24. 양성현 전주 예수병원 요청사항.
'========================================================

Public Function DiseaseSang() As String
    Dim DrRS          As Recordset
    Dim sSQL          As String
    Dim lngInOutGubun As Long
    
    On Error GoTo ErrMsg
    
    '외래와 병동환자의 상병테이블이 다르다
    lngInOutGubun = InOutGubun
    If lngInOutGubun < 0 Then: Exit Function
        
    sSQL = " select a.seq,a.icd diseacd,b." & F_IKNM & "  diseanm "

    '상병을 불러온다(병동/응급은 입원상병테이블에서)

    If lngInOutGubun = 2 Or lngInOutGubun = 3 Then
        sSQL = sSQL & _
              " from " & T_HIS006 & " b," & T_LAB106 & " a " & _
                      " where " & DBW("a.ptid=", mvarPtId) & " " & _
                      " and   " & DBW("a.bedindt=", mvarBedInDt) & " " & _
                      " and   " & DBJ("a.icd*=b." & F_ICD)
    Else
        sSQL = sSQL & _
                       " from " & T_HIS006 & " b," & T_LAB105 & " a " & _
                       " where " & DBW("a.ptid=", mvarPtId) & " " & _
                       " and   " & DBJ("a.icd*=b." & F_ICD) & " " & _
                       " and  a.orddt=(select max(z.orddt) from " & T_LAB105 & " z where " & DBW("z.ptid=", mvarPtId) & ") " & _
                       " and   " & DBJ("a.icd*=b." & F_ICD)
    End If

' 08.10.24. 양성현 전주 예수병원 요청사항.

    sSQL = sSQL & " And infegrd > 0 " & " order by a.seq"

    Set DrRS = New Recordset
    Call DrRS.Open(sSQL, DBConn)
'    If DrRS.DBerror = True Then
'        Disease = ""
'        Exit Function
'    End If
'    If DrRS.RecordCount > 0 Then
        Do Until DrRS.EOF
            If DrRS.Fields("seq").Value & "" = 1 Then
                DiseaseSang = DrRS.Fields("diseanm").Value & ""
                Exit Do
            End If
            DrRS.MoveNext
        Loop
'    Else
'        Disease = ""
'    End If
    
ErrMsg:

    Set DrRS = Nothing
    
End Function

'========================================================
' 2010.07.06. 온승호 전주 예수병원 요청사항.
' 감염주의 테이블 변경 MDCAUTNT테이블
' INFORM IS NOT NULL
'========================================================

Public Function DiseaseSang_New() As String
    Dim DrRS          As Recordset
    Dim sSQL          As String
    Dim lngInOutGubun As Long
    
    On Error GoTo ErrMsg
    
    sSQL = "SELECT * FROM MDCAUTNT"
    sSQL = sSQL & " WHERE PATNO = '" & mvarPtId & "'"
    sSQL = sSQL & "   AND INFORM IS NOT NULL"
    
    Set DrRS = New Recordset
    Call DrRS.Open(sSQL, DBConn)
    
    Do Until DrRS.EOF
        DiseaseSang_New = DrRS.Fields("PATNO").Value & ""
        DrRS.MoveNext
    Loop
    
ErrMsg:

    Set DrRS = Nothing
    
End Function

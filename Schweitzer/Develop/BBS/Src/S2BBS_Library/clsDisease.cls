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
Private mvarOrdDt   As String
Private mvarOrdNo   As String
Private mvarBedIndt As String
Private mvarDisease As Collection
Private index       As Long

Public Property Let PtId(ByVal vData As String)
    mvarPtId = vData
End Property

Public Property Get PtId() As String
    PtId = mvarPtId
End Property

Public Property Let orddt(ByVal vData As String)
    mvarOrdDt = vData
End Property

Public Property Get orddt() As String
    orddt = mvarOrdDt
End Property

Public Property Let OrdNo(ByVal vData As String)
    mvarOrdNo = vData
End Property

Public Property Get OrdNo() As String
    OrdNo = mvarOrdNo
End Property

Public Property Get DiseaseCd() As String
    If mvarDisease.Count = 0 Then
        DiseaseCd = ""
    Else
        DiseaseCd = medGetP(mvarDisease.Item(index), 1, vbTab)
    End If
End Property

Public Property Get DiseaseNm() As String
    If mvarDisease.Count = 0 Then
        DiseaseNm = ""
    Else
        DiseaseNm = medGetP(mvarDisease.Item(index), 2, vbTab)
    End If
End Property

Public Function GetDiseaseString(Optional ByVal pColDiv As String = " ", Optional ByVal pLineDiv As String = vbNewLine) As String
    Dim i As Long
    
    For i = 1 To mvarDisease.Count
        If GetDiseaseString = "" Then
            GetDiseaseString = medGetP(mvarDisease.Item(i), 1, vbTab) & pColDiv & _
                               medGetP(mvarDisease.Item(i), 2, vbTab)
        Else
            GetDiseaseString = GetDiseaseString & pLineDiv & _
                               medGetP(mvarDisease.Item(i), 1, vbTab) & pColDiv & _
                               medGetP(mvarDisease.Item(i), 2, vbTab)
        End If
    Next i
End Function

Public Property Get RecordCount() As Long
    If mvarDisease Is Nothing Then
        RecordCount = mvarDisease.Count
    Else
        RecordCount = 0
    End If
End Property

Public Sub MoveFirst()
    index = 1
End Sub

Public Sub MoveNext()
    index = index + 1
End Sub

Public Function EOF() As Boolean
    EOF = (index > mvarDisease.Count)
End Function

Public Sub Clear()
    mvarPtId = ""
    mvarOrdDt = ""
    mvarOrdNo = ""
    mvarBedIndt = ""
    index = 0
    Set mvarDisease = Nothing
    Set mvarDisease = New Collection
End Sub

Private Function InOutGubun() As Long
    '1:외래 2:입원
    Dim sSql As String
    Dim DrRS As Recordset
    
    sSql = "SELECT ptid,orddt,ordno,bussdiv,bedindt " & _
           "FROM " & T_LAB101 & " " & _
           "WHERE " & DBW("ptid=", mvarPtId) & " " & _
           "AND   " & DBW("orddt=", mvarOrdDt) & " " & _
           "AND   " & DBW("ordno=", mvarOrdNo)
    Set DrRS = New Recordset
    DrRS.Open sSql, DBConn
    
'    If DrRS.DBerror Then
'        InOutGubun = -1
'    Else
        If DrRS.RecordCount < 1 Then
            InOutGubun = -1
        Else
            InOutGubun = DrRS.Fields("bussdiv").Value & ""
            mvarBedIndt = DrRS.Fields("bedindt").Value & ""
        End If
'    End If
    Set DrRS = Nothing
End Function

Private Function GetInDiseaseSql() As String
    GetInDiseaseSql = "SELECT a.seq,a.icd diseacd,b." & F_IENM & " diseanm " & _
                      "FROM " & T_LAB106 & " a," & T_HIS006 & " b " & _
                      "WHERE " & DBW("a.ptid=", mvarPtId) & " " & _
                      "AND   " & DBW("a.bedindt=", mvarBedIndt) & " " & _
                      "AND   " & DBJ("a.icd*=b." & F_ICD) & " " & _
                      "ORDER BY a.seq"
End Function

Private Function GetOutDiseaseSql() As String
    GetOutDiseaseSql = "SELECT a.seq,a.icd diseacd,b." & F_IENM & " diseanm " & _
                       "FROM " & T_LAB105 & " a," & T_HIS006 & " b " & _
                       "WHERE " & DBW("a.ptid=", mvarPtId) & " " & _
                       "AND   " & DBW("a.orddt=", mvarOrdDt) & " " & _
                       "AND   " & DBW("a.ordno=", mvarOrdNo) & " " & _
                       "AND   " & DBJ("a.icd*=b." & F_ICD) & " " & _
                       "ORDER BY a.seq"
End Function

Public Function GetDisease() As Boolean
    Dim sSql As String
    Dim DrRS As Recordset
    Dim i As Long
    Dim lngInOutGubun As Long
    
    If mvarDisease Is Nothing Then Set mvarDisease = New Collection
    
    If mvarPtId = "" Then
        MsgBox "환자ID가 없습니다", vbCritical, "오류"
        GetDisease = False
        Exit Function
    End If
    If mvarOrdDt = "" Then
        MsgBox "처방일 정보가 없습니다", vbCritical, "오류"
        GetDisease = False
        Exit Function
    End If
    If mvarOrdNo = "" Then
        MsgBox "처방번호 정보가 없습니다", vbCritical, "오류"
        GetDisease = False
        Exit Function
    End If
    
    '외래와 병동환자의 상병테이블이 다르다
    lngInOutGubun = InOutGubun
    If lngInOutGubun < 0 Then GetDisease = False: Exit Function
        
    '상병을 불러온다----------------------------------------------
    If lngInOutGubun = 2 Then
        sSql = GetInDiseaseSql
    Else
        sSql = GetOutDiseaseSql
    End If

    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
'    If DrRS.DBerror = True Then
'        dbconn.DisplayErrors
'        GetDisease = False
'        Exit Function
'    End If
    
    With DrRS
        index = 1
        For i = 1 To .RecordCount
            If Trim(.Fields("diseacd")) <> "" Then
                mvarDisease.Add .Fields("diseacd").Value & "" & vbTab & Trim(.Fields("diseanm").Value & "")
            End If
            .MoveNext
        Next i
    End With
    Set DrRS = Nothing
    GetDisease = True
End Function

Private Sub Class_Initialize()
    Set mvarDisease = New Collection
End Sub

Private Sub Class_Terminate()
    Set mvarDisease = Nothing
End Sub

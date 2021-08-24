VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDonorAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetDonorOk(ByVal Donorid As String, ByVal donoraccdt As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT okdiv1,okdiv2,okdiv3,rmk1,rmk2,rmk3 " & _
           " FROM " & T_BBS603 & " " & _
           " WHERE " & DBW("donorid", Donorid, 2) & _
           " AND " & DBW("donoraccdt", donoraccdt, 2)

    Set GetDonorOk = New Recordset
    Call GetDonorOk.Open(sSql, DBConn)
'    If GetDonorOk.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetDonorOk = Nothing
'    End If
End Function

Public Function GetDonorAsk(ByVal Donorid As String, ByVal donoraccdt As String) As Recordset
    Dim sSql As String
    
    sSql = "SELECT askcd,yesno,okdiv " & _
           "FROM " & T_BBS604 & " " & _
           "WHERE " & DBW("donorid", Donorid, 2) & _
           "AND " & DBW("donoraccdt", donoraccdt, 2)
           
    Set GetDonorAsk = New Recordset
    Call GetDonorAsk.Open(sSql, DBConn)
'    If GetDonorAsk.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetDonorAsk = Nothing
'    End If
End Function

Public Function Save(ByVal Donorid As String, ByVal donoraccdt As String, _
                     ByVal okdiv As String, ByVal rmk As String, _
                     ByVal ask As String, _
                     ByVal IsHold As Boolean) As Boolean

    '-----------------------------------------------------
    'ask : 문진결과
    '      askcd COL_DIV yesno COL_DIV okdiv LINE_DIV ....
    '-----------------------------------------------------
    Dim askcd As String
    Dim askyesno As String
    Dim askokdiv As String
    
    Dim dSql As String  '삭제문장용
    Dim sSql() As String
    Dim i As Long
    Dim tmpRec As String
    Dim OkDt2 As String
    Dim stscd As String
    
    If IsHold Then
        stscd = DonorStatus.stsAskSave
    Else
        stscd = DonorStatus.stsAskVerify
    End If
    
    OkDt2 = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    ReDim Preserve sSql(0)
    sSql(0) = " update " & T_BBS603 & " " & _
              " set " & DBW("stscd", stscd, 3) & _
                        DBW("okdiv2", okdiv, 3) & _
                        DBW("okdt2", OkDt2, 3) & _
                        DBW("rmk2", rmk, 2) & _
              " WHERE " & _
                                DBW("donorid", Donorid, 2) & _
                      " AND " & DBW("donoraccdt", donoraccdt, 2)
    i = 0
    Do
        i = i + 1
        If i = 1 Then
            dSql = " delete FROM " & T_BBS604 & " " & _
                   " WHERE " & _
                                     DBW("donorid", Donorid, 2) & _
                           " AND " & DBW("donoraccdt", donoraccdt, 2)
        End If
        
        tmpRec = medGetP(ask, i, LINE_DIV)
        If tmpRec = "" Then Exit Do
        
        askcd = medGetP(tmpRec, 1, COL_DIV)
        askyesno = medGetP(tmpRec, 2, COL_DIV)
        askokdiv = medGetP(tmpRec, 3, COL_DIV)
        
        ReDim Preserve sSql(i)
        
        sSql(i) = "insert into " & T_BBS604 & "(donorid,donoraccdt,askcd,yesno,okdiv) " & _
                  "values(" & _
                         DBV("donorid", Donorid, 1) & DBV("donoraccdt", donoraccdt, 1) & DBV("askcd", askcd, 1) & _
                         DBV("yesno", askyesno, 1) & DBV("okdiv", askokdiv) & ")"

    Loop
    
On Error GoTo Save_error

    DBConn.BeginTrans
    DBConn.Execute dSql
    For i = LBound(sSql) To UBound(sSql)
        DBConn.Execute sSql(i)
    Next i
    DBConn.CommitTrans
    Save = True
    Exit Function
    
Save_error:
'    DBConn.DisplayErrors
    DBConn.RollbackTrans
    Save = False
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'----------------------------------------------------------------
'검체번호와 접수번호를(최종번호)를 가지고오며, Updat/Insert해준다.
'----------------------------------------------------------------

Public Function Get_AccdtFormat() As String
'-----------------------------------------
'접수일자의 형식별로 접수일자를 가지고온다
'-----------------------------------------
    Dim sSql As String
    Dim Rs As New Recordset
    Dim AccdtFormat As Long
    
    sSql = " SELECT a.field1" & _
           " FROM " & _
                    T_COM003 & " a" & _
           " WHERE" & _
                    " " & DBW("a.cdindex", BC2_ACCNO_CRITERION, 2) & _
                    " AND a.cdval1=(SELECT max(b.cdval1) FROM " & T_COM003 & " b" & _
                                  " WHERE" & DBW("b.cdindex", BC2_ACCNO_CRITERION, 2) & _
                                  " AND " & DBW(" b.cdval1<", Format(GetSystemDate, PRESENTDATE_FORMAT), 2) & _
                                  ") "
    
    Rs.Open sSql, DBConn
        
    If Not Rs.EOF = True Then
        Select Case Rs.Fields("field1").Value & ""
            Case "0": AccdtFormat = 8
            Case "1": AccdtFormat = 6
            Case "2": AccdtFormat = 4
        End Select
        Get_AccdtFormat = Mid(Format(GetSystemDate, PRESENTDATE_FORMAT), 1, AccdtFormat)
    Else
        Get_AccdtFormat = Format(GetSystemDate, PRESENTDATE_FORMAT)
    End If
        
    Set Rs = Nothing
End Function

Public Function Get_AccDT_Seq(ByVal accdt As String) As String
'---------------------
'접수번호를 가지고온다
'---------------------
    Dim sSql As String
    Dim Rs As New Recordset
    
    sSql = " SELECT divcd2,seq " & _
           " FROM  " & T_COM099 & " " & _
           " WHERE " & _
                                DBW("noindex=", BN_ACC_NO) & _
                     " AND  " & DBW("divcd1=", C_WORKAREA) & _
                     " AND  " & DBW("divcd2=", accdt) & _
                     " AND  " & DBW("divcd3=", "0")
           
    Rs.Open sSql, DBConn
    If Not Rs.EOF = True Then
        Get_AccDT_Seq = Val(Rs.Fields("seq").Value & "") + 1
    Else
        Get_AccDT_Seq = 1
    End If
    
    Set Rs = Nothing
End Function
Public Function Get_SpcYY_No(ByVal Year As String) As String
'---------------------
'검체번호를 가지고온다
'---------------------
    Dim sSql As String
    Dim Rs   As New Recordset
    
    sSql = "SELECT divcd2,seq " & _
           "FROM " & T_COM099 & " " & _
           "WHERE" & _
                   "    " & DBW("noindex", BN_SPC_NO, 2) & _
                   "AND " & DBW("divcd1", C_WORKAREA, 2) & _
                   "AND " & DBW("divcd2", Year, 2) & _
                   "AND " & DBW("divcd3", "0", 2)
                   
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF = True Then
        Get_SpcYY_No = Rs.Fields("divcd2").Value & "" & COL_DIV & Val(Rs.Fields("seq").Value & "") + 1
    Else
        Get_SpcYY_No = Year & COL_DIV & "1"
    End If
 
    Set Rs = Nothing
End Function

Public Function Set_NumbersCom099(ByVal CDINDEX As String, _
                                  ByVal Divcd2 As String, _
                                  ByVal seq As Long) As String
'--------------------------
'번호부여정보 저장 Sql 문장
'--------------------------
    Dim Rs   As Recordset
    Dim sSql As String
    
    sSql = " SELECT * FROM " & T_COM099 & _
          " WHERE " & _
                    DBW("noindex=", CDINDEX) & _
                    " AND " & DBW("divcd1=", C_WORKAREA) & _
                    " AND " & DBW("divcd2=", Divcd2) & _
                    " AND " & DBW("divcd3=", "0")
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.EOF Then
        Set_NumbersCom099 = "insert into " & T_COM099 & "( " & _
                           "noindex,divcd1,divcd2,divcd3,seq) " & _
                           "values(" & _
                                     DBV("noindex", CDINDEX, 1) & _
                                     DBV("divcd1", C_WORKAREA, 1) & _
                                     DBV("divcd2", Divcd2, 1) & _
                                     DBV("divcd3", "0", 1) & _
                                     DBV("seq", seq) & _
                                 ") "
    Else
        Set_NumbersCom099 = " update " & T_COM099 & _
                           " set    " & DBW("seq=", seq) & _
                           " WHERE  " & _
                                                DBW("noindex=", CDINDEX) & _
                                      " AND " & DBW("divcd1=", C_WORKAREA) & _
                                      " AND " & DBW("divcd2=", Divcd2) & _
                                      " AND " & DBW("divcd3=", "0")
    End If
    
    Set Rs = Nothing
End Function

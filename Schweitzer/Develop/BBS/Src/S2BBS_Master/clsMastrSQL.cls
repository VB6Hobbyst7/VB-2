VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMastrSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetCom002SQL(ByVal qCdindex As String, Optional ByVal qCdval1 As String = "", Optional ByVal qCdval2 As String = "") As String
    Dim SSQL    As String
    
    SSQL = " SELECT * FROM " & T_COM002 & " WHERE " & DBW("cdindex=", qCdindex)
    If qCdval1 <> "" Then
        SSQL = SSQL & " AND " & DBW("cdval1=", qCdval1)
    End If
    If qCdval2 <> "" Then
        SSQL = SSQL & " AND " & DBW("cdval2=", qCdval2)
    End If
    SSQL = SSQL & " ORDER BY cdval1"
    
    GetCom002SQL = SSQL
End Function

Public Function GetCompoSQL(Optional ByVal qCompocd As String = "") As String
    Dim SSQL    As String
    
    SSQL = " SELECT * FROM " & T_BBS006
    If qCompocd <> "" Then
        SSQL = SSQL & " WHERE " & DBW("compocd=", qCompocd) & " and (expdt ='' or expdt is null)"
    Else
        SSQL = SSQL & " WHERE  (expdt ='' or expdt is null)"
    End If
    SSQL = SSQL & " ORDER BY compocd"
    GetCompoSQL = SSQL
End Function

Public Function GetCompoMoney() As String
    Dim SSQL    As String
    
    SSQL = " SELECT a.cdval1,a.cdval2,a.field1,b.abbrnm " & _
           " FROM " & T_BBS006 & " b," & T_COM002 & " a" & _
           " WHERE " & DBW("a.cdindex=", BC2_CompoValue) & _
           " AND a.cdval1=b.compocd"
         
    GetCompoMoney = SSQL & " order by cdval1"
End Function

Public Function DeleteCOM002(ByVal CDINDEX As String, ByVal cdval1 As String, ByVal cdval2 As String) As String
    Dim SSQL As String
    
    SSQL = " delete  FROM " & T_COM002 & _
           " WHERE " & DBW("cdindex=", CDINDEX) & _
           " AND " & DBW("cdval1=", cdval1) & _
           " AND " & DBW("cdval2=", cdval2)
    
    DeleteCOM002 = SSQL
End Function


Public Function InsertCOM002(ByVal CDINDEX As String, ByVal cdval1 As String, ByVal cdval2 As String, _
                             Optional ByVal field1 As String = "", Optional ByVal field2 As String = "", _
                             Optional ByVal Field3 As String = "", Optional ByVal Field4 As String = "", _
                             Optional ByVal Text1 As String = "", Optional ByVal Text2 As String = "") As String
    Dim SSQL As String
    
    SSQL = " INSERT INTO " & T_COM002 & " (cdindex,cdval1,cdval2,field1,field2,field3,field4,text1,text2) " & _
           " VALUES(" & _
                    DBV("cdindex", CDINDEX, 1) & DBV("cdval1", cdval1, 1) & DBV("cdval2", cdval2, 1) & DBV("field1", field1, 1) & _
                    DBV("field2", field2, 1) & DBV("field3", Field3, 1) & DBV("field4", Field4, 1) & _
                    DBV("text1", Text1, 1) & DBV("text2", Text2) & ") "

    InsertCOM002 = SSQL
End Function

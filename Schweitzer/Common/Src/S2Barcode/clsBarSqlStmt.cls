VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarSqlStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function SQL_ReadConfig(ByVal strPrjNm) As String

    SQL_ReadConfig = " SELECT cdval2,field1,field2,field3,field4,field5 " & _
                     " FROM  " & objTables.T_COM002 & _
                     " WHERE " & DBW("cdindex = ", COM2_BAR_CONFIG) & _
                     " AND   " & DBW("cdval1  = ", strPrjNm)

End Function

Public Function SQL_ReadHan() As String
'zebra의 한글 폰트를 가지고 온다.
    SQL_ReadHan = " SELECT * FROM " & objTables.T_COM004 & _
                  " WHERE " & DBW("cdindex", "zebra", 2)

End Function

Public Function Get_Port() As String

    Get_Port = " SELECT distinct field1 FROM " & objTables.T_COM002 & _
               " WHERE " & DBW("cdindex=", COM2_BAR_CONFIG) & _
               " AND cdval1 in('LIS','BBS','APS') " & _
               " AND " & DBW("cdval2=", BAR_PORT)

End Function

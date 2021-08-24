VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMasters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'
'Public Function GetBldSrcNm(ByVal code As String) As String
'    Dim RS      As Recordset
'    Dim sSql    As String
'
'    sSql = "SELECT * FROM " & T_COM003 & " " & _
'           "WHERE " & DBW("cdindex=", BC2_BLD_SRC) & " " & _
'           "AND   " & DBW("cdval1=", code)
'
'    Set RS = New Recordset
'    RS.open ssql,dbconn
'    If RS.DBerror = True Then
'        GetBldSrcNm = ""
'    Else
'        If RS.RecordCount > 0 Then
'            GetBldSrcNm = RS.Fields("field1").Value & ""
'        Else
'            GetBldSrcNm = ""
'        End If
'    End If
'    RS.RsClose
'    Set RS = Nothing
'End Function
'
'Public Function GetBldSrcList(Optional fg As Integer = 0) As String
'    GetBldSrcList = " SELECT cdval1 ,field1 " & _
'                    " FROM " & T_COM003 & " " & _
'                    " WHERE " & DBW("cdindex=", BC2_BLD_SRC)
'End Function
'
'Public Function GetActiveBldSrc(code As String, name As String, Optional ByVal dt As String = "") As Boolean
'    Dim sSql As String
'    Dim RS As Recordset
'
'    If dt = "" Then dt = Format(GetSystemDate, PRESENTDATE_FORMAT)
'
'    sSql = "SELECT * FROM " & T_COM003 & " " & _
'           " WHERE " & DBW("cdindex=", BC2_ACTIVE_BLD_SRC) & " " & _
'           " AND   cdval1 = (SELECT max(cdval1) FROM " & T_COM003 & _
'                           " WHERE " & DBW("cdindex=", BC2_ACTIVE_BLD_SRC) & _
'                           " AND " & DBW("cdval1<", dt) & ")"
'
'    Set RS = New Recordset
'    RS.open ssql,dbconn
'
'    If RS.DBerror Then
'        'dbconn.DisplayErrors
'        code = ""
'        name = ""
'        Exit Function
'    End If
'
'    ' Ç÷¾×¿øÄÚµå ------------------------------------------------
'    code = RS.Fields("field1").Value & ""
'    RS.RsClose
'
'    sSql = "SELECT * FROM " & T_COM003 & " " & _
'           "WHERE " & DBW("cdindex=", BC2_BLD_SRC) & " " & _
'           "AND   " & DBW("cdval1=", code)
'
'    RS.open ssql,dbconn
'    If RS.DBerror Then
'        'dbconn.DisplayErrors
'        name = ""
'        Exit Function
'    End If
'
'    name = RS.Fields("field1").Value & ""
'    RS.RsClose
'
'    Set RS = Nothing
'End Function
'
'Public Function GetDestBuilding(ByVal code As String) As Recordset
'    Dim sSql As String
'
'    If code = "" Then GetDestBuilding = "": Exit Function
'
'    sSql = "SELECT * " & _
'           "FROM " & T_BBS004 & " " & _
'           "WHERE " & DBW("sbuilding=", code)
'
'    Set GetDestBuilding = New Recordset
'    Call GetDestBuilding.open(ssql,dbconn)
'
'    If GetDestBuilding.DBerror Then
'        'dbconn.DisplayErrors
'        Set GetDestBuilding = Nothing
'    End If
'End Function
'
'
'
'

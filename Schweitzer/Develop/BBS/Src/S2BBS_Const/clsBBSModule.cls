VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSModule"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function DBS(ByVal strValue As String, _
   Optional ByVal optNum As Long) As String
'String Conversion For Database INSERT,UPDATE

   strValue = Replace(strValue, "'", "''")
   DBS = "'" & CStr(strValue) & "'"
   If IsMissing(optNum) = False Then
      Select Case optNum
         Case 0
            DBS = DBS
         Case 1
            DBS = DBS & ","
         Case 2
            DBS = "=" & DBS
         Case 3
            DBS = "=" & DBS & ","
         Case Else
      End Select
   End If
   '
End Function

Public Function DBN(ByVal NumValue As String, _
   Optional ByVal optNum As Long) As String
'Number Conversion For Database INSERT,UPDATE
   DBN = CStr(NumValue)
   If DBN = "" Then DBN = "''"
   If IsMissing(optNum) = False Then
      Select Case optNum
         Case 0
            DBN = DBN
         Case 1
            DBN = DBN & ","
         Case 2
            DBN = " = " & DBN
         Case 3
            DBN = " = " & DBN & ","
         Case Else
      End Select
   End If
   '
End Function








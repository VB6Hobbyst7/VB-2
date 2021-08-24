VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PreparedStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private stmt() As String
Private values() As String

Public Function getPreparedStatement() As String
    Dim retnStr As String
    Dim i As Integer

    retnStr = ""
    
    For i = 0 To UBound(stmt)
        retnStr = retnStr & stmt(i) & values(i)
    Next i

    getPreparedStatement = retnStr
    
End Function

Public Sub setString(ByVal index As Integer, ByVal x As String)
    x = Replace(x, "'", "''")
    values(index - 1) = "'" & x & "'"
End Sub

Public Sub setNumber(ByVal index As Integer, ByVal x As Double)
    values(index - 1) = x
End Sub

Public Sub initPreparedStatement(ByVal Sql As String)
Attribute initPreparedStatement.VB_UserMemId = 0
    Dim count As Integer
    Dim i As Integer
    
    stmt = Split(Sql, "?")
    
    count = UBound(stmt) - LBound(stmt) + 1
    ReDim values(count)
End Sub




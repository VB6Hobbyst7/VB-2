VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStrState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private maryStmt() As String
Private maryValues() As String

Public Function GetPreparedStmt() As String
    Dim strRet As String
    Dim i As Integer

    strRet = ""
    
    For i = 0 To UBound(maryStmt) - 1
        strRet = strRet & maryStmt(i) & maryValues(i)
    Next i

    GetPreparedStmt = strRet
    
End Function

'flag
' 0 : string
' 1 : numeric
Public Sub SetString(ByVal Index As Integer, ByVal X As String, Optional ByVal flag As Integer = 0)
    
    X = Replace(X, "'", "''")
    
    If flag = "0" Then
        maryValues(Index - 1) = "'" & X & "'"
    ElseIf flag = "1" Then
        If X <> "" Then
            maryValues(Index - 1) = X
        Else
            maryValues(Index - 1) = "'" & X & "'"
        End If
    Else
        maryValues(Index - 1) = "'" & X & "'"
    End If
End Sub

Public Sub InitPreparedStmt(ByVal pSQL As String)
    Dim objToken As New clsStrToken
    Dim intCount As Integer
    Dim i As Integer
    
    objToken pSQL, "?"
    intCount = objToken.Count
    
    ReDim maryStmt(intCount)
    ReDim maryValues(intCount)
    
    i = 0
    While objToken.HasMoreTokens
        maryStmt(i) = objToken.NextToken
        maryValues(i) = ""
        i = i + 1
    Wend
End Sub



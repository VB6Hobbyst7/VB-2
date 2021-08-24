VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStrToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrStmt            As String
Private mstrDelim           As String
Private mintCurrentPosition As Integer

Public Function Count() As Integer
    Dim intStart As Integer
    Dim intFind As Integer
    Dim intCnt As Integer
    
    intFind = mintCurrentPosition
    Do
        intStart = intFind + 1
        intFind = InStr(intStart, mstrStmt, mstrDelim)
        
        If intFind = 0 Then Exit Do
        
        intCnt = intCnt + 1
    Loop
    
    Count = intCnt
End Function

Public Sub initStringTokenizer(ByVal Str As String, ByVal Delim As String)
    mstrStmt = Str & Delim
    mstrDelim = Delim
    mintCurrentPosition = 0
End Sub

Public Function NextToken() As String
    Dim strRet As String
    Dim intPosition As Integer
    
    intPosition = InStr(mintCurrentPosition + 1, mstrStmt, mstrDelim)
    strRet = Mid$(mstrStmt, mintCurrentPosition + 1, intPosition - (mintCurrentPosition + 1))
    mintCurrentPosition = intPosition
    
    NextToken = strRet
End Function

Public Function HasMoreTokens() As Boolean
    Dim intPosition As Integer
    
    intPosition = InStr(mintCurrentPosition + 1, mstrStmt, mstrDelim)
    
    If intPosition = 0 Then
        HasMoreTokens = False
    Else
        HasMoreTokens = True
    End If
End Function




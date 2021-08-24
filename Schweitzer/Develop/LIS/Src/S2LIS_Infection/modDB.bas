Attribute VB_Name = "modDB"
Option Explicit

Public DBConn As ADODB.Connection
Public Rep    As New ADODB.Connection
'Public Const D0COM_SERVER = " Provider=msdaora;Data Source=plis;User Id=plis;Password=plis;"
Public Const D0COM_SERVER = "Provider=MSDAORA.1;Data Source=pmc;User ID=oral1;Password=oral1;"


Public Function DBOpen(Server As String) As Boolean
    
    On Error GoTo ErrMsg
    
    Set DBConn = New ADODB.Connection
    
    With DBConn
        .CursorLocation = adUseServer
        .CommandTimeout = 0
        .Open Server
    End With
    
    DBOpen = True
    
    Exit Function
    
ErrMsg:
    Set DBConn = Nothing
    DBOpen = False
    
End Function

Public Sub DBClose()
    DBConn.Close
    Set DBConn = Nothing
End Sub

Public Sub MDBOpen(ByVal pPath As String)
    With Rep
        .CursorLocation = adUseClient
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pPath
        .Open
    End With
End Sub

Public Sub MDBClose()
    Rep.Close:   Set Rep = Nothing
End Sub



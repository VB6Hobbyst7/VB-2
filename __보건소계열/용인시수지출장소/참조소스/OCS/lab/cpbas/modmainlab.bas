Attribute VB_Name = "modBasic"
Option Explicit

Public i        As Integer
Public j        As Integer
Public sMsg     As String
Public sTitle   As String
Public nRet     As Integer

Public Function Spread_Clear(ByVal ssName As Object) As Integer
    
    ssName.Row = 1
    ssName.Row2 = ssName.DataRowCnt
    ssName.Col = 1
    ssName.Col2 = ssName.DataColCnt
    ssName.BlockMode = True
    ssName.Action = ActionClear
    ssName.BlockMode = False

End Function

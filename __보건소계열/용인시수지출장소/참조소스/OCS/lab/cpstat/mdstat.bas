Attribute VB_Name = "mdStat"
Option Explicit


'Public adoConnect       As ADODB.Connection
Public gMenuSelect      As Integer
Public GstrIdnumber        As String



Public Function SpreadSetClear(ByVal sSpread As Object) As Integer
    
    sSpread.Row = 1
    sSpread.Row2 = sSpread.DataRowCnt
    sSpread.Col = 1
    sSpread.Col2 = sSpread.MaxCols
    sSpread.BlockMode = True
    sSpread.Action = ActionClearText
    sSpread.BlockMode = False
    
    
    
End Function

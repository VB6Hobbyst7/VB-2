Attribute VB_Name = "ModuleLIB"
Option Explicit

Public interfacfrm  As Form
Public gtable       As Recordset
Public commtable    As Recordset
Public identb       As Recordset
Public resulttb     As Recordset
Public FrmFlag      As Integer

Public FileName     As String
Public strmmdd      As String
Public textmmdd     As String
Public commstr      As String
Public codestr      As String
Public failcomm     As Integer
Public Porttag      As Integer
Public Title        As String
Public machstr      As String
Public machinit     As String
Public fileInit     As String
Public ImgClickkey  As Integer
Public Errkey       As Integer

Public Msg          As Integer
Public ddate        As Integer
Public IdleFlag     As Integer
Public ResltFlag    As Integer
Public PendFlag     As Integer
Public OrderFlag    As Integer
Public OrderCnt     As Integer
Public slipno       As String
Public slipbuff     As String
Public MaxTestItem  As Integer


Type commset
    Port        As Integer
    data_bit    As Integer
    stop_bit    As Integer
    baud_rate   As Integer
    parity      As String
    blocksize   As Integer
End Type

Type TestNameTbl
    name        As String
    code        As String
    col_cnt     As Integer
End Type


Function GetByOne(ByVal sStr As String, sOriginal As String) As String
    Dim pos%
    
    pos = InStr(sStr, "|")
    
    If pos = 0 Then
    Else
        GetByOne = Trim(Mid$(sStr, 1, pos - 1))
        sOriginal = Trim(Mid$(sOriginal, pos + 1, Len(sOriginal) - pos))
    End If

End Function

Function Row_Plus(SpreadName As Object) As Integer
'1) 텍스트가 존재하는 현재의 MaxRow를 구함.
'2) 텍스트가 존재하는 현재의 MaxRow가 Spread의 MaxRow와 같으면 MaxRow + 1
    Dim i%
    Dim vTmp
    
    Row_Plus = 0
    
    If SpreadName.MaxRows = 0 Then
        SpreadName.MaxRows = 1
        Exit Function
    End If
    
    For i = 1 To SpreadName.MaxRows
        Call SpreadName.GetText(1, i, vTmp)
        
        If vTmp = "" Then
            Exit For
        Else
            Row_Plus = i
        End If
    Next
    
    If SpreadName.MaxRows = Row_Plus Then
        SpreadName.MaxRows = SpreadName.MaxRows + 1
    End If
    
'''    If SpreadName.Row >= SpreadName.MaxRows Then
'''        SpreadName.MaxRows = SpreadName.MaxRows + 1
'''        SpreadName.Row = SpreadName.Row + 1
'''    Else
'''        SpreadName.Row = SpreadName.Row + 1
'''    End If
End Function


Public Sub spdChksettext(SpreadName As Object, Colposition As Integer, RowPosition As Integer, byunsu As Variant)

    SpreadName.Col = Colposition
    SpreadName.Row = RowPosition
    SpreadName.TypeCheckText = byunsu
    
End Sub
Public Sub txbox_highlight(ibox As TextBox)

    ibox.SelStart = 0
    ibox.SelLength = Len(ibox.Text)
    
End Sub
Public Sub spdsettext(SpreadName As Object, Colposition As Integer, RowPosition As Integer, byunsu As Variant)

    SpreadName.Col = Colposition
    SpreadName.Row = RowPosition
    If IsNull(byunsu) Then
        SpreadName.Text = ""
    Else
        SpreadName.Text = byunsu
    End If
    
End Sub

'Sub MainTitle_Bold(Index As Integer)
'    Dim i%
'
'    For i = 0 To 5
'        If i = Index Then
'            INTmain00.Lblmain(i).ForeColor = RGB(255, 0, 0)
'            INTmain00.Lblmain(i).FontBold = True
'        Else
'            INTmain00.Lblmain(i).ForeColor = RGB(0, 0, 0)
'            INTmain00.Lblmain(i).FontBold = False
'        End If
'    Next
'
'End Sub


Attribute VB_Name = "Comm"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'통신설정
Type COMConfig
    ComPort       As String
    Speed      As String
    Parity     As String
    DataBit    As String
    StopBit    As String
    StartBit   As String
    RTSEnable  As String
    DTREnable  As String
    ExamUID    As String
    Gubun      As String
    ConnectFlag As Boolean
End Type
Public CA_COM As COMConfig

Public gArrExam()

Public gGubun As Integer
Public Const gEquip = "CLINILOG"

Public gSugaCode As String
Public gResult As String

Type DB_Parm
    Driver  As String
    User    As String
    Passwd  As String
    Server  As String
    DB      As String
    HostName    As String
End Type
Public gDB_Parm As DB_Parm

Type ExamRes
    res As String
    EquipRes As String
    RefLow As String
    RefHigh As String
    RefFlag As String
    EquipCode  As String
    ExamCode  As String
    ExamName As String
    SeqNo As String
    EquipGubun As String
    Equipcomment As String
End Type
Public gArrExamRes() As ExamRes

Public gExpireDate As String

Public raw_data As String

Public gArr_ExamCode() As String

Public gRow As Integer
Public gCol As Integer

Public gInsCode As String

Public gServerID As String

Public gUserID  As String

Public gIFUser  As String

Public Function chrSTX() As String
    chrSTX = Chr(2)
End Function

Public Function chrETX() As String
    chrETX = Chr(3)
End Function

Public Function chrEOT() As String
    chrEOT = Chr(4)
End Function

Public Function chrENQ() As String
    chrENQ = Chr(5)
End Function

Public Function chrACK() As String
    chrACK = Chr(6)
End Function

Public Function chrTAB() As String
    chrTAB = Chr(9)
End Function

Public Function chrLF() As String
    chrLF = Chr(10)
End Function

Public Function chrCR() As String
    chrCR = Chr(13)
End Function

Public Function chrNACK() As String
    chrNACK = Chr(15)
End Function

Public Function chrSPC() As String
    chrSPC = Chr(20)
End Function

Public Function chrETB() As String
    chrETB = Chr(23)
End Function

Public Function CheckSum(ByVal CheSum As String) As String
    Dim Tot  As Integer
    Dim sStr As String
    Dim i As Integer
    
    For i = 1 To Len(CheSum)
        Tot = Tot + Asc(Mid(CheSum, i, 1))
    Next i
    
    sStr = Hex(Tot)
    If Len(sStr) = 1 Then
        sStr = "0" & sStr
    End If
    
    CheckSum = Right(sStr, 2)
End Function

Public Function Get_OrderBody(strSid As String) As Variant

'Set cmdSQL = New ADODB.Command

On Error GoTo errtrap

    With cmdSQL
        .ActiveConnection = cn_Ser
        .CommandType = adCmdStoredProc
        .CommandText = "InterfaceBarcode_SELECT_sp2"
        .Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@i_barcodeNumber", adChar, adParamInput, 11, Trim(strSid))
        
        Set rs = New ADODB.Recordset
        rs.CursorType = adOpenStatic
        Set rs = .Execute
    End With
    
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Set cmdSQL = Nothing
        Get_OrderBody = Null
        
        Exit Function
    Else
        Get_OrderBody = rs.GetRows
        rs.Close
        Set rs = Nothing
        Set cmdSQL = Nothing
    End If
    
    Exit Function
    
errtrap:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing
    
    Get_OrderBody = Null
End Function

Public Function Get_OrderBody1(strSid As String) As Variant

'Set cmdSQL = New ADODB.Command

On Error GoTo errtrap

    With cmdSQL
        .ActiveConnection = cn_Ser
        .CommandType = adCmdStoredProc
        .CommandText = "InterfaceBarcode_SELECT_sp"
        .Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@i_barcodeNumber", adChar, adParamInput, 11, Trim(strSid))
        
        Set rs = New ADODB.Recordset
        rs.CursorType = adOpenStatic
        Set rs = .Execute
    End With
    
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Set cmdSQL = Nothing
        Get_OrderBody1 = Null
        
        Exit Function
    Else
        Get_OrderBody1 = rs.GetRows
        rs.Close
        Set rs = Nothing
        Set cmdSQL = Nothing
    End If
    
    Exit Function
    
errtrap:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing
    
    Get_OrderBody1 = Null
End Function


Attribute VB_Name = "COBAS_H232"
Option Explicit

Public Const H232_Connect = 1
Public Const H232_CState = 2
Public Const H232_InstrumentName = 3
Public Const H232_ModelID = 4
Public Const H232_SerialNum = 5

Public H232_State As String
Public H232_Connect_state As Boolean
Public H232_Output As String


Public H232_State_2 As String
Public H232_Connect_state_2 As Boolean
Public H232_Output_2 As String

Public H232_1 As String
Public H232_s_1 As String

Public H232_2 As String
Public H232_s_2 As String

Public Function H232_Function_2(asFunction As Integer) As String
    Dim lsSendData As String
    
    H232_Function_2 = ""
    lsSendData = ""
    
    Select Case asFunction
    Case H232_Connect
        lsSendData = chrCAN
'''        H232_State_2 = H232_Connect
        H232_2 = "1"
        
    Case H232_CState
        lsSendData = chrVT '& chrCR
        H232_2 = "2"
'''    Case H232_InstrumentName
'''        lsSendData = "I" '& chrCR
'''        H232_State_2 = H232_InstrumentName
'''    Case H232_ModelID
'''        lsSendData = "C" & Chr(9) & "4" & chrCR
'''        H232_State_2 = H232_ModelID
    Case H232_SerialNum
        lsSendData = "C" & Chr(9) & "3" & chrCR
        H232_State_2 = H232_SerialNum
        H232_2 = "2"
    End Select
    
    H232_Function_2 = lsSendData
End Function

Public Function H232_Function(asFunction As Integer) As String
    Dim lsSendData As String
    
    H232_Function = ""
    lsSendData = ""
    
    Select Case asFunction
    Case H232_Connect
        lsSendData = chrCAN
        H232_State = H232_Connect
        H232_1 = "1"
        
    Case H232_CState

        lsSendData = chrVT
        H232_1 = "2"
'''        H232_State = H232_CState
'''
'''    Case H232_InstrumentName
'''
'''        lsSendData = "I" & chrCR
'''        H232_State = H232_InstrumentName
'''
'''    Case H232_ModelID
'''        lsSendData = "C" & Chr(9) & "4" & chrCR
'''        H232_State = H232_ModelID
    Case H232_SerialNum
'''        lsSendData = "C" & Chr(9) & "3" & chrCR
'''        H232_1 = "2"
'''        H232_State = H232_SerialNum
    End Select
    
    H232_Function = lsSendData
End Function

Public Function H232_Result_Request(asStartSeq As String, asEndSeq As String) As String
'"a" & Chr(9) & Trim(Text1.Text) & Chr(9) & Trim(Text2.Text) & Chr(9) & "0" & vbCr
    Dim lsSendData As String
    H232_1 = "3"
'''    H232_Result_Request = ""
'''    lsSendData = ""
'''    lsSendData = "a" & Chr(9) & Trim(asStartSeq) & Chr(9) & Trim(asEndSeq) & Chr(9) & "0" & vbCr
'''    H232_Result_Request = lsSendData
    
End Function

Public Function H232_Result_Request_1(asStartSeq As String, asEndSeq As String) As String
'"a" & Chr(9) & Trim(Text1.Text) & Chr(9) & Trim(Text2.Text) & Chr(9) & "0" & vbCr
    Dim lsSendData As String
    H232_2 = "3"
'''    H232_Result_Request = ""
'''    lsSendData = ""
'''    lsSendData = "a" & Chr(9) & Trim(asStartSeq) & Chr(9) & Trim(asEndSeq) & Chr(9) & "0" & vbCr
'''    H232_Result_Request = lsSendData
    
End Function


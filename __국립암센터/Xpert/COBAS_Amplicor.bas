Attribute VB_Name = "COBAS_Amplicor"
Option Explicit

Public State_Check As Integer
Public Integra_BC As String
Public Host_BC As String
Public ComState As Boolean

Public gWaitStart As String
Public gStandard As String
Public gWait As Long

Public gLastSendData As String
Public gDays As String

Public Function Return_Init() As String
    Return_Init = "15 396612           00"
    'Return_Init = Chr(comSOH) & End_Char & _
                  "14 COBAS INTEGRA400 00" & End_Char & _
                  Chr(comSTX) & End_Char & _
                  Chr(comETX) & End_Char & _
                  Chr(comEOT) & End_Char
    
End Function

Public Function End_Char() As String
'기계의 세팅을 LF 로 할경우는 comLF
    'End_Char = Chr(comCR) & Chr(comLF)
    End_Char = chrLF
End Function

Public Sub Amplicor_INIT()
    Dim sSenData As String
    
    'sSenData = Chr(1) & End_Char & _
              "15 396612           00" & End_Char & _
              Chr(2) & End_Char & _
              Chr(3) & End_Char & _
              Chr(4) & End_Char
    sSenData = Chr(1) & End_Char & _
              "15 COBAmplicor Host 00" & End_Char & _
              Chr(2) & End_Char & _
              Chr(3) & End_Char & _
              Chr(4) & End_Char
    gLastSendData = sSenData
    Host_BC = "00"

    frmInterface.MSComm1.Output = sSenData
    SaveData "[TX:" & Format(Time, "hh:nn:ss") & "]" & sSenData
End Sub

Public Sub Amplicor_Res_Req()
    Dim sSenData As String
    
    If Trim(gWaitStart) = "" Then
        gWaitStart = Time
    End If
    
    sSenData = Chr(1) & End_Char & _
            "15 COBAmplicor Host 09" & End_Char & _
            Chr(2) & End_Char & _
            "00 0" & End_Char & _
            Chr(3) & End_Char & _
            Chr(4) & End_Char
    gLastSendData = sSenData
    Host_BC = "09"
    
    frmInterface.MSComm1.Output = sSenData
    SaveData "[TX:" & Format(Time, "hh:nn:ss") & "]" & sSenData
'    If Trim(gWaitStart) = "" Then
'        gWaitStart = Time
'    End If
End Sub

Public Sub Amplicor_OrderID_Req(Optional aiSelection As Integer = 1)
    Dim sSenData As String
    
    If aiSelection < 0 Or aiSelection > 4 Then
        aiSelection = 1
    End If
    
    Host_BC = "60_" & CStr(aiSelection)
    
    If Trim(gWaitStart) = "" Then
        gWaitStart = Time
    End If
    
    sSenData = Chr(1) & End_Char & _
            "09 36-2450          60" & End_Char & _
            Chr(2) & End_Char & _
            "40 " & CStr(aiSelection) & End_Char & _
            Chr(3) & End_Char & _
            Chr(4) & End_Char
    gLastSendData = sSenData
    Host_BC = "60"
        
    'SaveQuery sSenData
    
    frmInterface.MSComm1.Output = sSenData
    
    SaveData "[TX:" & Format(Time, "hh:nn:ss") & "]" & sSenData
End Sub

Public Sub Integra400_Order_Entry(ByVal asOrderID As String, asDate As String, ByVal asTestID As String, Optional asType As String = "")
    Dim sSenData As String
    
    sSenData = Chr(1) & End_Char & _
                "09 36-2450          10" & End_Char & _
                Chr(2) & End_Char & _
                "53 " & SetSpace(asOrderID, 15, 2) & " " & asDate & " " & Trim(asType) & End_Char & _
                "55 " & asTestID & End_Char & _
                Chr(3) & End_Char & _
                Chr(4) & End_Char
    gLastSendData = sSenData
    Host_BC = "10"
    
    frmInterface.MSComm1.Output = sSenData
    SaveData "[TX:" & Format(Time, "hh:nn:ss") & "]" & sSenData
End Sub

Public Sub Amplicor_Order_Entry(ByVal asOrder As String)
    Dim sSenData As String
    
    sSenData = Chr(1) & End_Char & _
                "15 COBAmplicor Host 10" & End_Char & _
                Chr(2) & End_Char & _
                asOrder & _
                Chr(3) & End_Char & _
                Chr(4) & End_Char
    gLastSendData = sSenData
    'Debug.Print sSenData
    Host_BC = "10"

    frmInterface.MSComm1.Output = sSenData
    SaveData "[TX:" & Format(Time, "hh:nn:ss") & "]" & sSenData
End Sub



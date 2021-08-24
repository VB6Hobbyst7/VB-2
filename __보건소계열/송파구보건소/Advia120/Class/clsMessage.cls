VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private objAnalysis     As clsMsg_Analysis
Private objInit         As clsMsg_Init
Private objOrder        As clsMsg_Order
'Private objOrderN       As clsMsg_OrderNo
'Private objOrderV       As clsMsg_OrderValid
Private objQuery        As clsMsg_Query
Private objResult       As clsMsg_Result
'Private objResultV      As clsMsg_ResultValid
Private objToken        As clsMsg_Token

Public Function Broadcast(ByVal Message As String) As String
    Dim stxPos      As Long
    Dim stxExcept   As String
    Dim strMsgType  As String * 1
    
    stxPos = InStr(1, Message, Chr(COM_STX))
    stxExcept = Mid(Message, stxPos)
    
    CHR_MT = Mid(stxExcept, 2, 1)
    strMsgType = Mid(stxExcept, 3, 1)
    
    Set objAnalysis = New clsMsg_Analysis
    
    Select Case strMsgType
        Case MID_INIT                   'As String = "I"
'            Set objInit = New clsMsg_Init
'            objAnalysis = objInit
'            Set objInit = Nothing
        Case MID_TOKEN                  'As String = "S"
            Set objToken = New clsMsg_Token
            
            With objToken
                .MSG_TOKEN = stxExcept
            End With
            Set objAnalysis = objToken
            
            Set objToken = Nothing
        Case MID_ORDER                  'As String = "Y"
            'Set objOrder = New clsMsg_Order
            
        Case MID_VALIDO                 'As String = "E"
'            Set objOrderV = New clsMsg_OrderValid
'            With objOrderV
'                .MSG_VALIDOREDER = stxExcept
'            End With
'            Set objAnalysis = objOrderV
'            Set objOrderV = Nothing
'            Call Send_Token
        Case MID_RESULT                 'As String = "R"
            Call Stop_Token
            Call Com_Output(CHR_MT)
            
            Set objResult = New clsMsg_Result
            Set frmComm.Result = objResult
            With objResult
                .MSG_RESULT = stxExcept
            End With
            
            Set objAnalysis = objResult
            Set objResult = Nothing
            
            Call Send_Token
            
        Case MID_QUERY                  'As String = "Q"
            Call Stop_Token
            Call Com_Output(CHR_MT)
            Set objQuery = New clsMsg_Query
            With objQuery
                .MSG_QUERY = stxExcept
            End With
            Set objAnalysis = objQuery
            
            Set objQuery = Nothing
            
        Case MID_ORDERN 'As String = "N"
            'Set objOrderN = New clsMsg_OrderNo
            
        Case MID_VALIDR 'As String = "Z"
            'Set objResultV = New clsMsg_ResultValid
            
        Case Else
        
    End Select
    
    Broadcast = objAnalysis.GetMessage
    
    Set objAnalysis = Nothing
    
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg_Header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ASTM Header"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mHEADER(12)     As String

Private vHeader()     As String

Private Sub Class_Initialize()
    mHEADER(0) = MSG_H                            '*Record Type
    mHEADER(1) = "\^&"                          '*Delimiters
    mHEADER(2) = ""                             '-Message Control ID
    mHEADER(3) = ""                             '-Access Password
    mHEADER(4) = ""                             'Sender name Or ID
    mHEADER(5) = ""                             'Sender Address
    mHEADER(6) = ""                             'Reserved
    mHEADER(7) = ""                             'Sender Telephone
    mHEADER(8) = ""                             'Charateristics of Sender
    mHEADER(9) = ""                             'Receiver ID
    mHEADER(10) = ""                            'Comment
    mHEADER(11) = "P"                           'Processing ID
    mHEADER(12) = "1"                           'Version Number
End Sub

Public Property Get MSG_HERDER() As String
    Dim strTemp     As String
    
    strTemp = Join(mHEADER, "|")
    strTemp = "1" & strTemp & Chr(COM_CR) & Chr(COM_ETX)
    
'   strTemp = "1H|\^&||||||||||P|1" & Chr(COM_CR) & Chr(COM_ETX)
    MSG_HERDER = Chr(COM_STX) & strTemp & CHK_SUM(strTemp, 2) & Chr(COM_CR) & Chr(COM_LF)
    
End Property

Public Property Let MSG_HERDER(ByVal sValue As String)
    Dim i As Integer
    vHeader = Split(sValue, "|")
    
    For i = LBound(vHeader) To UBound(vHeader)
        If i = UBound(mHEADER) Then Exit For
        mHEADER(i) = vHeader(i)
    Next
End Property

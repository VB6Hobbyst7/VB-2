VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "POCT "
   ClientHeight    =   9210
   ClientLeft      =   2355
   ClientTop       =   1725
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   10470
   Begin VB.CommandButton Command4 
      Caption         =   "오더"
      Height          =   375
      Left            =   8490
      TabIndex        =   15
      Top             =   360
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "장비코드요청"
      Height          =   375
      Left            =   7020
      TabIndex        =   14
      Top             =   390
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "서버접속"
      Height          =   345
      Left            =   5520
      TabIndex        =   13
      Top             =   420
      Width           =   1485
   End
   Begin VB.TextBox txtEquip 
      Height          =   480
      Left            =   4035
      TabIndex        =   11
      Text            =   "H05"
      Top             =   405
      Width           =   1395
   End
   Begin VB.TextBox txtXMLedit 
      Height          =   2025
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3300
      Width           =   9135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "clear"
      Height          =   315
      Left            =   7560
      TabIndex        =   9
      Top             =   780
      Width           =   1725
   End
   Begin VB.TextBox txtBarNo 
      Height          =   480
      Left            =   1200
      TabIndex        =   7
      Top             =   420
      Width           =   1395
   End
   Begin VB.TextBox txtSend 
      Height          =   3150
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5640
      Width           =   9060
   End
   Begin VB.TextBox txtXML 
      Height          =   2025
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   9135
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "XML Request"
      Height          =   315
      Left            =   5670
      TabIndex        =   2
      Top             =   750
      Width           =   1725
   End
   Begin VB.TextBox txtServerPath 
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Top             =   90
      Width           =   8055
   End
   Begin VB.Label Label5 
      Caption         =   "Equip : "
      Height          =   435
      Left            =   2925
      TabIndex        =   12
      Top             =   435
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "Barcode : "
      Height          =   435
      Left            =   90
      TabIndex        =   8
      Top             =   450
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "Send XML"
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   5415
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Return XML"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "ServerPath : "
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strXML As String

Private Sub cmdTest_Click()
    
  Dim XMLRequest            As New XMLHTTPRequest
  Dim txtSendXML               As String
  Dim txtResponseHeaders    As String
  Dim txtResponse           As String
  
  
On Error GoTo err_handler
  
    txtSendXML = "<?xml version='1.0' encoding='UTF-8'?>"
    txtSendXML = txtSendXML & vbCrLf & "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>"
    txtSendXML = txtSendXML & vbCrLf & "<soapenv:Body>"
    txtSendXML = txtSendXML & vbCrLf & "<registSpcmRcpn xmlns='http://svc.poct.ws.nhimc/'>"
    txtSendXML = txtSendXML & vbCrLf & "<arg0 xmlns=''>" & Trim(txtBarNo.Text) & "</arg0>"
    txtSendXML = txtSendXML & vbCrLf & "<arg1 xmlns=''>" & Trim(txtEquip.Text) & "</arg1>"
    txtSendXML = txtSendXML & vbCrLf & "</registSpcmRcpn>"
    txtSendXML = txtSendXML & vbCrLf & "</soapenv:Body>"
    txtSendXML = txtSendXML & vbCrLf & "</soapenv:Envelope>" & vbCrLf
      
    txtSend.Text = txtSendXML

    XMLRequest.open "POST", txtServerPath.Text, False
    XMLRequest.setRequestHeader "Content-Type", "text/xml"
    '  o.setRequestHeader "Connection", "close"
    XMLRequest.setRequestHeader "Connection", "PoctService"
    XMLRequest.setRequestHeader "SOAPAction", ""
    XMLRequest.send txtSendXML
    
    'XMLRequest.send "1607000010"
    
    
    txtResponseHeaders = XMLRequest.getAllResponseHeaders
    txtResponse = XMLRequest.responseText

    txtXML.Text = txtResponse
    txtResponse = Replace(txtResponse, ">", ">" & vbNewLine)
    txtXMLedit.Text = txtResponse
    
err_handler:
  If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description
    
    
End Sub

Private Sub Command1_Click()
    txtBarNo.Text = ""
    txtXML.Text = ""
    txtXMLedit.Text = ""
    txtSend.Text = ""
End Sub

Private Sub Command2_Click()
  Dim XMLRequest            As New XMLHTTPRequest
  Dim txtSendXML               As String
  Dim txtResponseHeaders    As String
  Dim txtResponse           As String
  
  
On Error GoTo err_handler

    Dim lsConnectReq As String
    Dim lsEquipCDReq As String
    Dim lsBarcodeReq As String
    
    lsConnectReq = "http://his032.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00113&business_id=lis&bcno=&testcd=&eqmtcd=H05&instcd=032&"
    lsEquipCDReq = "http://his032.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00103&business_id=lis&refgbn=2&instcd=032&eqmtcd=H05&"
    
    txtServerPath = lsConnectReq
    
    XMLRequest.open "POST", lsConnectReq, False
    'XMLRequest.setRequestHeader "Content-Type", "text/xml"
    'XMLRequest.setRequestHeader "Connection", "PoctService"
    'XMLRequest.setRequestHeader "SOAPAction", ""
    XMLRequest.send ""
    
    'XMLRequest.send "1607000010"
    
    
    'txtResponseHeaders = XMLRequest.getAllResponseHeaders
    txtResponse = XMLRequest.responseText
    
    txtXML.Text = txtResponse
    txtResponse = Replace(txtResponse, ">", ">" & vbNewLine)
    txtXMLedit.Text = txtResponse
    
err_handler:
  If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description
    
    
End Sub

Private Sub Command3_Click()
  Dim XMLRequest            As New XMLHTTPRequest
  Dim txtSendXML               As String
  Dim txtResponseHeaders    As String
  Dim txtResponse           As String
  
  
On Error GoTo err_handler

    Dim lsConnectReq As String
    Dim lsEquipCDReq As String
    Dim lsBarcodeReq As String
    
    lsEquipCDReq = "http://his032.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00103&business_id=lis&refgbn=2&instcd=032&eqmtcd=" & Trim(txtEquip.Text) & "&"
    
    txtServerPath = lsEquipCDReq
    
    XMLRequest.open "POST", lsEquipCDReq, False
    'XMLRequest.setRequestHeader "Content-Type", "text/xml"
    'XMLRequest.setRequestHeader "Connection", "PoctService"
    'XMLRequest.setRequestHeader "SOAPAction", ""
    XMLRequest.send ""
    
    'XMLRequest.send "1607000010"
    
    
    'txtResponseHeaders = XMLRequest.getAllResponseHeaders
    txtResponse = XMLRequest.responseText
    
    txtXML.Text = txtResponse
    
    txtResponse = Replace(txtResponse, ">", ">" & vbNewLine)
    txtXMLedit.Text = txtResponse
    
err_handler:
  If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub


Private Sub Command4_Click()
  Dim XMLRequest            As New XMLHTTPRequest
  Dim txtSendXML               As String
  Dim txtResponseHeaders    As String
  Dim txtResponse           As String
  
  
On Error GoTo err_handler

    Dim lsReqMsg As String
    
    lsReqMsg = "http://his032.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00101&business_id=lis&bcno=" & Trim(txtBarNo) & "&instcd=032&eqmtcd=" & Trim(txtEquip.Text) & "&"
    
    txtServerPath = lsReqMsg
    
    XMLRequest.open "POST", lsReqMsg, False
    'XMLRequest.setRequestHeader "Content-Type", "text/xml"
    'XMLRequest.setRequestHeader "Connection", "PoctService"
    'XMLRequest.setRequestHeader "SOAPAction", ""
    XMLRequest.send ""
    
    'XMLRequest.send "1607000010"
    
    
    'txtResponseHeaders = XMLRequest.getAllResponseHeaders
    txtResponse = XMLRequest.responseText
    
    SaveXMLFile txtResponse, "order"
    
    txtXML.Text = txtResponse
    
    txtResponse = Replace(txtResponse, ">", ">" & vbNewLine)
    txtXMLedit.Text = txtResponse
    
err_handler:
  If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description
  
End Sub



'New Result Trans sub end I08 사용 ===========================================================================
Public Sub SaveXMLFile(argXML As String, ByVal asGubun As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
        
    FilNum = FreeFile
    
    If Dir(App.Path & "\" & asGubun & ".xml") <> "" Then
        Kill App.Path & "\" & asGubun & ".xml"
    End If
    
    Open App.Path & "\" & asGubun & ".xml" For Append As FilNum
    Print #FilNum, argXML
    Close FilNum
    
End Sub


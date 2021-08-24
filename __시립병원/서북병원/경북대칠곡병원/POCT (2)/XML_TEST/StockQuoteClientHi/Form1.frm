VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   4215
   ClientTop       =   1950
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   6825
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   690
      Left            =   585
      TabIndex        =   0
      Top             =   810
      Width           =   1635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim client As Object
On Error GoTo ErrXML:
    Set client = CreateObject("MSSOAP.SoapClient")

'''    client.mssoapinit "http://localhost:8080/soap/StockQuoteService.wsdl", "StockQuoteService", "StockQuoteServicePort"
    client.mssoapinit "http://192.168.5.105:8800/service/PoctService?wsdl", "PoctService", "PoctPort"
    MsgBox client.getQuote("IBM")
ErrXML:
    MsgBox Err.Description
End Sub

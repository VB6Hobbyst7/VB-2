Attribute VB_Name = "StockQuoteClientHi"
Option Explicit

Sub Main()
    Dim client As Object

    Set client = CreateObject("MSSOAP.SoapClient")

    client.mssoapinit "http://localhost:8080/soap/StockQuoteService.wsdl", "StockQuoteService", "StockQuoteServicePort"
    MsgBox client.getQuote("IBM")
End Sub

VERSION 5.00
Begin VB.UserControl ControlURL 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ControlURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event GetUrlComplete(ByVal value As String)
Public Event GetUrlFailed()


Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    If AsyncProp.StatusCode = vbAsyncStatusCodeEndDownloadData Then
        RaiseEvent GetUrlComplete(StrConv(AsyncProp.value, vbUnicode))
        gStrXML = StrConv(AsyncProp.value, vbUnicode)
    Else
        RaiseEvent GetUrlFailed
    End If
End Sub

Public Sub Start(ByVal url As String)
    gStrXML = ""
    UserControl.AsyncRead url, vbAsyncTypeByteArray
End Sub


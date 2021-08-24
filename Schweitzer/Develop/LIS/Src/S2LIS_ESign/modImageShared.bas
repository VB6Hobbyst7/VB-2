Attribute VB_Name = "modImageShared"
Option Explicit

' Global variable to store reference to host application.
Public gobjAppInstance      As Object

' Public constants for menu item in Office application.
Public Const CBR_NAME       As String = "Tools"
Public Const CTL_CAPTION    As String = "Image &Gallery"
Public Const CTL_KEY        As String = "ImageGallery"
Public Const CTL_NAME       As String = "Image Gallery"

' Constants for characters surrounding ProgID.
Public Const PROG_ID_START  As String = "!<"
Public Const PROG_ID_END    As String = ">"

Sub AddInErr(errX As ErrObject)
    ' Displays message box with error information.

    Dim strMsg As String
    
    strMsg = "An error occurred in the COM add-in named '" _
        & App.Title & "'." & vbCrLf & "Error #:" & errX.Number _
        & vbCrLf & errX.Description
    MsgBox strMsg, , "Error!"
End Sub




Attribute VB_Name = "Mod_Function"
Option Explicit

Public Enum G_Color
    CWHITE = vbWhite
    CBLACK = vbBlack
    CDARKGRAY = &HC0C0C0
    CGRAY = &H8000000F
    CBLUE = &HFFC0C0   'vbBlue
    CGREEN = vbGreen
    CLIGHTGREEN = &HC0FFC0
    CRED = vbRed
    CPINK = &HC0E0FF
    CRIGHTRED = &H8080FF
    CDARKYELLOW = &HFFFF&
    CYELLOW = &HC0FFFF
    CORANGE = &H80FF&
    CMAGENTA = vbMagenta
End Enum

'API Declarations
Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

'API Structures
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'API constants
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80

' ****************************************************************************************************
' * Function Name: GfxSelectClear()
' * Description  : SpreadSheet에서 배경색 변환
' * Parameters   : objSpread (I)  SpreadSheet Object
' * Modification log :
' ****************************************************************************************************
Public Sub GfxSelectClear(objspread As Object, Optional ByVal nStartRow As Long = 1)
    Dim nRow As Long
    Dim nCol As Long
    On Error GoTo Err_Rtn
    Screen.MousePointer = vbHourglass
    With objspread
        For nRow = nStartRow To .MaxRows
            .Row = nRow
            .Col = 1
            If .BackColor <> vbWhite Then
                .Col = 1
                .Col2 = .MaxCols
                .Row = nRow
                .Row2 = nRow
                .BlockMode = True
                .BackColor = vbWhite
                .BlockMode = False
            End If
        Next
    End With
    Screen.MousePointer = vbNormal
    Exit Sub
Err_Rtn:
End Sub

' ****************************************************************************************************
' * Function Name: G_SET_SpreadColorChange()
' * Description  : SpreadSheet에서 Row 선택시 Row의 배경색 변환
' * Parameters   : objSpread (I)  SpreadSheet Object
' *                lngCol    (I)  SQL Script
' *                lngRow    (I)  SQL Script
' * Modification log :
' ****************************************************************************************************
Public Sub G_SET_SpreadColorChange(objspread As Object, LngCol As Long, LngRow As Long)
    On Error GoTo Err_Rtn
    With objspread
        .Row = LngRow
        .Row2 = LngRow
        .Col = 1
        .Col2 = .MaxCols
        .BlockMode = True
        If .BackColor = G_Color.CWHITE Then
            .BackColor = G_Color.CBLUE
        Else
            .BackColor = G_Color.CWHITE
        End If
        .BlockMode = False
        .BackColorStyle = 1
    End With
    Exit Sub
Err_Rtn:
    
End Sub
'***********************************************************************************
'***  Function Name : Mod_Function
'***  Description   : URL 연결 Module
'***  Function      : S_HomePage
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Public Sub S_HomePage(ByVal as_URL As String)

   Dim loIE As Object
   
   On Error Resume Next
   
   Set loIE = CreateObject("InternetExplorer.Application")
   loIE.Visible = True
   loIE.Navigate as_URL
   Set loIE = Nothing
   
End Sub

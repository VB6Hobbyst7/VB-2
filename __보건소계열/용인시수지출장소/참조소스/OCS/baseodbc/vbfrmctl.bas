Attribute VB_Name = "VbFrmCtl"
Option Explicit

Const HKEY_CURRENT_USER = &H80000001
Const Reg_Branch = "TWIN"

Private Type RECT

    Lft     As Long
    Top     As Long
    Rgt     As Long
    Bot     As Long
    
End Type

Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Integer, ByVal X As Integer, _
                                             ByVal Y As Integer, ByVal nWidth As Integer, _
                                             ByVal nHeight As Integer, ByVal hSrcDC As Integer, _
                                             ByVal XSrc As Integer, ByVal YSrc As Integer, _
                                             ByVal dwRop As Long) As Integer
                                             
Private Declare Function GetClientRect& Lib "USER32" (ByVal hwnd&, Rct As RECT)
Private Declare Function GetParent& Lib "USER32" (ByVal hwnd&)

Private Declare Function RegCloseKey& Lib "ADVAPI32" (ByVal hKey&)
Private Declare Function RegCreateKey& Lib "ADVAPI32" _
                  Alias "RegCreateKeyA" (ByVal hKey&, ByVal SubKey$, Result&)
Private Declare Function RegQueryValue& Lib "ADVAPI32" _
                  Alias "RegQueryValueA" (ByVal hKey&, ByVal SubKey$, _
                                          ByVal ReturnStr$, LenReturnstr&)
Private Declare Function RegSetValue& Lib "ADVAPI32" _
                  Alias "RegSetValueA" (ByVal hKey&, ByVal SubKey$, _
                                        ByVal StrType&, ByVal KeyValue$, ByVal KeyLen&)
                  
Sub CenterForm(Frm As Form)

    Dim nX              As Integer
    Dim nY              As Integer
    Dim nParWid         As Integer
    Dim nParHgt         As Integer
    Dim Rct             As RECT
    
    On Error GoTo Error_Process
    
    If Frm.MDIChild Then
        GetClientRect GetParent(Frm.hwnd), Rct
        nParWid = (Rct.Rgt - Rct.Lft) * Screen.TwipsPerPixelY
        nParHgt = (Rct.Bot - Rct.Top) * Screen.TwipsPerPixelX
        nX = (nParWid - Frm.Width) * 0.5
        nY = (nParHgt - Frm.Height) * 0.5
    Else
        nX = (Screen.Width - Frm.Width) * 0.5
        nY = (Screen.Height - Frm.Height) * 0.5
    End If
    
    Frm.Move nX, nY
    
    Exit Sub
    
'/-----------------------------------------------------------------------------

Error_Process:

    MsgBox "Center Form Error"
    
    Error = False
    
End Sub

Sub Picture_Copying(ArgFrPic As Object, ArgToPic As Object)

    Dim TileIt          As Integer
    
    TileIt = BitBlt(ArgToPic.hDC, 0, 0, ArgFrPic.Width, ArgFrPic.Height, ArgFrPic.hDC, 0, 0, &HCC0020)

End Sub

Sub Picture_Getting(ByVal ArgPic As Object, ArgRdoColumn As rdoColumn)

    Dim I                       As Integer
    Dim nPictureSize            As Long
    Dim nPicturePut()           As Byte
    Dim nChunk                  As Integer
    Const nFetchSize = 16384
    
    Open "c:\twin\test.bmp" For Binary Access Write As #1
    
    nPictureSize = ArgRdoColumn.ColumnSize
    
    nChunk = nPictureSize Mod nFetchSize
    ReDim nPicturePut(nChunk)
    nPicturePut() = ArgRdoColumn.GetChunk(nChunk)
    Put #1, , nPicturePut()

    nChunk = nPictureSize \ nFetchSize
    ReDim nPicturePut(nFetchSize)
    For I = 1 To nChunk
        nPicturePut() = ArgRdoColumn.GetChunk(nChunk)
        Put #1, , nPicturePut()
    Next I
    
    Close #1

    ArgPic.Picture = LoadPicture("")
    ArgPic.Picture = LoadPicture("c:\twin\test.bmp")
    
End Sub

Sub Picture_Setting(ByVal ArgPic As Object, ArgRdoColumn As rdoColumn)

    Dim I                       As Integer
    Dim nPictureSize            As Long
    Dim nPicturePut()           As Byte
    Dim nChunk                  As Integer
    Const nFetchSize = 16384
    
    SavePicture ArgPic.Picture, "c:\twin\test.bmp"
    
    ArgRdoColumn.AppendChunk Null
    
    Open "c:\twin\test.bmp" For Binary Access Read As #1
    nPictureSize = LOF(1)
    If nPictureSize = 0 Then Close #1: Exit Sub
    
    nChunk = nPictureSize Mod nFetchSize
    ReDim nPicturePut(nChunk)
    
    Get #1, , nPicturePut()
    ArgRdoColumn.AppendChunk nPicturePut()
    
    nChunk = nPictureSize \ nFetchSize
    ReDim nPicturePut(nFetchSize)
    
    For I = 1 To nChunk
        Get #1, , nPicturePut()
        ArgRdoColumn.AppendChunk nPicturePut()
    Next I
    
    Close #1
    
End Sub
Sub PreFormCheck()

    If App.PrevInstance = True Then
        Call MsgBox("This program is already running!", vbExclamation)
        End
    End If
    
End Sub

Sub ClearForm(ByVal ControlForm As Object)

    Dim I               As Integer
    
    For I = 0 To ControlForm.Count - 1
        
        If TypeOf ControlForm(I) Is TextBox Then ControlForm(I).Text = ""
        If TypeOf ControlForm(I) Is RichTextBox Then ControlForm(I).Text = ""
        If TypeOf ControlForm(I) Is ComboBox Then ControlForm(I).Clear
        If TypeOf ControlForm(I) Is ListBox Then ControlForm(I).Clear
        If TypeOf ControlForm(I) Is vaSpread Then
            If IsNumeric(ControlForm(I).Tag) And Val(ControlForm(I).Tag) <> 0 Then
                ControlForm(I).MaxRows = Val(ControlForm(I).Tag)
            End If
            ControlForm(I).Col = 1: ControlForm(I).Col2 = ControlForm(I).MaxCols
            ControlForm(I).Row = 1: ControlForm(I).Row2 = ControlForm(I).MaxRows
            ControlForm(I).BlockMode = True
            ControlForm(I).Action = SS_ACTION_CLEAR_TEXT
            ControlForm(I).BlockMode = False
        End If
    
    Next I
            
End Sub


Function Reg_Get(ByVal ArgSubKey As String) As String

    Dim nKeyHwnd            As Long
    Dim strInput            As String * 256
    
    If RegQueryValue(HKEY_CURRENT_USER, Reg_Branch & "\" & Trim(ArgSubKey), _
                     strInput, Len(strInput)) = 0 Then
        Reg_Get = MidB(strInput, 1, InStrB(strInput, Chr$(0)) - 1)
    Else
        Reg_Get = "NO"
    End If

End Function

Function Reg_Put(ByVal ArgSubKey$, ByVal ArgSubValue$) As String

    Dim nKeyHwnd            As Long
    
    RegCreateKey HKEY_CURRENT_USER, Reg_Branch, nKeyHwnd
    RegSetValue nKeyHwnd, Trim(ArgSubKey), 1, Trim(ArgSubValue), Len(Trim(ArgSubValue))

    RegCloseKey nKeyHwnd
    
End Function



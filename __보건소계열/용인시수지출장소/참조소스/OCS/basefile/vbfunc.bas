Attribute VB_Name = "VbFunction"
Option Explicit

'/========================================================================================================
Public Const IME_CMODE_NATIVE = &H1
Public Const IME_CMODE_HANGEUL = IME_CMODE_NATIVE
Public Const IME_CMODE_ALPHANUMERIC = &H0
Public Const IME_SMODE_NONE = &H0
Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

Public GstrSysDate          As String
Public GnAge_YY             As Integer
Public GnAge_MM             As Integer
Public GnAge_DD             As Integer

Sub Age_Gesan(ByVal SysDate, ByVal BirthDate)

    On Error GoTo Gesan_error
    
    GnAge_YY = 0
    GnAge_MM = 0
    GnAge_DD = 0

    GnAge_YY = DateDiff("yyyy", BirthDate, SysDate)
    GnAge_MM = DateDiff("m", BirthDate, SysDate)
    GnAge_DD = DateDiff("d", BirthDate, SysDate)

    Exit Sub
    

'/-------------------------------------------------------------------------------/

Gesan_error:
    
    GnAge_YY = 0
    GnAge_MM = 0
    GnAge_DD = 0

End Sub


Public Sub cvtToHan(ByRef ArgObject As Object)
   Dim hIMC                 As Long
   
   hIMC = ImmGetContext(ArgObject.hwnd)
   ImmSetConversionStatus hIMC, IME_CMODE_HANGEUL, IME_SMODE_NONE
   
End Sub

Public Sub cvtToEng(ByRef ArgObject As Object)
   Dim hIMC                 As Long
   
   hIMC = ImmGetContext(ArgObject.hwnd)
   ImmSetConversionStatus hIMC, IME_CMODE_ALPHANUMERIC, IME_SMODE_NONE
   
End Sub

Public Function Date_Format(ByVal AnyDateFormat As String) As String

    Dim strDate             As String
    
    AnyDateFormat = Trim(AnyDateFormat)

    If IsNumeric(AnyDateFormat) Then
        Select Case Len(AnyDateFormat)
            Case 4
                strDate = Left(AnyDateFormat, 2)
                strDate = strDate & "-" & Mid(AnyDateFormat, 3, 2)
            Case 6
                strDate = Left(AnyDateFormat, 2)
                strDate = strDate & "-" & Mid(AnyDateFormat, 3, 2)
                strDate = strDate & "-" & Mid(AnyDateFormat, 5, 2)
            Case 8
                strDate = Left(AnyDateFormat, 4)
                strDate = strDate & "-" & Mid(AnyDateFormat, 5, 2)
                strDate = strDate & "-" & Mid(AnyDateFormat, 7, 2)
        End Select
    Else
        strDate = AnyDateFormat
    End If
    
    If IsDate(strDate) Then
        Date_Format = Format(strDate, "YYYY-MM-DD")
    Else
        Date_Format = ""
    End If

End Function

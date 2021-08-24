VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Escer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum UnEscStates
    uesNotInEscapeSeq = 0
    uesProcessingEscSeq = 1
    uesProcessingLongEscSeq = 2
    uesProcessingLongEscSeqXEnd = 3
    uesProcessingLongEscSeqUEnd = 5
End Enum

Public Function UnEscape(ByVal EscapedText As String) As String
    'Un-escape a string value.  Escape sequences supported:
    '
    '   ^xhh   7 bit hex               ^r     vbCr
    '   ^uhhhh 16 bit hex ("Unicode")  ^l     vbLf
    '   ^b     vbBack                  ^n     vbNewLine/vbCrLf
    '   ^t     vbTab                   ^q     quote (")
    '   ^v     vbVerticalTab           ^0     vbNullChar
    '   ^f     vbFormFeed              ^^     ^
    
    Dim Pos As Long
    Dim ChS As String
    Dim ChW As Integer
    Dim UpChW As Integer
    Dim UnEscState As UnEscStates
    Dim UnPos As Long
    Dim Unicode As Boolean
    Dim Accum As Long
    
    UnEscape = Space$(Len(EscapedText))
    UnPos = 1&
    For Pos = 1& To Len(EscapedText)
        ChS = Mid$(EscapedText, Pos, 1&)
        ChW = AscW(ChS)
        UpChW = AscW(UCase$(ChS))
        Select Case UnEscState
            Case uesNotInEscapeSeq
                If ChW = 94# Then ' "^"
                    UnEscState = uesProcessingEscSeq
                End If
            Case uesProcessingEscSeq
                UnEscState = uesNotInEscapeSeq
                Select Case UpChW
                    Case 88# ' "X"
                        UnEscState = uesProcessingLongEscSeq
                    Case 85# ' "U"
                        Unicode = True
                        UnEscState = uesProcessingLongEscSeq
                    Case 66# ' "B"
                        ChS = vbBack
                    Case 84# ' "T"
                        ChS = vbTab
                    Case 86# ' "V"
                        ChS = vbVerticalTab
                    Case 70# ' "F"
                        ChS = vbFormFeed
                    Case 82# ' "R"
                        ChS = vbCr
                    Case 76# ' "L"
                        ChS = vbLf
                    Case 78# ' "N"
                        Mid$(UnEscape, UnPos, 1&) = vbCr
                        UnPos = UnPos + 1
                        ChS = vbLf
                    Case 81# ' "Q"
                        ChS = """"
                    Case 48# ' "0"
                        ChS = vbNullChar
                    Case 94# ' "^", i.e. ^^
                        '
                    Case Else
                        Err.Raise 5
                End Select
            Case Else
                If (48# <= UpChW And UpChW <= 57#) Then
                    Accum = 16& * Accum + (CLng(UpChW) And &HF&)
                ElseIf (65# <= UpChW And UpChW <= 70#) Then
                    Accum = 16& * Accum + (CLng(UpChW) - 55&)
                Else
                    Err.Raise 5
                End If
                UnEscState = UnEscState + 1&
                If Unicode Then
                    If UnEscState > uesProcessingLongEscSeqUEnd Then
                        ChS = ChrW$(Accum)
                        Accum = 0&
                        UnEscState = uesNotInEscapeSeq
                        Unicode = False
                    End If
                Else
                    If Accum > 127 Then Err.Raise 5
                    If UnEscState > uesProcessingLongEscSeqXEnd Then
                        ChS = ChrW$(Accum)
                        Accum = 0&
                        UnEscState = uesNotInEscapeSeq
                    End If
                End If
        End Select
        If UnEscState = uesNotInEscapeSeq Then
            Mid$(UnEscape, UnPos, 1&) = ChS
            UnPos = UnPos + 1&
        End If
    Next
    If UnEscState <> uesNotInEscapeSeq Then Err.Raise 5
    UnEscape = LEFT$(UnEscape, UnPos - 1&)
End Function


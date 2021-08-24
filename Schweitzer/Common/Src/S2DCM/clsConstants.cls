VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConstants"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public CS_DateMask As String
Public CS_DateLongMask As String
Public CS_DateShortMask As String
Public CS_TimeShortMask As String
Public CS_TimeLongMask As String
Public CS_BlankMask As String
Public CS_DateShortFormat As String
Public CS_DateLongFormat As String
Public CS_TimeShortFormat As String
Public CS_TimeLongFormat As String
Public CS_DateDbFormat As String
Public CS_TimeDbFormat As String

Public LINE_DIV  As String
Public COL_DIV  As String

Public CS_FingerMark As String
Public CS_QuestionMark As String

Private Sub Class_Initialize()

    CS_DateMask = "0###-##-##"
    CS_DateLongMask = "0###/##/##"
    CS_DateShortMask = "0#-##-##"
    CS_TimeLongMask = "0#:##:##"
    CS_TimeShortMask = "0#:##"
    CS_BlankMask = "____/__/__"
    CS_DateShortFormat = "YY-MM-DD"
    CS_DateLongFormat = "YYYY-MM-DD"
    CS_TimeShortFormat = "HH:MM"
    CS_TimeLongFormat = "HH:MM:SS"
    CS_DateDbFormat = "YYYYMMDD"
    CS_TimeDbFormat = "HHMMSS"
    
    LINE_DIV = Chr(17)
    COL_DIV = Chr(19)
    
    CS_FingerMark = "ขั"
    CS_QuestionMark = "?"
    
End Sub

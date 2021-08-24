VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTLab023"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public CtrlCd As String    '-- Control코드
Public LevelCd As String   '-- Level
Public LotNo As String     '-- Lot Number

Public OpenDt As String    '-- 시작일
Public ExpDt As String     '-- 폐기일
Public MakeCd As String    '-- 제조사
Public Remark As String    '-- Remark


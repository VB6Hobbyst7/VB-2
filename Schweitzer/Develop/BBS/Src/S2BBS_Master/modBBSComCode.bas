Attribute VB_Name = "modBBSComCode"
Option Explicit

Global DbConn As DrDatabase
Global objBBSComCode As clsHosComCode
Global objSysInfo As clsS2DSO
Global objMyUser As clsDSMLogOn

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long





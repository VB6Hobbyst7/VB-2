VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarDoDt As String      '--������
Private mvarDoTm As String      '--����ð�
Private mvarSectCd As String    '--Section�ڵ�
Private mvarCtrlCd As String    '--Control�ڵ�
Private mvarCtrlNm As String    '--Control�ڵ�
Private mvarLevelCd As String   '--Level(H, M, L)
Private mvarPrtFg As String     '--��¿���
Private mvarPrtId As String     '--������ID
Private mvarDoneFg As String    '--���࿩��

Private mvarNewFg As Boolean    '--�űԿ���

'Private mvarMyDb As DrDatabase
Private mvarChanged As Boolean  '--Item���濩��


'Database Connection
'Public Property Get MyDb() As DrDatabase
'    Set MyDb = mvarMyDb
'End Property
'Public Property Set MyDb(ByVal vNewValue As DrDatabase)
'    Set mvarMyDb = vNewValue
'End Property

'������
Public Property Get DoDt() As String
    DoDt = mvarDoDt
End Property
Public Property Let DoDt(ByVal vNewValue As String)
    mvarDoDt = vNewValue
End Property

'����ð�
Public Property Get DoTm() As String
    DoTm = mvarDoTm
End Property
Public Property Let DoTm(ByVal vNewValue As String)
    mvarDoTm = vNewValue
End Property

'Section�ڵ�
Public Property Get SectCd() As String
    SectCd = mvarSectCd
End Property
Public Property Let SectCd(ByVal vNewValue As String)
    mvarSectCd = vNewValue
End Property

'Control�ڵ�
Public Property Get CtrlCd() As String
    CtrlCd = mvarCtrlCd
End Property
Public Property Let CtrlCd(ByVal vNewValue As String)
    mvarCtrlCd = vNewValue
End Property

'Level�ڵ�(H,M,L)
Public Property Get LevelCd() As String
    LevelCd = mvarLevelCd
End Property
Public Property Let LevelCd(ByVal vNewValue As String)
    mvarLevelCd = vNewValue
End Property

'Control��
Public Property Get CtrlNm() As String
    CtrlNm = mvarCtrlNm
End Property
Public Property Let CtrlNm(ByVal vNewValue As String)
    mvarCtrlNm = vNewValue
End Property

'��¿���
Public Property Get PrtFg() As String
    PrtFg = mvarPrtFg
End Property
Public Property Let PrtFg(ByVal vNewValue As String)
    mvarPrtFg = vNewValue
End Property

'������ID
Public Property Get PrtId() As String
    PrtId = mvarPrtId
End Property
Public Property Let PrtId(ByVal vNewValue As String)
    mvarPrtId = vNewValue
End Property

'���࿩��
Public Property Get DoneFg() As String
    DoneFg = mvarDoneFg
End Property
Public Property Let DoneFg(ByVal vNewValue As String)
    mvarDoneFg = vNewValue
End Property

Public Sub ScheduleSave()
    
    Dim SqlStmt As String
    
    SqlStmt = "insert into " & T_LAB025 & _
              "(dodt, dotm, sectcd, ctrlcd, levelcd, prtfg, prtid, donefg) " & _
              "values (" & _
                         DBV("dodt", mvarDoDt, 1) & DBV("dotm", mvarDoTm, 1) & _
                         DBV("sectcd", mvarSectCd, 1) & DBV("ctrlcd", mvarCtrlCd, 1) & _
                         DBV("levelcd", mvarLevelCd, 1) & DBV("prtfg", mvarPrtFg, 1) & _
                         DBV("prtid", mvarPrtId, 1) & DBV("donefg", mvarDoneFg) & _
                      ")"
    dbconn.Execute SqlStmt
    
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarCtrlCd As String    '--Control코드
Private mvarLevelCd As String   '--Level(H, M, L)
Private mvarCtrlNm As String    '--Control명
Private mvarEqpCd As String     '--장비코드
Private mvarEqpNm As String     '--장비명
Private mvarSectCd As String    '--Section 코드
Private mvarSectNm As String    '--Section 코드
Private mvarCtrlDiv As String   '--정도관리구분(내부/외부)
Private mvarWorkArea As String  '--WorkArea
Private mvarBuildCd As String   '--Building 코드

Private mvarNewFg As Boolean    '--신규여부

'Private mvarMyDb As DrDatabase
Private mItemCol As Collection  '--검사항목 컬렉션
Private mvarItemCount As Integer '--컬렉션 카운트
Private mSchedCol As Collection  '--Schedule 컬렉션
Private mvarSchedCnt As Integer '--Schedule 카운트
Private mvarChanged As Boolean  '--Item변경여부


''Database Connection
'Public Property Get MyDb() As DrDatabase
'    Set MyDb = mvarMyDb
'End Property
'Public Property Set MyDb(ByVal vNewValue As DrDatabase)
'    Set mvarMyDb = vNewValue
'End Property

'Control코드
Public Property Get CtrlCd() As String
    CtrlCd = mvarCtrlCd
End Property
Public Property Let CtrlCd(ByVal vNewValue As String)
    mvarCtrlCd = vNewValue
End Property

'Level코드(H,M,L)
Public Property Get LevelCd() As String
    LevelCd = mvarLevelCd
End Property
Public Property Let LevelCd(ByVal vNewValue As String)
    mvarLevelCd = vNewValue
End Property

'Control명
Public Property Get CtrlNm() As String
    CtrlNm = mvarCtrlNm
End Property
Public Property Let CtrlNm(ByVal vNewValue As String)
    mvarCtrlNm = vNewValue
End Property

'장비코드
Public Property Get EqpCd() As String
    EqpCd = mvarEqpCd
End Property
Public Property Let EqpCd(ByVal vNewValue As String)
    mvarEqpCd = vNewValue
End Property

'장비명
Public Property Get EqpNm() As String
    EqpNm = mvarEqpNm
End Property
Public Property Let EqpNm(ByVal vNewValue As String)
    mvarEqpNm = vNewValue
End Property

'Section코드
Public Property Get SectCd() As String
    SectCd = mvarSectCd
End Property
Public Property Let SectCd(ByVal vNewValue As String)
    mvarSectCd = vNewValue
End Property

'Section명
Public Property Get SectNm() As String
    SectNm = mvarSectNm
End Property
Public Property Let SectNm(ByVal vNewValue As String)
    mvarSectNm = vNewValue
End Property

'정도관리구분
Public Property Get CtrlDiv() As String
    CtrlDiv = mvarCtrlDiv
End Property
Public Property Let CtrlDiv(ByVal vNewValue As String)
    mvarCtrlDiv = vNewValue
End Property

'WorkArea
Public Property Get WorkArea() As String
    WorkArea = mvarWorkArea
End Property
Public Property Let WorkArea(ByVal vNewValue As String)
    mvarWorkArea = vNewValue
End Property

'Building Code
Public Property Get BuildCd() As String
    BuildCd = mvarBuildCd
End Property
Public Property Let BuildCd(ByVal vNewValue As String)
    mvarBuildCd = vNewValue
End Property

'신규여부
Public Property Get NewFg() As Boolean
    NewFg = mvarNewFg
End Property
Public Property Let NewFg(ByVal vNewValue As Boolean)
    mvarNewFg = vNewValue
End Property

'변경여부
Public Property Get Changed() As Boolean
    Changed = mvarChanged
End Property
Public Property Let Changed(ByVal vNewValue As Boolean)
    mvarChanged = vNewValue
End Property

'검사항목 컬렉션
Public Property Get ItemCount() As Integer
    ItemCount = mItemCol.Count
End Property

Public Property Get Item(ByVal KeyIndex As Variant) As clsControlItem
    Set Item = mItemCol(KeyIndex)
End Property

'스케쥴 컬렉션
Public Property Get SchedCnt() As Integer
    SchedCnt = mSchedCol.Count
End Property

Public Property Get Schedule(ByVal KeyIndex As Variant) As clsSchedule
    Set Schedule = mSchedCol(KeyIndex)
End Property


Private Sub Class_Initialize()
    Set mItemCol = New Collection
    Set mSchedCol = New Collection
    mvarItemCount = 0
    mvarSchedCnt = 0
End Sub

Public Function AddQcItem(ByVal clsKey As Variant, ByVal clsData As clsControlItem)
   
    Dim newItem As New clsControlItem
    
    With newItem
        newItem.CtrlCd = clsData.CtrlCd
        newItem.LevelCd = clsData.LevelCd
        newItem.TestCd = clsData.TestCd
        newItem.TestNm = clsData.TestNm
        newItem.EqpCd = clsData.EqpCd
    End With
    
    mItemCol.Add newItem
    Set AddQcItem = newItem
    mvarItemCount = mvarItemCount + 1
   
End Function

Public Sub ItemClear()
    Set mItemCol = Nothing
    Set mItemCol = New Collection
    mvarItemCount = 0
End Sub

Public Function AddSchedule(ByVal clsKey As Variant, ByVal clsData As clsSchedule)
   
    Dim newItem As New clsSchedule
    
    With newItem
        newItem.DoDt = clsData.DoDt
        newItem.DoTm = clsData.DoTm
        newItem.CtrlCd = clsData.CtrlCd
        newItem.LevelCd = clsData.LevelCd
        newItem.PrtFg = clsData.PrtFg
        newItem.PrtId = clsData.PrtId
        newItem.DoneFg = clsData.DoneFg
        newItem.SectCd = clsData.SectCd
    End With
    
    mSchedCol.Add newItem
    Set AddSchedule = newItem
    mvarSchedCnt = mvarSchedCnt + 1
   
End Function

Public Sub SchedClear()
    Set mSchedCol = Nothing
    Set mSchedCol = New Collection
    mvarSchedCnt = 0
End Sub

Public Sub GetControlCd(ByVal pCtrlCd As String, Optional ByVal pLevelCd As String = "A")

    Dim objSqlQc As clsLISSqlQc
    Dim SqlStmt As String
    Dim Rs As Recordset
    
    Set objSqlQc = New clsLISSqlQc
    SqlStmt = objSqlQc.SqlControlCd(pCtrlCd, pLevelCd)
    Set objSqlQc = Nothing
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, dbconn
    
    If Rs.EOF Then
        NewFg = True
        Changed = True
        GoTo Nodata
    End If
    
    NewFg = False
    Changed = False
    
    mvarCtrlCd = pCtrlCd
    mvarLevelCd = "" & Rs.Fields("LevelCd").Value
    mvarCtrlNm = "" & Rs.Fields("CtrlNm").Value
    mvarEqpCd = "" & Rs.Fields("EqpCd").Value
    mvarEqpNm = "" & Rs.Fields("EqpNm").Value
    mvarSectCd = "" & Rs.Fields("SectCd").Value
    mvarSectNm = "" & Rs.Fields("SectNm").Value
    mvarCtrlDiv = "" & Rs.Fields("CtrlDiv").Value
    mvarWorkArea = "" & Rs.Fields("WorkArea").Value
    mvarBuildCd = "" & Rs.Fields("BuildCd").Value
    
Nodata:
    Set Rs = Nothing
    Set objSqlQc = Nothing
End Sub

Public Sub GetQCItems(ByVal pCtrlCd As String, ByVal pLevelCd As String)

    Dim objSqlQc As clsLISSqlQc
    Dim rs1 As Recordset
    Dim SqlStmt As String
    Dim newItem As clsControlItem
    
    Set objSqlQc = New clsLISSqlQc
    SqlStmt = objSqlQc.SqlQCItems(pCtrlCd, pLevelCd)
    Set objSqlQc = Nothing
    
    Set rs1 = New Recordset
    rs1.Open SqlStmt, dbconn
    
    Set mItemCol = New Collection
'    mvarCount = 0
    rs1.MoveFirst
    While Not rs1.EOF
        Set newItem = New clsControlItem
        newItem.CtrlCd = "" & rs1.Fields("CtrlCd").Value
        newItem.LevelCd = "" & rs1.Fields("LevelCd").Value
        newItem.TestCd = "" & rs1.Fields("TestCd").Value
        newItem.TestNm = "" & rs1.Fields("TestNm").Value
        newItem.EqpCd = "" & rs1.Fields("EqpCd").Value
        newItem.AvalVal = "" & rs1.Fields("AvalVal").Value
        newItem.RstUnit = "" & rs1.Fields("RstUnit").Value
        newItem.PanelFg = "" & rs1.Fields("PanelFg").Value
        
        mItemCol.Add newItem, "" & rs1.Fields("TestCd").Value
        mvarItemCount = mvarItemCount + 1
        
        If newItem.PanelFg = PN_Detail Then
            Call GetDetails(newItem.CtrlCd, newItem.LevelCd, newItem.TestCd, newItem.EqpCd)
        End If
        
        rs1.MoveNext
    Wend
    
    Set rs1 = Nothing
    
End Sub


Public Sub GetDetails(ByVal pControlCd As String, ByVal pLevelCd As String, _
                      ByVal pTestCd As String, ByVal pEqpCd As String)
     Dim objSqlQc As clsLISSqlQc
     Dim SqlStmt As String
     Dim rs2 As Recordset
     Dim newItem As clsControlItem
     Dim strTestcd As String
        
    Set objSqlQc = New clsLISSqlQc
    SqlStmt = objSqlQc.SqlDetails(pTestCd)
    Set objSqlQc = Nothing
    
    Set rs2 = New Recordset
    rs2.Open SqlStmt, dbconn
    
    On Error GoTo ErrTrab
    rs2.MoveFirst
    While Not rs2.EOF
        Set newItem = New clsControlItem
        newItem.CtrlCd = pControlCd
        newItem.LevelCd = pLevelCd
        newItem.TestCd = "" & rs2.Fields("TestCd").Value
'        Debug.Print "" & rs2.Fields("TestCd").Value
        newItem.TestNm = "" & rs2.Fields("TestNm").Value
        newItem.EqpCd = pEqpCd
        newItem.AvalVal = "" & rs2.Fields("AvalVal").Value
        newItem.RstUnit = "" & rs2.Fields("RstUnit").Value
        newItem.PanelFg = "" & rs2.Fields("PanelFg").Value
        
       mItemCol.Add newItem, "" & rs2.Fields("TestCd").Value
       mvarItemCount = mvarItemCount + 1
       rs2.MoveNext
    Wend
    
    Set rs2 = Nothing
    Exit Sub
    
ErrTrab:
    MsgBox Err.Description, vbExclamation
    
End Sub

Private Sub DataMove(ByRef objTLab021 As clsTLab021)
    With objTLab021
        .CtrlCd = mvarCtrlCd
        .LevelCd = mvarLevelCd
        .CtrlNm = mvarCtrlNm
        .EqpCd = mvarEqpCd
        .SectCd = mvarSectCd
        .CtrlDiv = mvarCtrlDiv
        .WorkArea = mvarWorkArea
        .BuildCd = mvarBuildCd
    End With
End Sub

Public Sub SaveData()

    Dim objSqlQc As clsLISSqlQc
    Dim objTLab021 As clsTLab021
    Dim SqlStmtSav As String
    Dim SqlStmtDel As String
    Dim i As Integer
    
    Set objSqlQc = New clsLISSqlQc
    Set objTLab021 = New clsTLab021
    Call DataMove(objTLab021)
    SqlStmtSav = objSqlQc.SqlSaveData(mvarNewFg, objTLab021)
    SqlStmtDel = objSqlQc.SqlDeleteData(mvarCtrlCd, mvarLevelCd)
    Set objTLab021 = Nothing
    Set objSqlQc = Nothing
    
    dbconn.BeginTrans
    dbconn.Execute SqlStmtSav
    'If mvarChanged Then
        dbconn.Execute SqlStmtDel
        For i = 1 To mvarItemCount
'            Set Item(i).DBConn = DBConn
            Call Item(i).ItemSave
        Next
    'End If
    dbconn.CommitTrans
    
End Sub

Public Sub DeleteData()

    Dim objSqlQc As clsLISSqlQc
    Dim objTLab021 As clsTLab021
    Dim SqlStmtSav As String
    Dim SqlStmtDel As String
    Dim i As Integer
    
    Set objSqlQc = New clsLISSqlQc
    Set objTLab021 = New clsTLab021
    Call DataMove(objTLab021)
    SqlStmtSav = objSqlQc.SqlSaveData(mvarNewFg, objTLab021)
    SqlStmtDel = objSqlQc.SqlDeleteData(mvarCtrlCd, mvarLevelCd)
    Set objTLab021 = Nothing
    Set objSqlQc = Nothing
    
    dbconn.BeginTrans
    dbconn.Execute SqlStmtSav
    'If mvarChanged Then
        dbconn.Execute SqlStmtDel
        For i = 1 To mvarItemCount
'            Set Item(i).DBConn = DBConn
            Call Item(i).ItemSave
        Next
    'End If
    dbconn.CommitTrans
    
End Sub

Public Sub DeleteControlAll(ByVal pCtrlCd As String, ByVal pLevelCd As String)

    Dim objSqlQc As clsLISSqlQc
    Dim SqlStmtDel As String
    Dim i As Integer
    
    Set objSqlQc = New clsLISSqlQc
    
    dbconn.BeginTrans
    
    For i = 1 To 4
        SqlStmtDel = objSqlQc.SqlDeleteAllData(mvarCtrlCd, mvarLevelCd, i)
        dbconn.Execute SqlStmtDel
    Next
    
    dbconn.CommitTrans
    Set objSqlQc = Nothing
    
End Sub



Public Sub CreateSchedule(ByVal pFromDt As String, ByVal pToDt As String, _
                          ByVal pCtrlCd As String, Optional ByVal pLevelCd As String = "A")

    Dim objSqlQc As clsLISSqlQc
    Dim SqlStmt As String
    Dim i As Long
    
    Set objSqlQc = New clsLISSqlQc
    SqlStmt = objSqlQc.SqlCreateSchedule(pFromDt, pToDt, pCtrlCd, pLevelCd)
    Set objSqlQc = Nothing
    
    dbconn.BeginTrans
    dbconn.Execute SqlStmt
    For i = 1 To mvarSchedCnt
'        Set Schedule(i).DBConn = DBConn
        Call Schedule(i).ScheduleSave
    Next
    dbconn.CommitTrans
    
End Sub

Public Sub GetSchedule(ByVal pCtrlCd As String, ByVal pFromDt As String, _
                       ByVal pToDt As String, Optional ByVal pLevelCd As String = "A")
    
    Dim objSqlQc As clsLISSqlQc
    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim newItem As clsSchedule
    
    Set objSqlQc = New clsLISSqlQc
    SqlStmt = objSqlQc.SqlSchedule(pCtrlCd, pFromDt, pToDt, pLevelCd)
    Set objSqlQc = Nothing
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, dbconn
    
    Set mSchedCol = New Collection
    mvarSchedCnt = 0
    While Not Rs.EOF
        Set newItem = New clsSchedule
        newItem.DoDt = "" & Rs.Fields("DoDt").Value
        newItem.DoTm = "" & Rs.Fields("DoTm").Value
        newItem.CtrlCd = "" & Rs.Fields("CtrlCd").Value
        newItem.LevelCd = "" & Rs.Fields("LevelCd").Value
        newItem.CtrlNm = "" & Rs.Fields("CtrlNm").Value
        newItem.PrtFg = "" & Rs.Fields("PrtFg").Value
        newItem.PrtId = "" & Rs.Fields("PrtId").Value
        newItem.DoneFg = "" & Rs.Fields("DoneFg").Value
        
        mSchedCol.Add newItem
        mvarSchedCnt = mvarSchedCnt + 1
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
End Sub


Public Sub LoadControl(ByRef lstList As Object)

    Dim Rs As Recordset
    Dim objCtrl As clsLISSqlQc
    
    Set objCtrl = New clsLISSqlQc
    
    With objCtrl
        Set Rs = New Recordset
        Rs.Open .GetCtrlMst, dbconn
    End With
    
    If Rs.EOF Then GoTo Nodata
    
    lstList.Clear
    While (Not Rs.EOF)
        lstList.addItem Format("" & Rs.Fields("CtrlCd").Value, "!" & String(12, "@")) & COL_DIV & _
                        Format("" & Rs.Fields("LevelCd").Value, "!" & String(3, "@")) & COL_DIV & _
                        Format("" & Rs.Fields("CtrlNm").Value, "!" & String(30, "@"))
        Rs.MoveNext
    Wend
    
Nodata:
    Set Rs = Nothing
    Set objCtrl = Nothing
End Sub


Public Sub LoadCtrlCd(ByRef lstList As Object)
'    Dim Rs As Recordset
'    Dim objCtrlCd As clsLISSqlQc
'
'    Set objCtrlCd = New clsLISSqlQc
'
'    Set Rs = OpenRecordSet(objCtrlCd.GetCtrlMst)
'
'    Do Until Rs.EOF
'        lstList.addItem
'
'        Rs.MoveNext
'    Loop
'
'    Set Rs = Nothing
'    Set objCtrlCd = Nothing
End Sub

Public Sub LoadEqpCd(ByRef lstList As Object)

    Dim Rs As Recordset
    Dim objEqp As clsLISSqlQc
    Dim i%
    
    
    Set objEqp = New clsLISSqlQc
    
    With objEqp
        Set Rs = New Recordset
        Rs.Open .GetEqpMst, dbconn
    End With
    
    If Rs.EOF Then GoTo Nodata
    
    lstList.Clear
    For i = 1 To Rs.RecordCount
        lstList.addItem "" & Rs.Fields("eqpcd").Value & vbTab & _
                        "" & Rs.Fields("eqpnm").Value
        Rs.MoveNext
    Next i
    
Nodata:
    Set Rs = Nothing
    Set objEqp = Nothing
End Sub


Public Sub LoadSection(ByRef cboCombo As Object)

    Dim Rs As Recordset
    Dim objSec As clsLISSqlQc
    Dim i%
        
    Set objSec = New clsLISSqlQc
        
    With objSec
        Set Rs = New Recordset
        Rs.Open .GetSection, dbconn
    End With
    
    If Rs.EOF Then GoTo Nodata
    
    cboCombo.Clear
    For i = 1 To Rs.RecordCount
        cboCombo.addItem "" & Rs.Fields("SectCd").Value & "   " & _
                         "" & Rs.Fields("SectNm").Value
        Rs.MoveNext
    Next i
    
Nodata:
    Set Rs = Nothing
    Set objSec = Nothing
End Sub

Public Sub LoadPrinter(ByRef cboCombo As Object)

    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim i As Long
    Dim tmpStr As String

    SqlStmt = "select * from " & T_LAB032 & " " & _
              "where cdindex = " & DBStr(LC3_PrinterId)
    Set Rs = New Recordset
    Rs.Open SqlStmt, dbconn
    
    If Rs.EOF Then GoTo Nodata
    
    cboCombo.Clear
    For i = 1 To Rs.RecordCount
        tmpStr = "" & Rs.Fields("CdVal1").Value & Space(9)
        cboCombo.addItem Mid(tmpStr, 1, 6) & _
                        "" & Rs.Fields("Field1").Value
        Rs.MoveNext
    Next i
    
Nodata:
    Set Rs = Nothing
    
End Sub



Public Sub LoadWorkArea(ByRef cboList As Object)

    Dim i%
    Dim tmpRs As Recordset
    Dim objWA As clsLISSqlQc
    
    Set objWA = New clsLISSqlQc
    
    With objWA
        Set tmpRs = New Recordset
        tmpRs.Open .GetWorkArea, dbconn
    End With

    If tmpRs.EOF = True Then ' record가 존재하지 않을 경우
        Set tmpRs = Nothing
        Set objWA = Nothing
        Exit Sub
    End If

    cboList.Clear
    With tmpRs
        .MoveFirst
        For i = 1 To .RecordCount
            cboList.addItem Format("" & .Fields("cdval1").Value, "!@@@@") & "" & .Fields("field1").Value
            .MoveNext
        Next i
    End With
    cboList.ListIndex = -1

    Set tmpRs = Nothing
    Set objWA = Nothing

End Sub

Public Sub LoadBuilding(ByRef cboList As Object)
    
    Dim RsBld As New Recordset
    Dim sBldCd As String, sBldNm As String, sBldNo As Integer
    Dim objBld As clsLISSqlQc
    
    Set objBld = New clsLISSqlQc
    
    With objBld
        Set RsBld = New Recordset
        RsBld.Open .GetBuilding, dbconn
    End With
    
    If RsBld.EOF Then GoTo Nodata
    
    cboList.Clear
    
    Do While Not RsBld.EOF
        sBldCd = "" & RsBld.Fields("cdval1").Value
        sBldNm = "" & RsBld.Fields("field1").Value
        sBldNo = Val("" & RsBld.Fields("field2").Value)
        cboList.addItem Format(sBldCd, "!@@@@") & sBldNm
        RsBld.MoveNext
    Loop

    If cboList.ListCount > 0 Then cboList.ListIndex = objsysinfo.BuildingNo - 1

Nodata:
    Set RsBld = Nothing
    Set objBld = Nothing
End Sub





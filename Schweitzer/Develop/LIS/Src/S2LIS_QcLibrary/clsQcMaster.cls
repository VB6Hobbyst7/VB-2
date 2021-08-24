VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQcMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarCtrlCd As String    '-- Control코드
Private mvarLevelCd As String   '-- Level
Private mvarLotNo As String     '-- Lot Number

Private mvarOpenDt As String    '-- 시작일
Private mvarExpDt As String     '-- 폐기일
Private mvarMakeCd As String    '-- 제조사
Private mvarRemark As String    '-- Remark
   
Private mvarNewFg As Boolean    '-- 신규여부
'Private mvarMyDb As DrDatabase

Private mItemCol As Collection
Private mvarItemCount As Integer '--컬렉션 카운트


'Database Connection
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

'Lot Number
Public Property Get Lotno() As String
    Lotno = mvarLotNo
End Property
Public Property Let Lotno(ByVal vNewValue As String)
    mvarLotNo = vNewValue
End Property

'시작일
Public Property Get OpenDt() As String
    OpenDt = mvarOpenDt
End Property
Public Property Let OpenDt(ByVal vNewValue As String)
    mvarOpenDt = vNewValue
End Property

'폐기일
Public Property Get ExpDt() As String
    ExpDt = mvarExpDt
End Property
Public Property Let ExpDt(ByVal vNewValue As String)
    mvarExpDt = vNewValue
End Property

'제조사
Public Property Get MakeCd() As String
    MakeCd = mvarMakeCd
End Property
Public Property Let MakeCd(ByVal vNewValue As String)
    mvarMakeCd = vNewValue
End Property

'Remark
Public Property Get Remark() As String
    Remark = mvarRemark
End Property
Public Property Let Remark(ByVal vNewValue As String)
    mvarRemark = vNewValue
End Property

'신규여부
Public Property Get NewFg() As Boolean
    NewFg = mvarNewFg
End Property
Public Property Let NewFg(ByVal vNewValue As Boolean)
    mvarNewFg = vNewValue
End Property


'검사항목 컬렉션
Public Property Get ItemCount() As Integer
    ItemCount = mItemCol.Count
End Property

Public Property Get Item(ByVal KeyIndex As Variant) As clsQcItem
    Set Item = mItemCol(KeyIndex)
End Property

Private Sub Class_Initialize()
    Set mItemCol = New Collection
    mvarItemCount = 0
End Sub

Public Sub ItemClear()
    Set mItemCol = Nothing
    Set mItemCol = New Collection
    mvarItemCount = 0
End Sub

Public Function AddQcItem(ByVal clsKey As Variant, ByVal clsData As clsQcItem)
   
    Dim newItem As New clsQcItem
    
    With newItem
        newItem.CtrlCd = clsData.CtrlCd
        newItem.LevelCd = clsData.LevelCd
        newItem.Lotno = clsData.Lotno
        newItem.TestCd = clsData.TestCd
        newItem.TestNm = clsData.TestNm
        newItem.AbbrNm = clsData.AbbrNm
        newItem.MeanVal = clsData.MeanVal
        newItem.SdVal = clsData.SdVal
        newItem.AvalVal = clsData.AvalVal
        newItem.RefCd = clsData.RefCd
        newItem.RstUnit = clsData.RstUnit
        newItem.CvVal = clsData.CvVal
        newItem.MinVal = clsData.MinVal
        newItem.MaxVal = clsData.MaxVal
        newItem.SchedFg = clsData.SchedFg
        newItem.WmSet = clsData.WmSet
        newItem.Selected = clsData.Selected
    End With
    
    mItemCol.Add newItem
    Set AddQcItem = newItem
    mvarItemCount = mvarItemCount + 1
   
End Function


Public Sub GetLotnoList(ByRef lstList As Object, ByVal pCtrlCd As String, ByVal pLevelCd As String)
    
    Dim objSqlQc As clsLISSqlQc
    Dim SqlStmt As String
    Dim Rs As Recordset
    
    Set objSqlQc = New clsLISSqlQc
    SqlStmt = objSqlQc.SqlMstLotnoList(pCtrlCd, pLevelCd)
    Set objSqlQc = Nothing
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, dbconn
    
    If Rs.EOF Then GoTo Nodata
    
    lstList.Clear
    While (Not Rs.EOF)
        lstList.addItem Rs.Fields("lotno").Value & "" & vbTab & _
                        Rs.Fields("opendt").Value & "" & vbTab & _
                        Rs.Fields("expdt").Value & ""
        Rs.MoveNext
    Wend
    
Nodata:
    Set Rs = Nothing
    Set objSqlQc = Nothing
End Sub

Public Sub GetQcData(ByVal pCtrlCd As String, ByVal pLevelCd As String, ByVal pLotNo As String)

    Dim objSqlQc As clsLISSqlQc
    Dim SqlStmt As String
    Dim Rs As Recordset
    
    Set objSqlQc = New clsLISSqlQc
    SqlStmt = objSqlQc.SqlMstQcData(pCtrlCd, pLevelCd, pLotNo)
    Set objSqlQc = Nothing
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, dbconn
    
    If Rs.EOF Then
        NewFg = True
'        Changed = True
        GoTo Nodata
    End If
    
    NewFg = False
'    Changed = False
    
    mvarCtrlCd = pCtrlCd
    mvarLevelCd = Rs.Fields("LevelCd").Value & ""
'    mvarCtrlNm = Rs.Fields("CtrlNm").Value & ""
    mvarLotNo = Rs.Fields("LotNo").Value & ""
    mvarOpenDt = Rs.Fields("OpenDt").Value & ""
    mvarExpDt = Rs.Fields("ExpDt").Value & ""
    mvarMakeCd = Rs.Fields("MakeCd").Value & ""
    mvarRemark = Rs.Fields("Remark").Value & ""
    
Nodata:
    Set Rs = Nothing
    
End Sub

Public Sub GetQCItems(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                      ByVal pLotNo As String, Optional ByVal pDiv As String = "D")

    Dim objSqlQc As clsLISSqlQc
    Dim i As Integer
    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim newItem As clsQcItem
    
    Set objSqlQc = New clsLISSqlQc
    SqlStmt = objSqlQc.SqlMstQCItems(pCtrlCd, pLevelCd, pLotNo, pDiv)
    Set objSqlQc = Nothing
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, dbconn
    
    Set mItemCol = New Collection
'    mvarCount = 0
    While Not Rs.EOF
        Set newItem = New clsQcItem
        newItem.CtrlCd = Rs.Fields("CtrlCd").Value & ""
        newItem.LevelCd = Rs.Fields("LevelCd").Value & ""
        newItem.Lotno = Rs.Fields("LotNo").Value & ""
        newItem.TestCd = Rs.Fields("TestCd").Value & ""
        newItem.AbbrNm = Rs.Fields("AbbrNm5").Value & ""
        If Trim("" & Rs.Fields("DetailFg").Value) <> "" Then
            newItem.TestNm = "   " & Rs.Fields("TestNm").Value
        Else
            newItem.TestNm = Rs.Fields("TestNm").Value & ""
        End If
        newItem.MeanVal = Val(Rs.Fields("MeanVal").Value)
        newItem.SdVal = Val(Rs.Fields("SdVal").Value)
        newItem.AvalVal = Val(Rs.Fields("AvalVal").Value)
        newItem.RefCd = Rs.Fields("RefCd").Value & ""
        newItem.RstUnit = Rs.Fields("RstUnit").Value & ""
        newItem.CvVal = Val(Rs.Fields("CvVal").Value)
        newItem.MinVal = Val(Rs.Fields("MinVal").Value)
        newItem.MaxVal = Val(Rs.Fields("MaxVal").Value)
        newItem.SchedFg = Rs.Fields("SchedFg").Value & ""
        newItem.WmSet = Rs.Fields("WmSet").Value & ""
        
        newItem.RstType = Rs.Fields("RstType").Value & ""
        newItem.RstDiv = Rs.Fields("RstDiv").Value & ""
        newItem.TestDiv = Rs.Fields("TestDiv").Value & ""
        newItem.PanelFg = Rs.Fields("PanelFg").Value & ""
        
        For i = 1 To 5
            newItem.WmFg(i) = Val(Mid(Rs.Fields("WmSet").Value & "", i, 1))
        Next
        newItem.CalFg = Rs.Fields("CalFg").Value & ""
        newItem.EntDt = Rs.Fields("EntDt").Value & ""
        newItem.EntId = Rs.Fields("EntId").Value & ""
        newItem.Selected = False
        
        mItemCol.Add newItem, Rs.Fields("TestCd").Value & ""
        mvarItemCount = mvarItemCount + 1
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
End Sub

Public Sub GetQCOrders(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                       ByVal pLotNo As String, ByVal pWorkArea As String, _
                       ByVal pAccDt As String, ByVal pAccSeq As String)

    Dim objSqlQc As clsLISSqlQc
    Dim i As Integer
    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim newItem As clsQcItem
    
    Set objSqlQc = New clsLISSqlQc
    SqlStmt = objSqlQc.SqlOrderQCItems(pCtrlCd, pLevelCd, pLotNo, pWorkArea, pAccDt, pAccSeq)
    Set objSqlQc = Nothing
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, dbconn
    
    Set mItemCol = New Collection
'    mvarCount = 0
    While Not Rs.EOF
        Set newItem = New clsQcItem
        newItem.CtrlCd = Rs.Fields("CtrlCd").Value & ""
        newItem.LevelCd = Rs.Fields("LevelCd").Value & ""
        newItem.Lotno = Rs.Fields("LotNo").Value & ""
        newItem.TestCd = Rs.Fields("TestCd").Value & ""
        newItem.AbbrNm = Rs.Fields("AbbrNm5").Value & ""
        If Trim("" & Rs.Fields("DetailFg").Value) <> "" Then
            newItem.TestNm = "   " & Rs.Fields("TestNm").Value & ""
        Else
            newItem.TestNm = Rs.Fields("TestNm").Value & ""
        End If
        newItem.MeanVal = Val(Rs.Fields("MeanVal").Value)
        newItem.SdVal = Val(Rs.Fields("SdVal").Value)
        newItem.AvalVal = Val(Rs.Fields("AvalVal").Value)
        newItem.RefCd = Rs.Fields("RefCd").Value & ""
        newItem.RstUnit = Rs.Fields("RstUnit").Value & ""
        newItem.CvVal = Val(Rs.Fields("CvVal").Value)
        newItem.MinVal = Val(Rs.Fields("MinVal").Value)
        newItem.MaxVal = Val(Rs.Fields("MaxVal").Value)
        newItem.SchedFg = Rs.Fields("SchedFg").Value & ""
        newItem.WmSet = Rs.Fields("WmSet").Value & ""
        
        newItem.RstType = Rs.Fields("RstType").Value & ""
        newItem.RstDiv = Rs.Fields("RstDiv").Value & ""
        newItem.TestDiv = Rs.Fields("TestDiv").Value & ""
        newItem.PanelFg = Rs.Fields("PanelFg").Value & ""
        
        For i = 1 To 5
            newItem.WmFg(i) = Val(Mid(Rs.Fields("WmSet").Value & "", i, 1))
        Next
        newItem.CalFg = Rs.Fields("CalFg").Value & ""
        newItem.EntDt = Rs.Fields("EntDt").Value & ""
        newItem.EntId = Rs.Fields("EntId").Value & ""
        If Trim(Rs.Fields("OrdFg").Value & "") = "" Then
            newItem.Selected = False
        Else
            newItem.Selected = True
        End If
        
        mItemCol.Add newItem, Rs.Fields("TestCd").Value & ""
        mvarItemCount = mvarItemCount + 1
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
End Sub

Private Sub DataMove(ByRef objTLab023 As clsTLab023)
    With objTLab023
        .CtrlCd = mvarCtrlCd
        .LevelCd = mvarLevelCd
        .Lotno = mvarLotNo
        .OpenDt = mvarOpenDt
        .ExpDt = mvarExpDt
        .MakeCd = mvarMakeCd
        .Remark = mvarRemark
    End With
End Sub

Public Sub SaveData()

    Dim objSqlQc As clsLISSqlQc
    Dim objTLab023 As clsTLab023
    Dim SqlStmtSav As String
    Dim SqlStmtDel As String
    Dim i As Integer
    
    Set objSqlQc = New clsLISSqlQc
    Set objTLab023 = New clsTLab023
    Call DataMove(objTLab023)
    SqlStmtSav = objSqlQc.SqlMstSaveData(mvarNewFg, objTLab023)
    SqlStmtDel = objSqlQc.SqlMstDeleteData(mvarCtrlCd, mvarLevelCd, mvarLotNo)
    Set objTLab023 = Nothing
    Set objSqlQc = Nothing
    
    
On Error GoTo Err_Trap

    dbconn.BeginTrans
    dbconn.Execute SqlStmtSav
    dbconn.Execute SqlStmtDel
    
    Dim blnItemSave As Boolean
    
    For i = 1 To mvarItemCount
        'Set Item(i).DBConn = DBConn
        blnItemSave = Item(i).ItemSave
        If Not blnItemSave Then GoTo Err_Trap
    Next
    dbconn.CommitTrans
    Exit Sub
    
Err_Trap:
    MsgBox Err.Description, vbExclamation

End Sub


Public Sub DeleteData()

    Dim objSqlQc As clsLISSqlQc
    Dim SqlStmtDel1 As String
    Dim SqlStmtDel2 As String
    Dim i As Integer
    
    Set objSqlQc = New clsLISSqlQc
    SqlStmtDel1 = objSqlQc.SqlMstDeleteAll(mvarCtrlCd, mvarLevelCd, mvarLotNo, "1")
    SqlStmtDel2 = objSqlQc.SqlMstDeleteAll(mvarCtrlCd, mvarLevelCd, mvarLotNo, "2")
    Set objSqlQc = Nothing
    
    
On Error GoTo Err_Trap

    dbconn.BeginTrans
    dbconn.Execute SqlStmtDel1
    dbconn.Execute SqlStmtDel2
    dbconn.CommitTrans
    Exit Sub
    
Err_Trap:
    MsgBox Err.Description, vbExclamation
End Sub

Public Function GetQcTestcd(ByVal Controlcd As String, ByVal LevelCd As String) As String

    GetQcTestcd = "select a.testcd,b.testnm from " & T_LAB022 & " a," & T_LAB001 & " b" & _
                  " where " & DBW("a.ctrlcd=", Controlcd) & " and " & DBW("a.levelcd=", LevelCd) & _
                  " and a.testcd=b.testcd"
End Function

Public Function GetQcResult(ByVal Controlcd As String, ByVal LevelCd As String, _
                            ByVal Lotno As String, ByVal TestCd As String, _
                            ByVal FDate As String, ByVal TDate As String) As String

'2003/11/07 remark by legends
'유효숫자도 갖고 오기위해서 테이블 조인을 해야되므로 부득이하게 수정해야 했음

'    GetQcResult = " select workarea,accdt,accseq,rstcd,radiv,vfydt from " & T_LAB026 & _
'                  " where " & DBW("ctrlcd=", Controlcd) & _
'                  " and " & DBW("levelcd=", LevelCd) & _
'                  " and " & DBW("lotno=", Lotno) & _
'                  " and " & DBW("testcd=", TestCd) & _
'                  " and " & DBW("vfydt>=", FDate) & _
'                  " and " & DBW("vfydt<=", TDate)
    
    GetQcResult = " select a.workarea,a.accdt,a.accseq,a.rstcd,a.radiv,a.vfydt,b.meanval,b.sdval,b.cvval,b.minval,b.maxval,b.avalval " & _
                  " from " & T_LAB026 & " a, " & T_LAB024 & " b " & _
                  " where " & DBW("a.ctrlcd=", Controlcd) & _
                  " and " & DBW("a.levelcd=", LevelCd) & _
                  " and " & DBW("a.lotno=", Lotno) & _
                  " and " & DBW("a.testcd=", TestCd) & _
                  " and " & DBW("a.vfydt>=", FDate) & _
                  " and " & DBW("a.vfydt<=", TDate) & _
                  " and a.ctrlcd=b.ctrlcd " & _
                  " and a.levelcd=b.levelcd " & _
                  " and a.lotno=b.lotno " & _
                  " and a.testcd=b.testcd "
End Function

Public Function SetLab026(ByVal CtrlCd As String, ByVal LevelCd As String, ByVal Lotno As String, ByVal TestCd As String, _
                         ByVal Mean As String, ByVal SD As String, ByVal Cv As String, ByVal sMax As String, ByVal sMin As String) As Boolean

    Dim Rs     As Recordset
    Dim sSql   As String
    Dim strTmp As String
    Dim ErrMsg As String
    
    On Error GoTo SaveERROR
    dbconn.BeginTrans
    sSql = " select * from " & T_LAB024 & _
           " where " & _
                     DBW("ctrlcd=", CtrlCd) & " and " & _
                     DBW("levelcd=", LevelCd) & " and " & _
                     DBW("lotno=", Lotno) & " and " & _
                     DBW("testcd=", TestCd)
    Set Rs = New Recordset
    Rs.Open sSql, dbconn
    
    If Not Rs.EOF Then
        strTmp = MsgBox("기존의 자료는 삭제됩니다." & vbCrLf & "저장하시겠습니까?", vbInformation + vbYesNo, "저장")
        
        If strTmp = vbNo Then GoTo Skip
        
        sSql = " delete   from " & T_LAB024 & _
               " where " & _
                          DBW("ctrlcd=", CtrlCd) & " and " & _
                          DBW("levelcd=", LevelCd) & " and " & _
                          DBW("lotno=", Lotno) & " and " & _
                          DBW("testcd=", TestCd)
    
        dbconn.Execute sSql
        
        sSql = "insert into " & T_LAB024 & "(ctrlcd,levelcd,lotno,testcd,meanval,sdval,avalval,refcd,rstunit," & _
                                        "cvval,minval,maxval,schedfg,wmset,calfg,entdt,entid) values (" & _
                                        DBV("ctrlcd", CtrlCd, 1) & _
                                        DBV("levelcd", LevelCd, 1) & _
                                        DBV("lotno", Lotno, 1) & _
                                        DBV("testcd", TestCd, 1) & _
                                        DBV("meanval", Mean, 1) & _
                                        DBV("sdval", SD, 1) & _
                                        DBV("avalval", Rs.Fields("avalval").Value & "", 1) & _
                                        DBV("refcd", Rs.Fields("refcd").Value & "", 1) & _
                                        DBV("rstunit", Rs.Fields("rstunit").Value & "", 1) & _
                                        DBV("cvval", Cv, 1) & _
                                        DBV("minval", sMin, 1) & _
                                        DBV("maxval", sMax, 1) & _
                                        DBV("schedfg", Rs.Fields("schedfg").Value & "", 1) & _
                                        DBV("wmset", Rs.Fields("wmset").Value & "", 1) & _
                                        DBV("calfg", "1", 1) & _
                                        DBV("entdt", Format(Getsystemdate, "yyyymmdd"), 1) & _
                                        DBV("entid", objsysinfo.EmpId) & ")"
             
        dbconn.Execute sSql
    Else
            
        sSql = " insert into " & T_LAB024 & "(ctrlcd,levelcd,lotno,testcd,meanval,sdval,cvval,minval,maxval,calfg," & _
                                            "entdt,entid) values (" & _
                                            DBV("ctrlcd", CtrlCd, 1) & _
                                            DBV("levelcd", LevelCd, 1) & _
                                            DBV("lotno", Lotno, 1) & _
                                            DBV("testcd", TestCd, 1) & _
                                            DBV("meanval", Mean, 1) & _
                                            DBV("sdval", SD, 1) & _
                                            DBV("cvval", Cv, 1) & _
                                            DBV("minval", sMin, 1) & _
                                            DBV("maxval", sMax, 1) & _
                                            DBV("calfg", "1", 1) & _
                                            DBV("entdt", Format(Getsystemdate, "yyyymmdd"), 1) & _
                                            DBV("entid", objsysinfo.EmpId) & ")"
        
        
        dbconn.Execute sSql
    End If
    Set Rs = Nothing
    
    dbconn.CommitTrans
    SetLab026 = True
    Exit Function
SaveERROR:
    
    MsgBox "저장시 오류가 발생하였습니다." & vbCr & Err.Description, vbExclamation
Skip:
    dbconn.RollbackTrans
    Set Rs = Nothing
End Function
                         


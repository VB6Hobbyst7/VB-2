VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHosComCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private mvarWARDID      As clsDictionary            '병상
'Private mvarDEPTCD      As clsDictionary            '부서
'Private mvarPTHDOCT     As clsDictionary            '판독의
'Private mvarSCREENID    As clsDictionary            '스크린ID
'Private mvarRESID       As clsDictionary            '레지던트
'Private mvarDOCT        As clsDictionary            '처방의
'Private mvarEMPID       As clsDictionary            '직원
'Private mvarTCd         As clsDictionary            '조직코드
'Private mvarMCd         As clsDictionary            '진단코드
'Private mvarOCD         As clsDictionary            '채취방법코드
'Private mvarBuilding    As clsDictionary            '건물정보
'Private mvarTemp        As clsDictionary            'Com004 결과템플릿
'Private mvarBarInfo     As clsBarcode
'Private mvarLISITEM     As clsDictionary            '검사항목
'Private mvarLISSPC      As clsDictionary            '검체정보
'Private objSQL          As New clsHosComSQLStmt
'Private objDiv          As New clsHosComConstants
'Private objTables       As New clsTables
'Private objFields       As New clsFields
'Private NewForm         As Form
'Private mvarProjectCd   As String
'
'Public Sub SetForm(ByRef pForm As Object)
'    Set NewForm = pForm
'End Sub
'
'Public Property Let ProjectCd(ByVal vData As String)
'    mvarProjectCd = vData
'End Property
'Public Property Get ProjectCd() As String
'    ProjectCd = mvarProjectCd
'End Property

'Public Function LoadAPSEntity() As Boolean
'    Dim objProgress As clsProgressBar
'    Dim i As Long
'
'    Set objProgress = New clsProgressBar
'    Set objProgress.StatusBar = NewForm.stsbar
'    NewForm.MousePointer = 13
'    objProgress.Msg = "자료를 읽고 있습니다."
'
'
''    Call LoadEmp
'    DoEvents
'    For i = 1 To 10
'        objProgress.Value = i
'    Next
'
'    Call LoadDept
'    DoEvents
'    For i = 11 To 20
'        objProgress.Value = i
'    Next
'
'    Call LoadWard
'    DoEvents
'    For i = 21 To 30
'        objProgress.Value = i
'    Next
'
'    Call LoadDoct
'    DoEvents
'    For i = 31 To 40
'        objProgress.Value = i
'    Next
'
'    Call LoadMSnomed
'    DoEvents
'
'    For i = 41 To 50
'        objProgress.Value = i
'    Next
'
'    Call LoadTSnomed
'    DoEvents
'    For i = 51 To 60
'        objProgress.Value = i
'    Next
'
'    Call LoadOrgans
'    DoEvents
'    For i = 61 To 70
'        objProgress.Value = i
'    Next
'
'    Call LoadResDoct
'    DoEvents
'    For i = 71 To 80
'        objProgress.Value = i
'    Next
'
'    Call LoadScreen
'    DoEvents
'    For i = 81 To 90
'        objProgress.Value = i
'    Next
'
'    Call LoadPthDoct
'    DoEvents
'    For i = 91 To 95
'        objProgress.Value = i
'    Next
'
'    Call LoadBuilding
'    DoEvents
'    For i = 95 To 100
'        objProgress.Value = i
'    Next
'
'    Set objProgress = Nothing
'    NewForm.MousePointer = 0
'
'    LoadAPSEntity = True
'
'End Function

'Public Function LoadLISEntity() As Boolean
'    Dim objProgress As clsProgressBar
'    Dim i           As Long
'
'    Set objProgress = New clsProgressBar
'    Set objProgress.StatusBar = NewForm.stsbar
'    NewForm.MousePointer = 13
'    objProgress.Msg = "자료를 읽고 있습니다."
'
'
''    Call LoadEmp
'    DoEvents
'    For i = 1 To 10
'        objProgress.Value = i
'    Next
'
'    Call LoadDept
'    DoEvents
'    For i = 11 To 30
'        objProgress.Value = i
'    Next
'
'    Call LoadWard
'    DoEvents
'    For i = 31 To 40
'        objProgress.Value = i
'    Next
'
'    Call LoadDoct
'    DoEvents
'    For i = 41 To 60
'        objProgress.Value = i
'    Next
'
'    Call LoadBuilding
'    DoEvents
'    For i = 61 To 70
'        objProgress.Value = i
'    Next
'
'    Call LoadLisSpc
'    DoEvents
'    For i = 61 To 100
'        objProgress.Value = i
'    Next
'
'    DoEvents
'
'    Set objProgress = Nothing
'    NewForm.MousePointer = 0
'
'    LoadLISEntity = True
'
'End Function

'Public Function LoadBuilding()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarBuilding Is Nothing Then
'        Set mvarBuilding = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarBuilding.Clear
'
'    mvarBuilding.FieldInialize "buildcd", "buildnm,buildno"
'
'    mvarBuilding.Sort = False
'
'    '병동ID Load
'    RS.Open objSQL.LoadBuilding, DbConn
'''    If ColCnt > 0 Then
''        Do until rs.eof
'    Do Until RS.EOF
'        mvarBuilding.AddNew RS.Fields("buildcd").Value & "", RS.Fields("buildnm").Value & "" & objDiv.COL_DIV & RS.Fields("buildno").Value & ""
'        RS.MoveNext
'    Loop
''        Rs.movenext
''     loop
''    End If
'
'    '
'    Set RS = Nothing
'
'    mvarBuilding.Sort = True
'End Function

'Public Function LoadWard()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarWARDID Is Nothing Then
'        Set mvarWARDID = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarWARDID.Clear
'
'    mvarWARDID.FieldInialize "wardid", "wardnm"
'
'    mvarWARDID.Sort = False
'
'    '병동ID Load
'
'    RS.Open objSQL.LoadWardId, DbConn
'''    If ColCnt > 0 Then
''        Do until rs.eof
'    Do Until RS.EOF
'        mvarWARDID.AddNew RS.Fields("wardid").Value & "", RS.Fields("wardnm").Value & ""
'        RS.MoveNext
'    Loop
''        Rs.movenext
''     loop
''    End If
'
'    '
'    Set RS = Nothing
'
'    mvarWARDID.Sort = True
'End Function


'Public Function LoadMSnomed()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarMCd Is Nothing Then
'        Set mvarMCd = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarMCd.Clear
'
'    mvarMCd.FieldInialize "mcd", "mnm,mid"
'
'    mvarMCd.Sort = False
'
'    '진단코드 Load
'    RS.Open objSQL.SqlTcdMcdOcd(objTables.T_APS005), DbConn
'    'If rs.RecordCount > 0 Then
''    If ColCnt > 0 Then
'    Do Until RS.EOF
'        mvarMCd.AddNew RS.Fields("mcd").Value & "", RS.Fields("mnm").Value & "" & objDiv.COL_DIV & RS.Fields("mid").Value & ""
'        RS.MoveNext
'     Loop
''    End If
'
'    '
'    Set RS = Nothing
'
'    mvarMCd.Sort = True
'
'End Function

'Public Function LoadTSnomed()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarTCd Is Nothing Then
'        Set mvarTCd = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarTCd.Clear
'
'    mvarTCd.FieldInialize "tcd", "tnm,tid"
'
'    mvarTCd.Sort = False
'
'    '조직코드 Load
'    RS.Open objSQL.SqlTcdMcdOcd(objTables.T_APS004), DbConn
''    If ColCnt > 0 Then
'        Do Until RS.EOF
'            mvarTCd.AddNew RS.Fields("tcd").Value & "", RS.Fields("tnm").Value & "" & objDiv.COL_DIV & RS.Fields("tid").Value & ""
'        RS.MoveNext
'     Loop
''    End If
'
'    '
'    Set RS = Nothing
'    '
'    mvarTCd.Sort = True
'
'End Function

'Public Function LoadOrgans()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarOCD Is Nothing Then
'        Set mvarOCD = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarOCD.Clear
'
'    mvarOCD.FieldInialize "ocd", "onm,oid"
'
'    mvarOCD.Sort = False
'
'    '채취방법 Load
'    RS.Open objSQL.SqlTcdMcdOcd(objTables.T_APS006), DbConn
''    If ColCnt > 0 Then
'        Do Until RS.EOF
'            mvarOCD.AddNew RS.Fields("ocd").Value & "", RS.Fields("onm").Value & "" & objDiv.COL_DIV & RS.Fields("oid").Value & ""
'        RS.MoveNext
'     Loop
''    End If
'    '
'    Set RS = Nothing
'
'    mvarOCD.Sort = True
'
'
'End Function


'Public Function LoadEmp()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarEMPID Is Nothing Then
'        Set mvarEMPID = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarEMPID.Clear
'
'    mvarEMPID.FieldInialize "empid", "empnm"
'
'    mvarEMPID.Sort = False
'
'    '직원정보 Load
'    RS.Open objSQL.LoadEmpId, DbConn
''    If ColCnt > 0 Then
'        Do Until RS.EOF
'            mvarEMPID.AddNew RS.Fields("empid").Value & "", RS.Fields("empnm").Value & ""
'        RS.MoveNext
'     Loop
''    End If
'
'    '
'    Set RS = Nothing
'
'    mvarEMPID.Sort = True
'
'End Function

'Public Function LoadResDoct()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarRESID Is Nothing Then
'        Set mvarRESID = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarRESID.Clear
'
'    mvarRESID.FieldInialize "resid", "resnm"
'
'    mvarRESID.Sort = False
'
'    '레지던트 Load
'    RS.Open objSQL.LoadResId, DbConn
''    If ColCnt > 0 Then
'        Do Until RS.EOF
'            mvarRESID.AddNew RS.Fields("resid").Value & "", RS.Fields("resnm").Value & ""
'        RS.MoveNext
'     Loop
''    End If
'    '
'    Set RS = Nothing
'
'    mvarRESID.Sort = True
'
'End Function

'Public Function LoadScreen()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarSCREENID Is Nothing Then
'        Set mvarSCREENID = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarSCREENID.Clear
'
'    mvarSCREENID.FieldInialize "screenid", "screennm"
'
'    mvarSCREENID.Sort = False
'
'    '스크린 Load
'    RS.Open objSQL.LoadScreenId, DbConn
''    If ColCnt > 0 Then
'        Do Until RS.EOF
'            mvarSCREENID.AddNew RS.Fields("screenid").Value & "", RS.Fields("screennm").Value & ""
'        RS.MoveNext
'     Loop
''    End If
'    '
'    Set RS = Nothing
'
'    mvarSCREENID.Sort = True
'
'End Function

'Public Function LoadPthDoct()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarPTHDOCT Is Nothing Then
'        Set mvarPTHDOCT = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarPTHDOCT.Clear
'
'    mvarPTHDOCT.FieldInialize "pthdoctid", "pthdoctnm"
'
'    mvarPTHDOCT.Sort = False
'    '
'    '판독의 Load
'    RS.Open objSQL.LoadPthDoctId, DbConn
''    If ColCnt > 0 Then
'        Do Until RS.EOF
'            mvarPTHDOCT.AddNew RS.Fields("pthdoct").Value & "", RS.Fields("pthdoctnm").Value & ""
'        RS.MoveNext
'     Loop
''    End If
'    '
'    Set RS = Nothing
'
'    mvarPTHDOCT.Sort = True
'
'End Function
'Public Function LoadDoct()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarDOCT Is Nothing Then
'        Set mvarDOCT = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarDOCT.Clear
'
'    mvarDOCT.FieldInialize "doctid", "doctnm"
'
'    mvarDOCT.Sort = False
'
'    '의사정보 Load
'    RS.Open objSQL.LoadDoctId, DbConn
''    If ColCnt > 0 Then
'        Do Until RS.EOF
'            mvarDOCT.AddNew RS.Fields("doctid").Value & "", RS.Fields("doctnm").Value & ""
'        RS.MoveNext
'     Loop
''    End If
'    '
'    Set RS = Nothing
'
'    mvarDOCT.Sort = True
'
'End Function

'Public Function LoadDept()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarDEPTCD Is Nothing Then
'        Set mvarDEPTCD = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarDEPTCD.Clear
'
'    mvarDEPTCD.FieldInialize "deptcd", "deptnm"
'
'    mvarDEPTCD.Sort = False
'    '
'    '부서 Load
'    RS.Open objSQL.LoadDeptCd, DbConn
''    If ColCnt > 0 Then
'        Do Until RS.EOF
'            mvarDEPTCD.AddNew RS.Fields("deptcd").Value & "", RS.Fields("deptnm").Value & ""
'        RS.MoveNext
'     Loop
''    End If
'    '
'    Set RS = Nothing
'
'    mvarDEPTCD.Sort = True
'
'End Function

'Public Function LoadLisItem()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarLISITEM Is Nothing Then
'        Set mvarLISITEM = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarLISITEM.Clear
'
'    mvarLISITEM.FieldInialize "testcd", "testnm,abbrnm5,spccd,workarea,storecd,rndfg,labelcnt," & _
'                                        "statflags,testdiv,multifg,spcgrp,spcnm,labdiv,labrange," & _
'                                        "insurfg,statfg,testfg"
'    mvarLISITEM.Sort = False
'    '
'    '부서 Load
'    RS.Open objSQL.SqlItemList, DbConn
''    If ColCnt > 0 Then
'        Do Until RS.EOF
'            If Not mvarLISITEM.Exists("" & RS.Fields("testcd").Value) Then
'                mvarLISITEM.AddNew "" & RS.Fields("testcd").Value, _
'                                    "" & RS.Fields("testnm").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("abbrnm5").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("spccd").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("workarea").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("storecd").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("rndfg").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("labelcnt").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("statflags").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("testdiv").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("multifg").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("spcgrp").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("spcnm").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("labdiv").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("labrange").Value & objDiv.COL_DIV & _
'                                    "" & RS.Fields("insurfg").Value & objDiv.COL_DIV & _
'                                    "0" & objDiv.COL_DIV & "0"
'            End If
'        RS.MoveNext
'     Loop
''    End If
'    '
'    Set RS = Nothing
'
'    mvarLISITEM.Sort = True
'
'End Function

'Public Function LoadLisSpc()
'    Dim RS As New Recordset
'    Dim ColCnt As Long
'
'    If mvarLISSPC Is Nothing Then
'        Set mvarLISSPC = New clsDictionary
'    Else
'        Exit Function
'    End If
'
'    mvarLISSPC.Clear
'
'    mvarLISSPC.FieldInialize "spccd", "spcnm,spcabbr,spcbarnm,multifg,spcgrp,labrange"
'    mvarLISSPC.Sort = False
'
'    '
'    '부서 Load
'    RS.Open objSQL.SqlSpcList, DbConn
''    If ColCnt > 0 Then
'        Do Until RS.EOF
'            mvarLISSPC.AddNew "" & RS.Fields("spccd").Value, _
'                              "" & RS.Fields("spcnm").Value & objDiv.COL_DIV & _
'                              "" & RS.Fields("spcabbr").Value & objDiv.COL_DIV & _
'                              "" & RS.Fields("spcbarnm").Value & objDiv.COL_DIV & _
'                              "" & RS.Fields("multifg").Value & objDiv.COL_DIV & _
'                              "" & RS.Fields("spcgrp").Value & objDiv.COL_DIV & _
'                              "" & RS.Fields("labrange").Value & objDiv.COL_DIV
'        RS.MoveNext
'     Loop
''    End If
'
'    '
'    Set RS = Nothing
'
'    mvarLISSPC.Sort = True
'
'End Function

'Public Function LoadBarcodeInfo()
'
'
'    If mvarBarInfo Is Nothing Then
'        Set mvarBarInfo = New clsBarcode
'    Else
'        Exit Function
'    End If
'
'    With mvarBarInfo
''        Set .MyDB = DbConn
'        Set .TableInfo = objTables
'        Set .FieldInfo = objFields
'        Call .GetBarConfig
'    End With
'
'End Function


'Private Sub Class_Terminate()
'    Set mvarWARDID = Nothing
'    Set mvarDEPTCD = Nothing
'    Set mvarPTHDOCT = Nothing
'    Set mvarSCREENID = Nothing
'    Set mvarDOCT = Nothing
'    Set mvarEMPID = Nothing
'    Set mvarLISITEM = Nothing
'    Set mvarLISSPC = Nothing
'    Set mvarTCd = Nothing
'    Set mvarMCd = Nothing
'    Set mvarOCD = Nothing
'    Set mvarTemp = Nothing
'    Set objDiv = Nothing
'    Set objSQL = Nothing
'    Set objTables = Nothing
'    Set objFields = Nothing
'    Set mvarBarInfo = Nothing
'End Sub

'------------------
' PROPERTY DECLARE
'------------------

'Public Sub setDbConn(ByVal pDbConn As Connection)
'    Set DbConn = pDbConn
'End Sub

'Public Property Set EMPID(ByVal vData As clsDictionary)
'    Set mvarEMPID = vData
'End Property
'
'Public Property Get EMPID() As clsDictionary
'    If mvarEMPID Is Nothing Then LoadEmp
'    Set EMPID = mvarEMPID
'End Property
'
'Public Property Set Doct(ByVal vData As clsDictionary)
'    Set mvarDOCT = vData
'End Property
'
'Public Property Get Doct() As clsDictionary
'    If mvarDOCT Is Nothing Then LoadDoct
'    Set Doct = mvarDOCT
'End Property
'
'Public Property Set SCREENID(ByVal vData As clsDictionary)
'    Set mvarSCREENID = vData
'End Property
'
'Public Property Get SCREENID() As clsDictionary
'    If mvarSCREENID Is Nothing Then LoadScreen
'    Set SCREENID = mvarSCREENID
'End Property
'
'Public Property Set PTHDOCT(ByVal vData As clsDictionary)
'    Set mvarPTHDOCT = vData
'End Property
'
'Public Property Get PTHDOCT() As clsDictionary
'    If mvarPTHDOCT Is Nothing Then LoadPthDoct
'    Set PTHDOCT = mvarPTHDOCT
'End Property
'
'Public Property Set DEPTCD(ByVal vData As clsDictionary)
'    Set mvarDEPTCD = vData
'End Property
'
'Public Property Get DEPTCD() As clsDictionary
'    If mvarDEPTCD Is Nothing Then LoadDept
'    Set DEPTCD = mvarDEPTCD
'End Property
'
'Public Property Set WARDID(ByVal vData As clsDictionary)
'    Set mvarWARDID = vData
'End Property
'
'Public Property Get WARDID() As clsDictionary
'    If mvarWARDID Is Nothing Then LoadWard
'    Set WARDID = mvarWARDID
'End Property
'
'Public Property Set RESID(ByVal vData As clsDictionary)
'    Set mvarRESID = vData
'End Property
'
'Public Property Get RESID() As clsDictionary
'    If mvarRESID Is Nothing Then LoadResDoct
'    Set RESID = mvarRESID
'End Property

'Public Property Set TCd(ByVal vData As clsDictionary)
'    Set mvarTCd = vData
'End Property
'
'Public Property Get TCd() As clsDictionary
'    If mvarTCd Is Nothing Then LoadTSnomed
'    Set TCd = mvarTCd
'End Property

'Public Property Set MCd(ByVal vData As clsDictionary)
'    Set mvarMCd = vData
'End Property
'
'Public Property Get MCd() As clsDictionary
'    If mvarMCd Is Nothing Then LoadMSnomed
'    Set MCd = mvarMCd
'End Property

'Public Property Set OCD(ByVal vData As clsDictionary)
'    Set mvarOCD = vData
'End Property
'
'Public Property Get OCD() As clsDictionary
'    If mvarOCD Is Nothing Then LoadOrgans
'    Set OCD = mvarOCD
'End Property

'Public Property Set Building(ByVal vData As clsDictionary)
'    Set mvarBuilding = vData
'End Property
'
'Public Property Get Building() As clsDictionary
'    If mvarBuilding Is Nothing Then LoadBuilding
'    Set Building = mvarBuilding
'End Property
'
'Public Property Set LisItem(ByVal vData As clsDictionary)
'    Set mvarLISITEM = vData
'End Property
'
'Public Property Get LisItem() As clsDictionary
'    If mvarLISITEM Is Nothing Then LoadLisItem
'    Set LisItem = mvarLISITEM
'End Property
'
'Public Property Set LisSpc(ByVal vData As clsDictionary)
'    Set mvarLISSPC = vData
'End Property
'
'Public Property Get LisSpc() As clsDictionary
'    If mvarLISSPC Is Nothing Then Call LoadLisSpc
'    Set LisSpc = mvarLISSPC
'End Property

'Public Property Set Temp(ByVal vData As clsDictionary)
'    Set mvarTemp = vData
'End Property
'
'Public Property Get Temp() As clsDictionary
'    If mvarTemp Is Nothing Then LoadCom004Temp
'    Set Temp = mvarTemp
'End Property

'Public Property Set BarInfo(ByVal vData As clsBarcode)
'    Set mvarBarInfo = vData
'End Property

'Public Property Get BarInfo() As clsBarcode
'    If mvarBarInfo Is Nothing Then LoadBarcodeInfo
'    Set BarInfo = mvarBarInfo
'End Property


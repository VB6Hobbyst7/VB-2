VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "comCLSCHlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'
' 수정일자 : 2000.01.31 YEJ
'              #1 clsLQryWhere2 에 수정
'

Private mCallForm       As Form     '콜한 폼이름
Private mFormName       As Form     '폼 이름
Private mFormWidth      As Long     '폼 넓이
Private mFormHeight     As Long     '폼 길이
Private mPosX           As Integer  'X 좌표
Private mPosY           As Integer  'Y 좌표
Private mObjectHeight   As Integer  '
Private mSpreadName     As vaSpread 'Spread 이름

Private mCol            As Integer  '현재 Colum이나 배열첨자
Private mMaxCols        As Integer  'Spread Max Colum
Private mMaxRows        As Integer  'Spread Max Row
Private mColTitle()     As String   'Spread Colum Title
Private mColAlign()     As Integer  'Spread Colum 정렬
Private mColWidth()     As Long     'Spread Colum 넓이
Private mColHidden()    As Integer  'Spread Colum 숨김여부
Private mColFormat()    As String   '값의 형식

Private mRecordCount    As Long     '레코드의 갯수
Private mServerName     As String   'Server Name
Private mLogin          As Long     'Server Login Value
Private mDBOpen         As Integer  '데이타베이스 Open 여부(True/False)
Private mSqlConn        As Long     'Server Connent Value
Private mDBTable        As String   'Table Name
Private mDBField()      As String   'Select하고 하는 필드명
Private mDBValue()      As String   'Select한 값
Private mQryWhere       As String   'Where절
Private mQryWhere2      As String   '추가 Where절
Private mQryOrder       As String   'Order By절
Private mQryGroup       As String   'Group By절
Private mQrySqlStr      As String   'Sql문장
Private mQrySetRow      As Integer  'set Rowcount 여부

Private mOColum         As Integer
Private mNColum         As Integer
Private mSort           As Integer

Private Type Gbn_Member
    code    As String
    Name    As String
End Type

Private mGbnMemberCnt   As Integer
Private mGbnData()      As Gbn_Member

'/  검사항목선택화면사용시
Private mORDDept    As String   '과코드
Private mORDSlip    As String   'Slip코드
Private mActionGbn  As String   '99.11.26 YEJ추가

'local variable(s) to hold property value(s)

Private mclsLSqlDistinct As Boolean  'Distinct   99/1/3 ojm add


Public Property Get clsGActionGbn() As String
    
    clsGActionGbn = mActionGbn
    
End Property

Public Property Let clsLActionGbn(sPara As String)

    mActionGbn = sPara
    
End Property

Public Property Let clsLSpreadName(spd As vaSpread)
    
    Set mSpreadName = spd
    
End Property

' 99/1/3 ojm add
Public Property Let clsLSqlDistinct(ByVal vData As Boolean)
    
    mclsLSqlDistinct = vData
    
End Property

Public Property Get clsGCallForm() As Form

    Set clsGCallForm = mCallForm
    
End Property

'/* *   *   *   *   *   *   *   *   *   */
'/*                                     */
'/* 코드 Check에 화면을 디자인한다.     */
'/*                                     */
'/* *   *   *   *   *   *   *   *   *   */

'/* *   *   *   *   *   *   *   *   *   *   *   */
'/*                                             */
'/* 코드CHECK화면에 데이타를 표시한다.          */
'/*                                             */
'/* *   *   *   *   *   *   *   *   *   *   *   */

Public Property Get clsGClassName() As Object

    Set clsGClassName = mClassName
    
End Property

Public Property Get clsGColFormat() As String

    clsGColFormat = mColFormat(mCol)
    
End Property

Public Property Get clsGColHidden() As Integer

    clsGColHidden = mColHidden(mCol)

End Property



Public Property Get clsGFormHeight() As Long

    clsGFormHeight = mFormHeight
    
End Property

Public Property Get clsGGbnDataCode() As String

    clsGGbnDataCode = mGbnData(mCol).code
    
End Property

Public Property Get clsGGbnDataName() As String

    clsGGbnDataName = mGbnData(mCol).Name
    
End Property


Public Property Get clsGGbnmenberCnt() As String

    clsGGbnmenberCnt = mGbnMemberCnt
    
End Property

Public Property Get clsGObjectHeight() As Integer

    clsGObjectHeight = mObjectHeight
    
End Property

Public Property Get clsGOColum()

    clsGOColum = mOColum
    
End Property

Public Property Get clsGDbField()
    
    clsgGbfield = mDBField(mCol)
    
End Property



Public Property Get clsGNColum()

    clsGNColum = mNColum
    
End Property


Public Property Get clsGPosX()

    clsGPosX = mPosX
    
End Property

Public Property Get clsGPosY()

    clsGPosY = mPosY
    
End Property


Public Property Get clsGQryWhere() As String

    clsGQryWhere = mQryWhere
    
End Property

Public Property Get clsGRecordCount()

    clsGRecordCount = mRecordCount
    
End Property

Public Property Let clsLCallForm(frm As Form)

    Set mCallForm = frm
    
End Property

Public Property Let clsLColFormat(Para As String)

    mColFormat(mCol) = Para
    
End Property

Public Property Let clsLColHidden(flag As Integer)
    
    mColHidden(mCol) = flag
    
End Property

Public Property Let clsLORDDept(Para As String)

    mORDDept = Para
    
End Property

Public Property Let clsLORDSlip(Para As String)

    mORDSlip = Para
    
End Property



Public Property Let clsLFormHeight(l As Long)

    mfromheight = l
    
End Property

Public Property Let clsLGbnDataCode(Para As String)

    mGbnData(mCol).code = Para
    
End Property

Public Property Let clsLGbnDataName(Para As String)

    mGbnData(mCol).Name = Para
    
End Property

Public Property Let clsLGbnMemberCnt(n As Integer)

    mGbnMemberCnt = n
    
End Property

Public Property Let clsLObjectHeight(i As Integer)

    mObjectHeight = i
    
End Property

Public Property Let clsLOColum(n As Integer)

    mOColum = n
    
End Property


Public Property Let clsLNColum(n As Integer)

    mNColum = n
    
End Property



Public Property Let clsLPosX(n As Integer)

    mPosX = n
    
End Property


Public Property Let clsLPosY(n As Integer)

    mPosY = n
    
End Property



Public Property Let clsLQryGroup(Para As String)

    mQryGroup = Para

End Property

Public Property Let clsLQrySetRow(n As Integer)

    mQrySetRow = n
    
End Property

Public Property Let clsLQrySqlStr(Para As Integer)

    mQrySqlStr = Para
    
End Property


Public Property Let clsLQryWhere2(Para As String)

    If mCol = 0 Then mCol = 1   '00.01.31 YEJ 추가
    
    '00.01.31 YEJ (mDBField(1) -> mDBFied(mCol))
    If mQryWhere = "" Then
        mQryWhere2 = " WHERE " + mDBField(mCol) + " LIKE '" & Para & "%'"
    Else
        mQryWhere2 = "   AND " _
                   + mDBField(mCol) + " LIKE '" & Para & "%'"
    End If
    
End Property

Public Property Let clsLMaxRows(n As Integer)

    mMaxRows = n
    
End Property

Public Property Set clsSClassName(cls As Object)

    Set mClassName = cls
    
End Property



Public Sub subDefindCodeCHK()

    With mFormName
    
    If mMaxRows = 0 Then mMaxRows = 10
    
    .pnlMain.Height = 9000
    
    .spdDisp.MaxCols = mMaxCols + 1: .spdDisp.MaxRows = mMaxRows
    
    For i = 1 To mMaxCols
        .spdDisp.Col = i + 1
        If mColHidden(i) = True Then .spdDisp.ColHidden = True
                
        .spdDisp.Row = 1
        .spdDisp.Col = i + 1: .spdDisp.TypeButtonText = mColTitle(i)
        
        If mColWidth(i) < 10 _
           And Not .spdDisp.ColHidden Then mColWidth(i) = 10

        .spdDisp.ColWidth(i + 1) = mColWidth(i)
        
    Next
    
    ret = DoEvents

    mFormWidth = .spdDisp.Width + 260
    
    .Label1.TOP = .spdDisp.TOP + .spdDisp.Height + 120
    
    .CmdEsc.TOP = .spdDisp.TOP + .spdDisp.Height + 75
    .CmdEsc.LEFT = mFormWidth - .CmdEsc.Width - 200
    
    .CmdView.TOP = .spdDisp.TOP + .spdDisp.Height + 75
    .CmdView.LEFT = mFormWidth - .CmdEsc.Width * 2 - 200
    
    .txtCd.TOP = .spdDisp.TOP + .spdDisp.Height + 75
    .txtCd.Width = mFormWidth - .CmdEsc.Width * 2 _
                 - .txtCd.LEFT - 250

    .pnlMsg.Width = mFormWidth - 200
    
    .pnlMain.Height = .CmdEsc.TOP + .CmdEsc.Height + 100
        
    mFormHeight = .pnlMain.TOP + .pnlMain.Height _
                + .pnlMsg.TOP + .pnlMsg.Height + 150
    
    End With

End Sub

Public Sub subDefindGbnData()

    ReDim mDBField(1 To 2) As String
    ReDim mDBValue(1 To 2) As String
    ReDim mGbnData(1 To mGbnMemberCnt) As Gbn_Member
    
End Sub



Public Sub subDefindCodeRST()

    With mFormName
    
    If mMaxRows = 0 Then mMaxRows = 10
    
    .pnlMain.Height = 9000
    
    .spdDisp.MaxCols = mMaxCols: .spdDisp.MaxRows = mMaxRows
    
    For i = 1 To mMaxCols
        .spdDisp.Row = 1
        .spdDisp.Col = i: .spdDisp.TypeButtonText = mColTitle(i)
        
        If mColWidth(i) < 10 _
           And Not mColHidden(i) Then mColWidth(i) = 10

        .spdDisp.ColWidth(i) = mColWidth(i)
        
    Next
    
    ret = DoEvents

    mFormWidth = .spdDisp.Width + 260
    
    .pnlMsg.Width = mFormWidth - 200
    
    .pnlMain.Height = .spdDisp.TOP + .spdisp.Height + 100
        
    mFormHeight = .pnlMain.TOP + .pnlMain.Height _
                + .pnlMsg.TOP + .pnlMsg.Height + 150
    
    End With

End Sub

Public Sub subDisplayCodeCHK()

    Dim sqldoc      As String, sqlRet   As Integer
    Dim qryTmp      As String
    
    Dim SqlData()   As String, record   As String
    
    Dim idx As Integer
    
    Screen.MousePointer = vbHourglass
    
    mRecordCount = 2
    Call D0SUB_SpreadClear(mFormName.spdDisp, 100)
    
    With mFormName.spdDisp
    
    For idx = 1 To mMaxCols
        .Col = idx + 1: .Col2 = idx + 1
        .Row = 2:       .Row2 = .MaxRows
        .BlockMode = True
        .TypeHAlign = mColAlign(idx)
        .ColHidden = mColHidden(idx)
        .BlockMode = False
    Next
    
    End With
    
    With D0COM_TempDB
    
    sqldoc = "select distinct "
    
    For idx = 1 To mMaxCols
        sqldoc = sqldoc & mDBField(idx) & ", "
    Next
    
    .clsLTempDBIndex = 1
    sqldoc = sqldoc + "T." + .clsGTempDBKeyFd _
           + "  from " & mDBTable + ", " + .clsGTempDBTable + " T"
    If Not mDBValue(1) = "" Then
        sqldoc = sqldoc + " where " & mDBField(1) & " = '" & mDBValue(1) & "'"
        qryTmp = "   and " + mDBField(1) + " *= " + "T." + .clsGTempDBKeyFd _
               + "   and "
    Else
        qryTmp = " where " + mDBField(1) + " *= " + "T." + .clsGTempDBKeyFd _
               + "   and "
    End If
    
    If Not mQryWhere = "" Then sqldoc = sqldoc & qryTmp & mQryWhere
    If Not mQryWhere2 = "" Then sqldoc = sqldoc & mQryWhere2
    If Not mQryGroup = "" Then sqldoc = sqldoc + " group by " & mQryGroup
    If Not mQryOrder = "" Then sqldoc = sqldoc + " order by " & mQryOrder

    End With
    
    sqlRet = QSqlDBExec(sqldoc, mSqlConn)
    If sqlRet = QSQL_SUCCESS Then
        Do Until Not QSqlGetRow(record, mSqlConn) = QSQL_SUCCESS
        
            QSqlGetField mMaxCols + 1, record, SqlData()

            With mFormName.spdDisp
            
                If mRecordCount > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    
                    For idx = 1 To mMaxCols
                        .Col = idx + 1:     .Col2 = idx + 1
                        .Row = .MaxRows:    .Row2 = .MaxRows
                        .BlockMode = True
                        .TypeHAlign = mColAlign(idx)
                        .BlockMode = False
                    Next
                    
                End If
                
                .Row = mRecordCount
                For idx = 1 To mMaxCols
                    .Col = idx + 1: .Text = SqlData(idx)
                Next
                
                If SqlData(1) = SqlData(mMaxCols + 1) Then
                    .Col = 1: .Text = "1"
                End If
                    
            End With
            
            mRecordCount = mRecordCount + 1
            
        Loop
    End If
    Call QSqlSelectFree(mSqlConn)
    
    If mRecordCount >= mMaxRows + 1 Then
        mFormName.spdDisp.MaxRows = mRecordCount - 1
    Else
        mFormName.spdDisp.MaxRows = mMaxRows + 1
    End If

    If mRecordCount = 2 Then
        Beep
        mRecordCount = 0
        mFormName.pnlMsg = "데이타가 없습니다."
    Else
        mRecordCount = mRecordCount - 1
        mFormName.pnlMsg = " 데이타 조회 건수 :  " & mRecordCount
    End If
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub subDisplayCodeRST()

    Dim sqldoc      As String, sqlRet   As Integer
    Dim qryTmp      As String
    
    Dim SqlData()   As String, record   As String
    
    Dim idx As Integer, Row As Integer
    
    Screen.MousePointer = vbHourglass
    
    mRecordCount = 2
    Call D0SUB_SpreadClear(mFormName.spdDisp, 100)
    
    With mFormName.spdDisp
    
    For idx = 1 To mMaxCols
        .Col = idx: .Col2 = idx
        .Row = 2:   .Row2 = .MaxRows
        .BlockMode = True
        .TypeHAlign = mColAlign(idx)
        .BlockMode = False
    Next
    
    End With
    
    sqlRet = QSqlDBExec(mQrySqlStr, mSqlConn)
    If sqlRet = QSQL_SUCCESS Then
        Do Until Not QSqlGetRow(record, mSqlConn) = QSQL_SUCCESS
        
            QSqlGetField mMaxCols, record, SqlData()

            With mFormName.spdDisp
            
            If mRecordCount > .MaxRows Then
                .MaxRows = .MaxRows + 1
                
                For idx = 1 To mMaxCols
                    .Col = idx:         .Col2 = idx
                    .Row = .MaxRows:    .Row2 = .MaxRows
                    .BlockMode = True
                    .TypeHAlign = mColAlign(idx)
                    .BlockMode = False
                Next
                
            End If
            
            .Row = mRecordCount
            For idx = 1 To mMaxCols
                .Col = idx
                If mColFormat(idx) = "" Then
                    .Text = SqlData(idx)
                Else
                    .Text = Format$(Format$(SqlData(idx), "#######"), mColFormat(idx))
                End If
            Next
            
            End With
            
            mRecordCount = mRecordCount + 1
            
        Loop
    End If
    Call QSqlSelectFree(mSqlConn)
     
    If mRecordCount >= mMaxRows + 1 Then
        mFormName.spdDisp.MaxRows = mRecordCount - 1
    Else
        mFormName.spdDisp.MaxRows = mMaxRows + 1
    End If

    If mRecordCount = 2 Then
        Beep
        mRecordCount = 0
        mFormName.pnlMsg = "데이타가 없습니다."
    Else
        mRecordCount = mRecordCount - 2
    End If
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub subDisplayORDOrd()

    Dim sqldoc      As String, sqlRet   As Integer
    Dim qryTmp      As String
    
    Dim SqlData()   As String, record   As String
    
    Dim iRow        As Integer
    
    Screen.MousePointer = vbHourglass
    
    iRow = 2:   Call D0SUB_SpreadClear(mFormName.spdOrd, 11)
    
    With mFormName.spdOrd
    
    If mORDDept = "" Then
        sqldoc = "select distinct " _
               & "       A.SLIPCD+A.ORDCD+A.SPCCD, A.ORDNM, A.ORDCHK, A.ETCGBN, A.DEPTCD+A.SUGAGBN+A.SUGACD" _
               & "     , A.DEPTCD,                 T.ORDCD, T.ETCGBN, T.SUGACD" _
               & "  from BAS_DB..BAS060M A, " & D0COM_TempDB.clsGTempDBTable & " T" _
               & " where A.SLIPCD = '" & mORDSlip & "'" _
               & "     , A.SUBCD = ''" _
               & "   and A.SLIPCD+A.ORDCD+A.SPCCD *= T.ORDCD"
    Else
        sqldoc = "select distinct " _
               & "       A.SLIPCD+A.ORDCD+A.SPCCD, B.ORDNM, B.ORDCHK, B.ETCGBN, B.DEPTCD+B.SUGAGBN+B.SUGACD" _
               & "     , B.DEPTCD,                 T.ORDCD, T.ETCGBN, T.SUGACD" _
               & "  from BAS_DB..BAS210M A, BAS_DB..BAS060M B" _
               & "     , " & D0COM_TempDB.clsGTempDBTable & " T" _
               & " where A.DEPTCD = '" & mORDDept & "'" _
               & "   and A.SLIPCD = '" & mORDSlip & "'" _
               & "   and A.SLIPCD = B.SLIPCD" _
               & "   and A.ORDCD  = B.ORDCD and A.SPCCD = B.SPCCD" _
               & "   and B.SUBCD  = ''" _
               & "   and A.SLIPCD+A.ORDCD+A.SPCCD *= T.ORDCD"
    End If
    
    sqlRet = QSqlDBExec(sqldoc, mSqlConn)
    If sqlRet = QSQL_SUCCESS Then
        Do Until Not QSqlGetRow(record, mSqlConn) = QSQL_SUCCESS
        
            QSqlGetField 9, record, SqlData()

            If iRow > .MaxRows Then .MaxRows = .MaxRows + 1
            
            .Row iRow
            .Col = 2:   .Text = SqlData(2)
            .Col = 3:   .Text = SqlData(1)
            .Col = 4:   .Text = SqlData(3)
            
            If SqlData(1) = SqlData(7) Then
                .Col = 5:   .Text = SqlData(8)
                .Col = 1:   .Text = "1"
            Else
                .Col = 5:   .Text = SqlData(4)
            End If
            
            .Col = 6:   .Text = SqlData(5)
            .Col = 7:   .Text = SqlData(6)
            
            iRow = iRow + 1
        Loop
    End If
    Call QSqlSelectFree(mSqlConn)
    End With
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub subDisplayORDSlip()

    Dim sqldoc      As String, sqlRet       As Integer
    Dim SqlData()   As String, sqlRecord    As String
    
    Dim iRow        As Integer
    
    Screen.MousePointer = vbHourglass
    
    Call D0SUB_SpreadClear(mFormName.spdSlip, 11, , , 2)
     
    With mFormName.spdSlip
    
    For iRow = 2 To .MaxRows
        
        .Row = iRow
        .Col = 1:   .TypeButtonText = ""
    Next
    
    iRow = 2
    
    If mORDDept = "" Then
        sqldoc = "select SCGBCD, CDGBNM" _
               + "  from BAS_DB..BAS031M"
    Else
        sqldoc = "select distinct A.SLIPCD, B.CDGBNM" _
               + "  from BAS_DB..BAS210M A, BAS_DB..BAS031M B" _
               + " where A.DEPTCD = '" + mORDDept + "'" _
               + "   and A.SLIPCD = B.SCGBCD"
    End If
    sqlRet = QSqlDBExec(sqldoc, mSqlConn)
    If sqlRet = QSQL_SUCCESS Then
        Do Until Not QSqlGetRow(sqlRecord, mSqlConn) = QSQL_SUCCESS
        
            QSqlGetField 2, sqlRecord, SqlData()

            If iRow > .MaxRows Then .MaxRows = .MaxRows + 1
            
            .Row = iRow
            .Col = 1:   .TypeButtonText = SqlData(2)
            .Col = 2:   .Text = SqlData(1)
            
            iRow = iRow + 1
        Loop
    End If
    Call QSqlSelectFree(mSqlConn)
    End With
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub subSpreadDataSort(ByVal Col As Integer)

    With mFormName.spdDisp
    
        .Col = 1:   .Col2 = .MaxCols
        .Row = 2:   .Row2 = mRecordCount + 1
        .SortBy = SS_SORT_BY_ROW
        .SortKey(1) = Col
        If Col = mOColum Then
            mSort = IIf(mSort = SS_SORT_ORDER_ASCENDING _
                              , SS_SORT_ORDER_DESCENDING _
                              , SS_SORT_ORDER_ASCENDING)
        Else
            mSort = SS_SORT_ORDER_ASCENDING
        End If
        
        .SortKeyOrder(1) = mSort
        
        If Col = .MaxCols Then
            .SortKey(2) = 1
        Else
            .SortKey(2) = Col + 1
        End If
        
        .SortKeyOrder(2) = SS_SORT_ORDER_ASCENDING
        .Action = SS_ACTION_SORT
     
    End With
    
    mOColum = Col

End Sub

'/* *   *   *   *   *   *   *   *   *   *   *   */
'/*                                             */
'/* TABLE에서 특정필드의 데이타값을 얻어온다.   */
'/*                                             */
'/* *   *   *   *   *   *   *   *   *   *   *   */
Public Sub subGetCodeName()

    Dim SqlData()   As String, record   As String
    
    Dim sqldoc  As String, sqlRet As Integer
    Dim SelectCount As Integer
    
    Dim idx As Integer
    
    mRecordCount = 0
    
    If mSqlConn = 0 Then
         '본 화면에서 사용할 Index Open
        If QSqlOpen(D0COM_SERVER01, mFormName.hWnd, mSqlConn) _
           <> QSQL_SUCCESS Then Exit Sub
    
        mDBOpen = True
        
    End If

    If mQrySetRow > 0 Then
        sqldoc = "  set rowcount " & mQrySetRow & " "
    End If
    
'    sqldoc = sqldoc & "select distinct "  ' 99/1/3 ojm rem
    '99/1/3 ojm 수정
    
    If mclsLSqlDistinct = False Then
        sqldoc = sqldoc & "select "
    Else
        sqldoc = sqldoc & "select distinct "
    End If
    
    For idx = 1 To mMaxCols
        sqldoc = sqldoc & mDBField(idx) & ", "
    Next
    
    sqldoc = Mid$(sqldoc, 1, Len(sqldoc) - 2) _
           & "  from " & mDBTable
    If Not mDBValue(1) = "" Then
        sqldoc = sqldoc & " where " & mDBField(1) & " = '" & mDBValue(1) & "'"
        If Not mQryWhere = "" Then sqldoc = sqldoc & "   and "
    ElseIf Not mQryWhere = "" Then
        sqldoc = sqldoc & " where "
    End If
    If Not mQryWhere = "" Then sqldoc = sqldoc & mQryWhere
    If Not mQryOrder = "" Then sqldoc = sqldoc & " order by " & mQryOrder
    If Not mQryGroup = "" Then sqldoc = sqldoc & " group by " & mQryGroup
    If mQrySetRow > 0 Then sqldoc = sqldoc & "  set rowcount 0"
    
    sqlRet = QSqlDBExec(sqldoc, mSqlConn)
    If sqlRet = QSQL_SUCCESS Then
        Do Until Not QSqlGetRow(record, mSqlConn) = QSQL_SUCCESS
            
            QSqlGetField mMaxCols, record, SqlData()

            For idx = 1 To mMaxCols
                mDBValue(idx) = SqlData(idx)
            Next
                       
            mRecordCount = mRecordCount + 1
            
            Exit Do
        Loop
    End If
    Call QSqlSelectFree(mSqlConn)

    Dim DbFlag  As Integer
    
    If mDBOpen = True Then Call QSqlClose(mSqlConn)
    
End Sub

Public Property Get clsGDBOpen()
    
    clsGDBOpen = mDBOpen
    
End Property


Public Property Get clsGDbValue()

    clsGDbValue = mDBValue(mCol)
    
End Property





Public Property Let clsLLogin(n As Long)

    mLogin = n
    
End Property


Public Property Get clsGSqlConn()
 
    clsGSqlConn = mSqlConn
    
End Property

 
Public Property Let clsLColAlign(n As Integer)

    mColAlign(mCol) = n
    
End Property

Public Property Let clsLDBOpen(n As Integer)

    mDBOpen = n
    
End Property

Public Property Let clsLDBTable(Para As String)

    mDBTable = Para
    
End Property


Public Property Let clsLColWidth(Para As Long)

    mColWidth(mCol) = Para
    
End Property

Public Property Let clsLFormName(frm As Form)

    Set mFormName = frm
    
End Property




Public Property Let clsLFormWidth(l As Long)

    mFormWidth = l
    
End Property

Public Property Get clsGFormWidth()

    clsGFormWidth = mFormWidth
    
End Property




Public Property Let clsLMaxCols(n As Integer)

    mMaxCols = n
    
End Property


Public Property Let clsLCol(n As Integer)

    mCol = n
    
End Property




Public Property Let clsLColTitle(Para As String)
    
    mColTitle(mCol) = Para
    
End Property

Public Property Let clsLDbField(Para As String)

    mDBField(mCol) = Para
    
End Property



Public Property Let clsLDBValue(Para As String)

    mDBValue(mCol) = Para
    
End Property


Public Property Let clsLQryOrder(Para As String)

    mQryOrder = Para
    
End Property

Public Property Let clsLQryWhere(Para As String)

    mQryWhere = Para
    
End Property



'/* *   *   *   *   *   *   *   *   *   */
'/*                                     */
'/* 코드 Help에 화면을 디자인한다.      */
'/*                                     */
'/* *   *   *   *   *   *   *   *   *   */
Public Sub subDefindCodeHLP()

    With mFormName
    
    If mMaxRows = 0 Then mMaxRows = 11
    
    .pnlMain.Height = 9000
    
    .spdDisp.MaxCols = mMaxCols: .spdDisp.MaxRows = mMaxRows + 1
    
    For i = 1 To mMaxCols
        .spdDisp.Row = 1
        .spdDisp.Col = i: .spdDisp.TypeButtonText = mColTitle(i)
        
        If mColWidth(i) < 10 And Not mColHidden(i) Then
            If mMaxCols = 1 Then
                mColWidth(i) = 15
            Else
                mColWidth(i) = 10
            End If
        End If

        .spdDisp.ColWidth(i) = mColWidth(i)
        
    Next
    
    ret = DoEvents

    mFormWidth = .spdDisp.Width + 260
    
    .Label1.TOP = .spdDisp.TOP + .spdDisp.Height + 120
    
    .CmdEsc.TOP = .spdDisp.TOP + .spdDisp.Height + 75
    .CmdEsc.LEFT = mFormWidth - .CmdEsc.Width - 200
    
    .CmdView.TOP = .spdDisp.TOP + .spdDisp.Height + 75
    .CmdView.LEFT = mFormWidth - .CmdEsc.Width * 2 - 200
    
    .txtCd.TOP = .spdDisp.TOP + .spdDisp.Height + 75
    .txtCd.Width = mFormWidth - .CmdEsc.Width * 2 _
                 - .txtCd.LEFT - 250

    .pnlMsg.Width = mFormWidth - 200
    
    .pnlMain.Height = .CmdEsc.TOP + .CmdEsc.Height + 100
        
    mFormHeight = .pnlMain.TOP + .pnlMain.Height _
                + .pnlMsg.TOP + .pnlMsg.Height + 150
    
    End With
    
End Sub

Public Property Let clsLServerName(Para As String)

    mServerName = Para
    
End Property

Public Property Let clsLSqlConn(l As Long)

    mSqlConn = l

End Property


'/* *   *   *   *   *   *   *   *   *   */
'/*                                     */
'/* 화면에 관련되 변수를 선언한다.      */
'/* (칼럼에 타이틀, 길이, 정렬방법)     */
'/*                                     */
'/* *   *   *   *   *   *   *   *   *   */
Public Sub subDefindCols()

    ReDim mColTitle(1 To mMaxCols) As String
    ReDim mColAlign(1 To mMaxCols) As Integer
    ReDim mColWidth(1 To mMaxCols) As Long
    ReDim mColHidden(1 To mMaxCols) As Integer
    ReDim mColFormat(1 To mMaxCols) As String
    
End Sub

'/* *   *   *   *   *   *   *   *   *   *   *   */
'/*                                             */
'/* SELECT할 데이타 관련된 변수를 선언한다.     */
'/* (칼럼에 타이틀, 길이, 정렬방법)             */
'/*                                             */
'/* *   *   *   *   *   *   *   *   *   *   *   */
Public Sub subDefindFieldValue()

    ReDim mDBField(1 To mMaxCols) As String
    ReDim mDBValue(1 To mMaxCols) As String

End Sub



'/* *   *   *   *   *   *   *   *   *   *   *   */
'/*                                             */
'/* 코드HELP화면에 데이타를 표시한다.           */
'/*                                             */
'/* *   *   *   *   *   *   *   *   *   *   *   */
Public Sub subDisplayCodeHLP()

    Dim sqldoc      As String, sqlRet   As Integer
    Dim qryTmp      As String
    
    Dim SqlData()   As String, record   As String
    
    Dim idx As Integer, Row As Integer
    
    Screen.MousePointer = vbHourglass
    
    mRecordCount = 2
    Call D0SUB_SpreadClear(mFormName.spdDisp, 100)
    
    With mFormName.spdDisp
    
    For idx = 1 To mMaxCols
        .Col = idx: .Col2 = idx
        .Row = 2:   .Row2 = .MaxRows
        .BlockMode = True
        .TypeHAlign = mColAlign(idx)
        .BlockMode = False
    Next
    
    End With
    
    If mGbnMemberCnt > 0 Then
        For Row = 1 To mGbnMemberCnt
        
            With mFormName.spdDisp
            
            If mRecordCount > .MaxRows Then
                .MaxRows = .MaxRows + 1
                
                For idx = 1 To mMaxCols
                    .Col = idx:         .Col2 = idx
                    .Row = .MaxRows:    .Row2 = .MaxRows
                    .BlockMode = True
                    .TypeHAlign = mColAlign(idx)
                    .BlockMode = False
                Next idx
                
            End If
            
            .Row = mRecordCount
            .Col = 1: .Text = mGbnData(Row).Name
            .Col = 2: .Text = mGbnData(Row).code
            
            mRecordCount = mRecordCount + 1
            
            End With
        Next Row
    Else
    
    '    sqldoc = sqldoc & "select distinct "  ' 99/1/3 ojm rem
        '99/1/3 ojm 수정
        
        If mclsLSqlDistinct = False Then
            sqldoc = sqldoc & "select "
        Else
            sqldoc = sqldoc & "select distinct "
        End If
        
        For idx = 1 To mMaxCols
            sqldoc = sqldoc & mDBField(idx) & ", "
        Next
        
        sqldoc = Mid$(sqldoc, 1, Len(sqldoc) - 2) _
               + "  from " & mDBTable
        If Not mDBValue(1) = "" Then
            sqldoc = sqldoc + " where " & mDBField(1) & " = '" & mDBValue(1) & "'"
            qryTmp = "   and "
        Else
            qryTmp = " where "
        End If
        
        If Not mQryWhere = "" Then sqldoc = sqldoc & qryTmp & mQryWhere
        If Not mQryWhere2 = "" Then sqldoc = sqldoc & mQryWhere2
        If Not mQryGroup = "" Then sqldoc = sqldoc + " group by " & mQryGroup
        If Not mQryOrder = "" Then sqldoc = sqldoc + " order by " & mQryOrder
    
        sqlRet = QSqlDBExec(sqldoc, mSqlConn)
        If sqlRet = QSQL_SUCCESS Then
            Do Until Not QSqlGetRow(record, mSqlConn) = QSQL_SUCCESS
            
                QSqlGetField mMaxCols, record, SqlData()
    
                With mFormName.spdDisp
                
                If mRecordCount > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    
                    For idx = 1 To mMaxCols
                        .Col = idx:         .Col2 = idx
                        .Row = .MaxRows:    .Row2 = .MaxRows
                        .BlockMode = True
                        .TypeHAlign = mColAlign(idx)
                        .BlockMode = False
                    Next
                    
                End If
                
                .Row = mRecordCount
                For idx = 1 To mMaxCols
                    .Col = idx
                    If mColFormat(idx) = "" Then
                        .Text = SqlData(idx)
                    Else
                        .Text = Format$(Format$(SqlData(idx), "#######"), mColFormat(idx))
                    End If
                Next
                
                End With
                
                mRecordCount = mRecordCount + 1
                
            Loop
        End If
        Call QSqlSelectFree(mSqlConn)
    End If
    
    If mRecordCount >= mMaxRows + 1 Then
        mFormName.spdDisp.MaxRows = mRecordCount '00.01.20 YEJ 수정 - 1
    Else
        mFormName.spdDisp.MaxRows = mMaxRows + 1
    End If

    If mRecordCount = 2 Then
        Beep
        mRecordCount = 0
        mFormName.pnlMsg = "데이타가 없습니다."
    Else
        mRecordCount = mRecordCount - 2
        mFormName.pnlMsg = " 데이타 조회 건수 :  " & mRecordCount
        
        mFormName.spdDisp.Row = 2: mFormName.spdDisp.Col = 2
        mFormName.spdDisp.Action = SS_ACTION_ACTIVE_CELL
        
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub Class_Initialize()
    
    mclsLSqlDistinct = True     ' 99/1/3 ojm add
    
End Sub

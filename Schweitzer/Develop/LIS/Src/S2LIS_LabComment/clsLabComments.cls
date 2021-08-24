VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLabComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarMyDb As Object '지역 복사
'Private mvarServer As String
'Private mvarDatabase As String
'Private mvarUID  As String
'Private mvarPWD As String

Private mvarDoctId As String
Private mvarDoctNm As String

Private mvarPatId As String
Private mvarPatNm As String

Private mvarErrMsg As String
Private mvarVersion As String

Private mvarBedInDt As String

'Server
'Public Property Let Server(ByVal vData As String)
'    mvarServer = vData
'End Property
'Public Property Get Server() As String
'    Server = mvarServer
'End Property

'Database
'Public Property Let Database(ByVal vData As String)
'    mvarDatabase = vData
'End Property
'
'Public Property Get Database() As String
'    Database = mvarDatabase
'End Property

'UID
'Public Property Let uID(ByVal vData As String)
'    mvarUID = vData
'End Property
'Public Property Get uID() As String
'    uID = mvarUID
'End Property
'
''PWD
'Public Property Let PWD(ByVal vData As String)
'    mvarPWD = vData
'End Property
'Public Property Get PWD() As String
'    PWD = mvarPWD
'End Property

'DB Connection 설정
Public Property Set MyDb(ByVal vData As Object)
    Set DBConn = vData
End Property

Public Property Get MyDb() As Object
    Set MyDb = DBConn
End Property

''DB Connection 설정
Public Property Set SysInfo(ByVal vData As Object)
    Set ObjSysInfo = vData
End Property

Public Property Get SysInfo() As Object
    Set SysInfo = ObjSysInfo
End Property

'의사ID
Public Property Let DoctId(ByVal vData As String)
    mvarDoctId = vData
    gDoctId = vData
End Property

Public Property Get DoctId() As String
    DoctId = mvarDoctId
End Property

'의사명
Public Property Let DoctNm(ByVal vData As String)
    mvarDoctNm = vData
    gDoctNm = vData
End Property

Public Property Get DoctNm() As String
    DoctNm = mvarDoctNm
End Property

'환자명
Public Property Let PatNm(ByVal vData As String)
    mvarPatNm = vData
    gPatNm = vData
End Property

Public Property Get PatNm() As String
    PatNm = mvarPatNm
End Property

'환자ID
Public Property Let PatId(ByVal vData As String)
    mvarPatId = vData
    gPatId = vData
End Property

Public Property Get PatId() As String
    PatId = mvarPatId
End Property

'입원일
Public Property Let BedinDt(ByVal vData As String)
    mvarBedInDt = vData
    gPatId = vData
End Property

Public Property Get BedinDt() As String
    BedinDt = mvarBedInDt
End Property


''Error Message
'Public Property Let ErrMsg(ByVal vData As String)
'    mvarErrMsg = vData
'End Property
'Public Property Get ErrMsg() As String
'    ErrMsg = mvarErrMsg
'End Property

''Version
'Public Property Get Version() As String
'    Version = App.EXEName & vbTab & App.Major & "." & App.Minor & "." & App.Revision & vbTab & "2000.3.9"
'End Property

Public Sub ShowForm()
   
    Dim RetDbCon As Integer
    Dim blnSuccess As Boolean

    If DBConn Is Nothing Then
       RetDbCon = DBConnect
        If Not IsDBOpen Then
            MsgBox "Database 연결이 안됐습니다. 전산실로 문의 바랍니다."
            Exit Sub
        End If
    End If
   
'    DCM_DbType = ObjSysInfo.dbtype
    gDoctId = mvarDoctId
    gDoctNm = mvarDoctNm
    
    gPatId = mvarPatId
    gPatNm = mvarPatNm
   
    sBedInDT = mvarBedInDt
    
    Set fMain = frmMain
   
    fMain.Show
   
End Sub

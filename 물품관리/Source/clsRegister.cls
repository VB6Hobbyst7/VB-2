VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const gRegKeyApp As String = "[SWIT] 물품관리시스템"
Const gRegKeySec As String = "개발자 설정옵션"
Const fRegSqlServer As String = "서버이름"
Const fRegSqlDb As String = "데이터베이스이름"
Const fRegSqlUser As String = "사용자이름"
Const fRegSqlPswd As String = "비밀번호"

Const fRegOraTns As String = "검진 서버이름"
Const fRegOraUser As String = "검진 사용자이름"
Const fRegOraPswd As String = "검진 비밀번호"

Const gRegKeyUser As String = "사용자 설정옵션"
Const fRegUserName As String = "사용자명"

Private mvarsqlserver As String '지역 복사
Private mvarsqldb As String '지역 복사
Private mvarsqluser As String '지역 복사
Private mvarsqlpswd As String '지역 복사
Private mvarusername As String '지역 복사
Private mvaroratns As String '지역 복사
Private mvarorauser As String '지역 복사
Private mvarorapswd As String '지역 복사

Public Property Let orapswd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.orapswd = 5
    mvarorapswd = vData
End Property


Public Property Get orapswd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.orapswd
    orapswd = mvarorapswd
End Property



Public Property Let orauser(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.orauser = 5
    mvarorauser = vData
End Property


Public Property Get orauser() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.orauser
    orauser = mvarorauser
End Property



Public Property Let oratns(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.oratns = 5
    mvaroratns = vData
End Property


Public Property Get oratns() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.oratns
    oratns = mvaroratns
End Property

Public Property Let username(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.username = 5
    mvarusername = vData
End Property


Public Property Get username() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.username
    username = mvarusername
End Property




Public Property Let sqlpswd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.sqlpswd = 5
    mvarsqlpswd = vData
End Property


Public Property Get sqlpswd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.sqlpswd
    sqlpswd = mvarsqlpswd
End Property



Public Property Let sqluser(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.sqluser = 5
    mvarsqluser = vData
End Property


Public Property Get sqluser() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.sqluser
    sqluser = mvarsqluser
End Property



Public Property Let sqldb(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.sqldb = 5
    mvarsqldb = vData
End Property


Public Property Get sqldb() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.sqldb
    sqldb = mvarsqldb
End Property



Public Property Let sqlserver(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.sqlserver = 5
    mvarsqlserver = vData
End Property


Public Property Get sqlserver() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.sqlserver
    sqlserver = mvarsqlserver
End Property

Public Sub csRegisterSave()

    SaveSetting gRegKeyApp, gRegKeySec, fRegSqlServer, mvarsqlserver
    SaveSetting gRegKeyApp, gRegKeySec, fRegSqlDb, mvarsqldb
    SaveSetting gRegKeyApp, gRegKeySec, fRegSqlUser, mvarsqluser
    SaveSetting gRegKeyApp, gRegKeySec, fRegSqlPswd, mvarsqlpswd

    SaveSetting gRegKeyApp, gRegKeySec, fRegUserName, mvarusername

    SaveSetting gRegKeyApp, gRegKeySec, fRegOraTns, mvaroratns
    SaveSetting gRegKeyApp, gRegKeySec, fRegOraUser, mvarorauser
    SaveSetting gRegKeyApp, gRegKeySec, fRegOraPswd, mvarorapswd

End Sub

Private Sub Class_Initialize()

    mvarusername = GetSetting(appname:=gRegKeyApp, section:=gRegKeyUser, Key:=fRegUserName, Default:="(주)성원")

    mvarsqlserver = GetSetting(appname:=gRegKeyApp, section:=gRegKeySec, Key:=fRegSqlServer, Default:="127.0.0.1")
    mvarsqldb = GetSetting(appname:=gRegKeyApp, section:=gRegKeySec, Key:=fRegSqlDb, Default:="GOODSMNG")
    mvarsqluser = GetSetting(appname:=gRegKeyApp, section:=gRegKeySec, Key:=fRegSqlUser, Default:="sa")
    mvarsqlpswd = GetSetting(appname:=gRegKeyApp, section:=gRegKeySec, Key:=fRegSqlPswd, Default:="7297")

    mvaroratns = GetSetting(appname:=gRegKeyApp, section:=gRegKeySec, Key:=fRegOraTns, Default:="KAHP")
    mvarorauser = GetSetting(appname:=gRegKeyApp, section:=gRegKeySec, Key:=fRegOraUser, Default:="tw_mis_med")
    mvarorapswd = GetSetting(appname:=gRegKeyApp, section:=gRegKeySec, Key:=fRegOraPswd, Default:="hospital")

End Sub

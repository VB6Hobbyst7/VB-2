VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis999"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarareacd As String '지역 복사
Private mvarareanm As String '지역 복사
Private mvarchanggofg As String '지역 복사
Private mvarareatype As String '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
Private mvarpswd As String '지역 복사
'속성값을 가질 지역 변수
Private mvarautoentfg As String '지역 복사
Public Property Let autoentfg(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.autoentfg = 5
    mvarautoentfg = vData
End Property


Public Property Get autoentfg() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.autoentfg
    autoentfg = mvarautoentfg
End Property




Public Property Let pswd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.pswd = 5
    mvarpswd = vData
End Property


Public Property Get pswd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.pswd
    pswd = mvarpswd
End Property

Public Property Let moddt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.moddt = 5
    mvarmoddt = vData
End Property


Public Property Get moddt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.moddt
    moddt = mvarmoddt
End Property



Public Property Let wrtdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.wrtdt = 5
    mvarwrtdt = vData
End Property


Public Property Get wrtdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.wrtdt
    wrtdt = mvarwrtdt
End Property



Public Property Let areatype(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.areatype = 5
    mvarareatype = vData
End Property


Public Property Get areatype() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.areatype
    areatype = mvarareatype
End Property



Public Property Let changgofg(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.changgofg = 5
    mvarchanggofg = vData
End Property


Public Property Get changgofg() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.changgofg
    changgofg = mvarchanggofg
End Property



Public Property Let areanm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.areanm = 5
    mvarareanm = vData
End Property


Public Property Get areanm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.areanm
    areanm = mvarareanm
End Property



Public Property Let areacd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.areacd = 5
    mvarareacd = vData
End Property


Public Property Get areacd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.areacd
    areacd = mvarareacd
End Property

Public Function cfSave() As Boolean

    On Error GoTo errSave
    gSql = "SELECT * FROM S2PIS999 WHERE LOCALCD='1'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("LOCALCD").Value = "1"
                .Fields("WRTDT").Value = gfSystemDate
            End If
            .Fields("AREACD").Value = areacd
            .Fields("AREANM").Value = areanm
            .Fields("CHANGGOFG").Value = changgofg
            .Fields("AREATYPE").Value = areatype
            .Fields("PSWD").Value = pswd
            .Fields("AUTOENTFG").Value = autoentfg
            .Fields("MODDT").Value = gfSystemDate
            .Update
            .Close
        End If
    End With
    cfSave = True
    Exit Function
    
errSave:
    MsgBox Err.Description, vbCritical
    cfSave = False
    
End Function

Public Function cfSeek() As Boolean

    gSql = "SELECT * FROM S2PIS999 WHERE LOCALCD='1'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                cfSeek = True
                
                areacd = "" & .Fields("AREACD").Value
                areanm = "" & .Fields("AREANM").Value
                changgofg = "" & .Fields("CHANGGOFG").Value
                areatype = "" & .Fields("AREATYPE").Value
                wrtdt = "" & .Fields("WRTDT").Value
                moddt = "" & .Fields("MODDT").Value
                pswd = "" & .Fields("PSWD").Value
                autoentfg = "" & .Fields("AUTOENTFG").Value
            End If
            .Close
        End If
    End With

End Function

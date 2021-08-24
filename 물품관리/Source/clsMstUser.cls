VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMstUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarusercd As String '지역 복사
Private mvarusernm As String '지역 복사
Private mvarpswd As String '지역 복사
Private mvarlevelfg As Integer '지역 복사
Private mvardelfg As Integer '지역 복사
Private mvardutycd As String '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
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



Public Property Let dutycd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.dutycd = 5
    mvardutycd = vData
End Property


Public Property Get dutycd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.dutycd
    dutycd = mvardutycd
End Property



Public Property Let delfg(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.delfg = 5
    mvardelfg = vData
End Property


Public Property Get delfg() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.delfg
    delfg = mvardelfg
End Property



Public Property Let levelfg(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.levelfg = 5
    mvarlevelfg = vData
End Property


Public Property Get levelfg() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.levelfg
    levelfg = mvarlevelfg
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



Public Property Let usernm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.usernm = 5
    mvarusernm = vData
End Property


Public Property Get usernm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.usernm
    usernm = mvarusernm
End Property



Public Property Let usercd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.usercd = 5
    mvarusercd = vData
End Property


Public Property Get usercd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.usercd
    usercd = mvarusercd
End Property

Public Function cfSave() As Boolean
Dim sDate As String

    On Error GoTo errSave
    sDate = Format(gfSystemDate, "yyyy-MM-dd Hh:Nn:Ss")
    gSql = "select * from mstUSER where usercd = '" & usercd & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("usercd").Value = usercd
                .Fields("wrtdt").Value = sDate
            End If
            .Fields("usernm").Value = usernm
            .Fields("pswd").Value = pswd
            .Fields("levelfg").Value = levelfg
            .Fields("delfg").Value = delfg
            .Fields("dutycd").Value = dutycd
            .Fields("moddt").Value = sDate
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

Public Function cfDelete(ByVal brCd As String) As Boolean
Dim sReturn As Boolean

    sReturn = True
    ' 출고자료 확인
    gSql = "select usercd from outL with (nolock) where usercd = '" & brCd & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            sReturn = .EOF
            .Close
        End If
    End With
    
    If sReturn Then
        ' 요청자료 확인
        gSql = "select usercd from reqL with (nolock) where usercd = '" & brCd & "'"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        ' 발주자료 확인
        gSql = "select usercd from ordH with (nolock) where usercd = '" & brCd & "'"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        ' 구매자료 확인
        gSql = "select usercd from buyL with (nolock) where usercd = '" & brCd & "'"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        ' 장비운영자료 확인
        gSql = "select usercd from operL with (nolock) where usercd = '" & brCd & "'"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        ' 수동검사자료 확인
        gSql = "select usercd from dayMENU with (nolock) where usercd = '" & brCd & "'"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        gSql = "delete from mstUSER where usercd = '" & brCd & "'"
    Else
        gSql = "update mstUSER set delfg = " & gDelYes & " where usercd = '" & brCd & "'"
    End If
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(Optional ByVal brDel As Byte = gAllData) As ADODB.Recordset

    gSql = "select a.*, b.dutynm from mstUSER a, mstDUTY b where a.dutycd *= b.dutycd "
    If brDel <> gAllData Then
        gSql = gSql & " and a.delfg = " & brDel
    End If
    gSql = gSql & " order by a.usercd"
    Set cfList = cDb.cfRecordSet(gSql)
    
End Function

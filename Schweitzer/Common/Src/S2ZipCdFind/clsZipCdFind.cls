VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZipCdFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarZipCd As String
Private mvarProvince As String
Private mvarDistrict As String
Private mvarVillage As String
Private mvarAddrNo As String

Private MyForm As frmZipCdFind

'Public Sub setDbConn(ByRef pDbConn As DrDatabase)
''디비 커넥션
'
'    Set ZipFindDbConn = pDbConn
'End Sub

'Public Function RecordSetOpen(ByVal pSQL As String) As DrRecordSet
''레코드셋을 오픈
''쿼리문만 전달하면 레코드셋을 오푼하여 리턴한다.
'
'    Set RecordSetOpen = New DrRecordSet
'
'    RecordSetOpen.RsOpen , pSQL
'
''에러 체크
'    If RecordSetOpen.DBerror = True Then
'        MsgBox ZipFindDbConn.Errors.Item(1).Description
'        Exit Function
'    End If
'
'End Function

Public Function GetZipCd(ByVal Village As String) As String
'COM012
'우편번호 마스터에서 우편번호를 얻어온다.
    
    GetZipCd = "select * from " & T_COM012 & _
               " where village like '" & Village & "%'"
End Function

Public Function FormShow()
    Set MyForm = frmZipCdFind
        
    With frmZipCdFind
        Set .objMySQL = Me
        .Show vbModal
    End With
End Function

Public Property Let AddrNo(ByVal vData As String)
    mvarAddrNo = vData
End Property

Public Property Get AddrNo() As String
    AddrNo = mvarAddrNo
End Property

Public Property Let Village(ByVal vData As String)
    mvarVillage = vData
End Property

Public Property Get Village() As String
    Village = mvarVillage
End Property

Public Property Let District(ByVal vData As String)
    mvarDistrict = vData
End Property

Public Property Get District() As String
    District = mvarDistrict
End Property

Public Property Let Province(ByVal vData As String)
    mvarProvince = vData
End Property

Public Property Get Province() As String
    Province = mvarProvince
End Property

Public Property Let ZipCd(ByVal vData As String)
    mvarZipCd = vData
End Property

Public Property Get ZipCd() As String
    ZipCd = mvarZipCd
End Property


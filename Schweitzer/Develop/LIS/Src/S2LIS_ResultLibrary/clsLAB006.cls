VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lab006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsEquipMaster"
Attribute VB_Ext_KEY = "Member0" ,"clsEquipMaster"
Option Explicit
Private mCol As Collection

Public Function Add(EqpCd As String, EqpNm As String, _
   PurchDt As String, VandCd As String, _
   Remark As String, TempLow As Single, _
   TempHigh As Single, ModelNm As String, _
   EqpDiv As String, RecordCount As Long, _
   LAB006 As LAB006, Optional sKey As String) As clsEquipments
Dim objNewMember As clsEquipments
    '
   On Error GoTo ErrorAddCollection
   Set objNewMember = New clsEquipments
   '
   With objNewMember
      .EqpCd = EqpCd
      .EqpNm = EqpNm
      .PurchDt = PurchDt
      .VandCd = VandCd
      .Remark = Remark
      .TempLow = TempLow
      .TempHigh = TempHigh
      .ModelNm = ModelNm
      .RecordCount = RecordCount
      .EqpDiv = EqpDiv
      .Key = sKey                                  '** Collection Key **
   End With
   Set objNewMember.LAB006 = LAB006
   If Len(sKey) = 0 Then
       mCol.Add objNewMember
   Else
       mCol.Add objNewMember, sKey
   End If
   '
   Set Add = objNewMember
   Set objNewMember = Nothing
   Exit Function
ErrorAddCollection:
   OraErr.Reset
   glngErrorNo = Err.Number
   gstrErrorMsg = Err.Description
   Set Add = Nothing
   Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As clsEquipments
Attribute Item.VB_UserMemId = 0
   Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


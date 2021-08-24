VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LAB306"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(TransDt As String, EqpCd As String, TransTm As String, TransNo As String, _
SpcYy As String, SpcNo As String, WorkArea As String, AccDt As String, AccNo As String, StatFg As String, _
RecordCount As Long, LAB306 As LAB306, Optional sKey As String) As clsEquipTransfer
'
Dim objNewMember As clsEquipTransfer
   Set objNewMember = New clsEquipTransfer
    '
   With objNewMember
      .TransDt = TransDt
      .EqpCd = EqpCd
      .TransTm = TransTm
      .SpcYy = SpcYy
      .SpcNo = SpcNo
      .WorkArea = WorkArea
      .AccDt = AccDt
      .AccNo = AccNo
      .StatFg = StatFg
      .TransNo = TransNo
      
      .RecordCount = RecordCount
      .Key = sKey
   End With
   '
   Set objNewMember.LAB306 = LAB306
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
   '
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As clsEquipTransfer
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

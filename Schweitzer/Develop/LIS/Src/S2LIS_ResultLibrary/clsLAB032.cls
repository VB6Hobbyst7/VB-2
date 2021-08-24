VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LAB032"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Departments"
Attribute VB_Ext_KEY = "Member0" ,"Departments"
Option Explicit

Private mCol As Collection

Public Function Add(CdIndex As String, CdVal1 As String, _
Field1 As String, Field2 As String, Field3 As String, Field4 As String, Field5 As String, Text1 As String, Text2 As String, _
RecordCount As Long, LAB032 As LAB032, Optional sKey As String) As clsComcode032
'
Dim objNewMember As clsComcode032
   Set objNewMember = New clsComcode032
    '
   With objNewMember
      .CdIndex = CdIndex
      .CdVal1 = CdVal1
      .Field1 = Field1
      .Field2 = Field2
      .Field3 = Field3
      .Field4 = Field4
      .Field5 = Field5
      .Text1 = Text1
      .Text2 = Text2
      .RecordCount = RecordCount
      .Key = sKey
   End With
   '
   Set objNewMember.LAB032 = LAB032
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
   '
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As clsComcode032
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


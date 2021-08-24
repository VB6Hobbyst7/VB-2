VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_objParent As clsJSON
Private m_strkey    As String
Private m_strValue  As String
Private m_clcChild  As New Collection

Public Sub dolnitialize(ByRef v_strkey As String, v_strValue As String, Optional ByRef v_objParent As clsJSON)
    m_strkey = v_strkey
    m_strValue = v_strValue
    Set m_objParent = v_objParent
End Sub

Public Function getParent() As clsJSON
    Set getParent = m_objParent
End Function

Public Function getChild() As Collection
    Set getChild = m_clcChild
End Function

Public Function addChild(Optional ByRef v_strkey As String, Optional ByRef v_strValue As String) As clsJSON
    Dim objltem As New clsJSON
    
    Call objltem.dolnitialize(v_strkey, v_strValue, Me)
    
    If Len(v_strkey) = 0 Then
        Call m_clcChild.Add(objltem)
    Else
        Call m_clcChild.Add(objltem, v_strkey)
    End If
    Set addChild = objltem
    
End Function

Public Function getSize() As Long
    getSize = m_clcChild.Count
End Function

Public Function getItem(ByRef v_varlndex As Variant) As clsJSON
    On Error Resume Next
    Set getItem = m_clcChild(v_varlndex)
End Function

Public Function getKey() As String
    getKey = m_strkey
End Function

Public Function getValue() As String
    getValue = m_strValue
End Function





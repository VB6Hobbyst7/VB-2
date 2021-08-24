VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mvarSort As Boolean
Private mvarKey As String
Private mvarRecordCount As Long
Private mvarFieldName As String
Private mvarEOF As Boolean
Private mvarBOF As Boolean
Private m_Index As Long
Private mvarTagDic As New Scripting.Dictionary
Private mvarDictionary As New Scripting.Dictionary
Private mvarFieldAlias As New Scripting.Dictionary        'Field Position
Private mvarKeyAlias As New Scripting.Dictionary          'Key Field Position
Private mvarDataAlias As New Scripting.Dictionary         'Data Field Position
Private mvarTagAlias As New Scripting.Dictionary          'Tag Position

'/* Class Default Property */
Friend Property Get FieldName() As String
    '
    FieldName = mvarFieldName
    '
End Property

Public Property Get DataColCount() As String
    '
    DataColCount = mvarDataAlias.Count
    '
End Property

Public Property Get KeyColCount() As String
    '
    KeyColCount = mvarKeyAlias.Count
    '
End Property

Public Property Get TagColCount() As String
    '
    TagColCount = mvarTagAlias.Count
    '
End Property

Public Property Get ColCount() As String
    '
    ColCount = mvarFieldAlias.Count
    '
End Property

Public Property Get Key() As String
    '
    Key = mvarKey
    '
End Property

Public Property Get ItemData() As String
    '
    ItemData = mvarDictionary(mvarKey)

End Property

Public Property Get Sort() As Boolean
    '
    Sort = mvarSort
    '
End Property

Public Property Let Sort(ByVal vData As Boolean)
Dim blnLast As Boolean
    '
    blnLast = mvarSort
    mvarSort = vData
    If blnLast = False And mvarSort = True Then
        SortDictionary
    End If
    '
End Property

Public Function GetLine() As String
    '
    If mvarKey = "" Then Exit Function
    If mvarDictionary.Exists(mvarKey) = True Then
        GetLine = mvarKey & COL_DIV & mvarDictionary(mvarKey)
    End If
    '
End Function

Public Function GetLineClip() As String
    '
    If mvarKey = "" Then Exit Function
    If mvarDictionary.Exists(mvarKey) = True Then
        GetLineClip = mvarKey & vbTab & Replace(mvarDictionary(mvarKey), COL_DIV, vbTab)
    End If
    '
End Function

Public Sub KeyChange(ByVal ChangeKey As String)
Dim aryT As Variant
Dim ii As Long
    '
    If mvarDictionary.Exists(ChangeKey) = True Then
        If ChangeKey = mvarKey Then Exit Sub
        mvarKey = ChangeKey
        aryT = mvarDictionary.Keys
        For ii = LBound(aryT) To UBound(aryT)
            If mvarKey = CStr(aryT(ii)) Then
                m_Index = ii + 1
                Exit For
            End If
        Next ii
    End If
    '
End Sub

Public Property Let Fields(ByVal Field As String, ByVal vData As String)
Dim aryFld() As String
Dim newKey As String
    
    Field = UCase(Field)
    
    If mvarFieldAlias.Exists(Field) = True Then
        If mvarKeyAlias.Exists(Field) = True Then
            aryFld = Split(mvarKey, COL_DIV)
            aryFld(mvarKeyAlias(Field)) = vData
            newKey = Join(aryFld, COL_DIV)
            If newKey <> mvarKey Then
                '현재Dictionary키값과 새로운Dictionary키값이 틀리면
                '새로운키로 대체한다.
                mvarDictionary.Key(mvarKey) = newKey
                mvarKey = newKey
            End If
            SortDictionary
        Else
            aryFld = Split(mvarDictionary(mvarKey), COL_DIV)
            If mvarDictionary(mvarKey) = "" And mvarDataAlias.Count = 1 Then
                ReDim Preserve aryFld(mvarDataAlias.Count - 1)
            End If
            aryFld(mvarDataAlias(Field)) = vData
            mvarDictionary(mvarKey) = Join(aryFld, COL_DIV)
        End If
    End If

End Property

Public Property Get Fields(ByVal Field As String) As String
'
Dim aryFld() As String
Dim newKey As String
    '
    If mvarKey = "" Then
        Exit Function
    End If
    '
    Field = UCase(Field)
 'Debug.Print "mvarkey = " & mvarKey
   If mvarFieldAlias.Exists(Field) = True Then
        If mvarKeyAlias.Exists(Field) = True Then
            aryFld = Split(mvarKey, COL_DIV)
             Fields = aryFld(mvarKeyAlias(Field))
        Else
            aryFld = Split(mvarDictionary(mvarKey), COL_DIV)
            If mvarDictionary(mvarKey) = "" And mvarDataAlias.Count = 1 Then
                ReDim Preserve aryFld(mvarDataAlias.Count - 1)
            End If
            Fields = aryFld(mvarDataAlias(Field))
        End If
    End If
    '
End Property

Public Property Let Tags(ByVal Field As String, ByVal vData As String)
Dim aryTag() As String
Dim newKey As String
    
    Field = UCase(Field)
    If mvarTagAlias.Exists(Field) = True Then
        aryTag = Split(mvarTagDic(mvarKey), COL_DIV)
        aryTag(mvarTagAlias(Field)) = vData
        mvarTagDic(mvarKey) = Join(aryTag, COL_DIV)
    End If

End Property

Public Property Get Tags(ByVal Field As String) As String
'
Dim aryTag() As String
Dim newKey As String
    '
    If mvarKey = "" Then
        Exit Property
    End If
    
    '
    Field = UCase(Field)
 'Debug.Print "mvarkey = " & mvarKey
    If mvarTagAlias.Exists(Field) = True Then
        aryTag = Split(mvarTagDic(mvarKey), COL_DIV)
        Tags = aryTag(mvarTagAlias(Field))
    End If
    '
End Property

Public Property Get RecordCount() As Double
'
    RecordCount = mvarRecordCount
'
End Property

Public Property Get BOF() As Boolean
    BOF = mvarBOF
End Property

Public Property Get EOF() As Boolean
    EOF = mvarEOF
End Property

'/* Public Fucntion */
Public Function FieldInialize(ByVal KeyFieldName As String, _
                              ByVal DataFieldName As String, _
                              Optional ByVal TagFieldName As String = "") As Boolean
Dim strFields As String
    Dim aryFld() As String
    Dim ii As Double
    '
    If KeyFieldName = "" Then Exit Function
    '
    aryFld = Split(KeyFieldName, ",")
    For ii = LBound(aryFld) To UBound(aryFld)
        mvarKeyAlias.Add UCase(aryFld(ii)), CStr(ii)
    Next ii
    '
    strFields = KeyFieldName
    If DataFieldName <> "" Then
        aryFld = Split(DataFieldName, ",")
        For ii = LBound(aryFld) To UBound(aryFld)
            mvarDataAlias.Add UCase(aryFld(ii)), CStr(ii)
        Next ii
        strFields = strFields & "," & DataFieldName
    End If
    '
    aryFld = Split(strFields, ",")
    For ii = LBound(aryFld) To UBound(aryFld)
        mvarFieldAlias.Add UCase(aryFld(ii)), CStr(ii)
    Next ii
    mvarFieldName = strFields
    '
    '2001.2.2 추가
    If TagFieldName <> "" Then
        aryFld = Split(TagFieldName, ",")
        For ii = LBound(aryFld) To UBound(aryFld)
            mvarTagAlias.Add UCase(aryFld(ii)), CStr(ii)
        Next ii
    End If
    
    FieldInialize = True
    '
End Function

Public Sub MoveFirst()
'
   CopyDataToProperty 1
   '
End Sub

Public Sub MoveLast()
'
   CopyDataToProperty mvarDictionary.Count
   '
End Sub

Public Sub MoveNext()
'
   CopyDataToProperty (m_Index + 1)
'
End Sub

Public Sub MovePrevious()
'
    CopyDataToProperty (m_Index - 1)
'
End Sub

Public Function GetString() As String
    '
    GetString = DctToStr(mvarDictionary)
    '
End Function

Public Function GetClipText() As String
    '
    GetClipText = DctToClip(mvarDictionary)
    '
End Function

Public Function GetTagString() As String
    '
    GetTagString = DctToStr(mvarTagDic)
    '
End Function

Public Function AddNew(ByVal Key As String, ByVal Data As String, Optional ByVal Tags As String)
Dim aryTmp() As String
Dim aryKey() As String
Dim aryTag() As String
    '
    If InStr(1, Key, LINE_DIV) = 0 Then
        '1 Row Insert
        aryKey = Split(Key, COL_DIV)
        aryTmp = Split(Data, COL_DIV)
        aryTag = Split(Tags, COL_DIV)
        If UBound(aryKey) < (mvarKeyAlias.Count - 1) Then
            ReDim Preserve aryTmp(mvarKeyAlias.Count - 1)
        End If
        If UBound(aryTmp) < (mvarDataAlias.Count - 1) Then
            ReDim Preserve aryTmp(mvarDataAlias.Count - 1)
        End If
        If UBound(aryTag) < (mvarTagAlias.Count - 1) Then
            ReDim Preserve aryTag(mvarTagAlias.Count - 1)
        End If
        mvarDictionary.Add Join(aryKey, COL_DIV), Join(aryTmp, COL_DIV)
        mvarTagDic.Add Join(aryKey, COL_DIV), Join(aryTag, COL_DIV)
        SortDictionary
        mvarRecordCount = mvarDictionary.Count
        '
        mvarKey = Join(aryKey, COL_DIV)
    Else
        'Multi Row Insert
        Dim aryKeys() As String
        Dim strData As String
        Dim strTag As String
        Dim ii As Long
        '
        aryKeys = Split(Key, LINE_DIV)
        For ii = LBound(aryKeys) To UBound(aryKeys)
            '
            strData = medGetP(Data, ii + 1, LINE_DIV)
            strTag = medGetP(Tags, ii + 1, LINE_DIV)
            aryKey = Split(aryKeys(ii), COL_DIV)
            aryTmp = Split(strData, COL_DIV)
            aryTag = Split(strTag, COL_DIV)
            If mvarDictionary.Exists(Join(aryKey, COL_DIV)) = False Then
                If UBound(aryKey) < (mvarKeyAlias.Count - 1) Then
                    ReDim Preserve aryTmp(mvarKeyAlias.Count - 1)
                End If
                If UBound(aryTmp) < (mvarDataAlias.Count - 1) Then
                    ReDim Preserve aryTmp(mvarDataAlias.Count - 1)
                End If
                If UBound(aryTag) < (mvarTagAlias.Count - 1) Then
                    ReDim Preserve aryTag(mvarTagAlias.Count - 1)
                End If
                '
                mvarDictionary.Add Join(aryKey, COL_DIV), Join(aryTmp, COL_DIV)
                '
                mvarTagDic.Add Join(aryKey, COL_DIV), Join(aryTag, COL_DIV)
                SortDictionary
                mvarRecordCount = mvarDictionary.Count
                '
                mvarKey = Join(aryKey, COL_DIV)
                '
            End If
        Next ii
        '
    End If
    '
End Function

Private Sub SortDictionary()
Dim frmSort As New frmDLP001
Dim aryK() As Variant
Dim aryT() As Variant
Dim aryF() As Variant
Dim arySort() As String
Dim strKey As String
Dim dctTmp As New Scripting.Dictionary
Dim dctTag As New Scripting.Dictionary
Dim ii As Long
    '
    If mvarSort = False Then Exit Sub
    If mvarDictionary.Count > 1 Then
        dctTmp.CompareMode = BinaryCompare
        dctTag.CompareMode = BinaryCompare
        aryK = mvarDictionary.Keys
        aryT = mvarDictionary.Items
        aryF = mvarTagDic.Items '2001.2.2
        With frmDLP001
            medDataLoadLvw .lvwSort, LINE_DIV, Chr(16), Join(aryK, LINE_DIV)
            .lvwSort.SortKey = 0
            .lvwSort.SortOrder = lvwAscending
            .lvwSort.Sorted = True
            For ii = 1 To .lvwSort.ListItems.Count
               strKey = .lvwSort.ListItems(ii).Text
               If mvarDictionary.Exists(strKey) = True Then
                    dctTmp.Add strKey, mvarDictionary(strKey)
                    mvarDictionary.Remove (strKey)
                    dctTag.Add strKey, mvarTagDic(strKey)
                    mvarTagDic.Remove (strKey)
               End If
            Next ii
        End With
        frmDLP001.lvwSort.ListItems.Clear
        Set mvarDictionary = dctTmp
        Set mvarTagDic = dctTag
        Set dctTmp = Nothing
        Set dctTag = Nothing
        Set frmSort = Nothing
    End If
    '
End Sub

Public Function Exists(ByVal Key As String) As Boolean
    '
    Exists = mvarDictionary.Exists(Key)
    '
End Function

Public Sub Clear()
    '
    m_Index = 0
    mvarBOF = True
    mvarEOF = True
    mvarFieldName = ""
    mvarDictionary.RemoveAll
    mvarTagDic.RemoveAll
    mvarFieldAlias.RemoveAll
    mvarKeyAlias.RemoveAll
    mvarDataAlias.RemoveAll
    mvarTagAlias.RemoveAll
    mvarKey = ""
    mvarSort = True
    '
End Sub

Public Sub DeleteAll()
    '
    m_Index = 0
    mvarBOF = True
    mvarEOF = True
    mvarDictionary.RemoveAll
    mvarTagDic.RemoveAll
    mvarKey = ""
    '
End Sub

Public Sub Delete()
    '
    If mvarKey <> "" Then
        mvarDictionary.Remove (mvarKey)
        mvarTagDic.Remove (mvarKey)
        mvarRecordCount = mvarDictionary.Count
        mvarKey = ""
        If mvarRecordCount = 0 Then
            mvarBOF = True
            mvarEOF = True
            m_Index = 0
            mvarKey = ""
        Else
            m_Index = m_Index - 1
        End If
    End If
    '
End Sub

Private Sub Class_Initialize()
    '
    mvarBOF = False
    mvarEOF = False
    mvarSort = True
    '
    mvarDictionary.CompareMode = BinaryCompare
    mvarTagDic.CompareMode = BinaryCompare
    mvarFieldAlias.CompareMode = BinaryCompare
    mvarKeyAlias.CompareMode = BinaryCompare
    mvarDataAlias.CompareMode = BinaryCompare
    mvarTagAlias.CompareMode = BinaryCompare
    '
    m_Index = 0
    mvarRecordCount = 0
    '
End Sub

Private Sub Class_Terminate()
    '
    mvarRecordCount = 0
    m_Index = 0
    mvarDictionary.RemoveAll
    mvarTagDic.RemoveAll
    mvarFieldAlias.RemoveAll
    mvarKeyAlias.RemoveAll
    mvarDataAlias.RemoveAll
    mvarTagAlias.RemoveAll
    '
    Set mvarDictionary = Nothing
    Set mvarTagDic = Nothing
    Set mvarFieldAlias = Nothing
    Set mvarKeyAlias = Nothing
    Set mvarDataAlias = Nothing
    Set mvarTagAlias = Nothing
    '
End Sub

Private Function DctToStr(ByRef dctTmp As Scripting.Dictionary) As String
    
    Dim varKey As Variant
    Dim aryTmp() As String
    Dim blnFirst As Boolean
   
    'varkeyTmp = dctTmp.Keys
    If dctTmp.Count = 0 Then Exit Function
    For Each varKey In dctTmp.Keys
        If blnFirst = False Then
            ReDim aryTmp(0)
            blnFirst = True
        Else
            ReDim Preserve aryTmp(UBound(aryTmp) + 1)
        End If
        aryTmp(UBound(aryTmp)) = varKey & COL_DIV & dctTmp.Item(varKey)
    Next
   '
    DctToStr = Join(aryTmp, LINE_DIV)
   '
End Function

Private Function DctToClip(ByRef dctTmp As Scripting.Dictionary) As String
    
    Dim varKey As Variant
    Dim aryTmp() As String
    Dim blnFirst As Boolean
   
    'varkeyTmp = dctTmp.Keys
    If dctTmp.Count = 0 Then Exit Function
    For Each varKey In dctTmp.Keys
        If blnFirst = False Then
            ReDim aryTmp(0)
            blnFirst = True
        Else
            ReDim Preserve aryTmp(UBound(aryTmp) + 1)
        End If
        aryTmp(UBound(aryTmp)) = varKey & vbTab & Replace(dctTmp.Item(varKey), COL_DIV, vbTab)
    Next
   '
    DctToClip = Join(aryTmp, vbCrLf)
   '
End Function


Private Sub CopyDataToProperty(ByVal index As Long)
Dim aryT As Variant
   '
   mvarRecordCount = mvarDictionary.Count
   If (index < 1) Or (index > Me.RecordCount) Then
         If index < 1 Then
            m_Index = 0
         Else
            m_Index = Me.RecordCount + 1
         End If
         m_Index = index
         mvarKey = ""
    Else
        aryT = mvarDictionary.Keys
        m_Index = index
        mvarKey = aryT(m_Index - 1)
    End If
   '
    With mvarDictionary
       If .Count > 0 Then
          Select Case m_Index
             Case Is < 1
                mvarBOF = True
                mvarEOF = False
             Case Is > .Count
                mvarEOF = True
                mvarBOF = False
             Case Else
                mvarBOF = False
                mvarEOF = False
          End Select
       Else
          mvarEOF = True
          mvarBOF = True
       End If
    End With
   '
End Sub


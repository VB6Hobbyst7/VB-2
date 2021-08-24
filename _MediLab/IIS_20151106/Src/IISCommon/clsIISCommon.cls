VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISCommon"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISCommon.cls
'   �ۼ���  : ������
'   ��  ��  : ������Ʈ���� �������� ����ϴ� �Լ� Ŭ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'
Option Explicit

'## AlwaysOnEnum
Public Enum AlwaysOnEnum
    ccOff = 0
    ccOn = 1
End Enum

'-----------------------------------------------------------------------------'
'   ��� : �÷� ������
'-----------------------------------------------------------------------------'
Public Property Get DIV() As String
    DIV = Chr(19)
End Property

'-----------------------------------------------------------------------------'
'   ��� : ���� �ֻ����� ������� ǥ��
'   �μ� :
'       1.frmFrom   : �ش���
'       2.OnOff     : 0(����), 1(����)
'-----------------------------------------------------------------------------'
Public Sub mAlwaysOn(ByVal pForm As Object, ByVal pFlag As AlwaysOnEnum)
    Dim hWndMode As Integer

    hWndMode = Choose(pFlag + 1, HWND_NOTOPMOST, HWND_TOPMOST)
    SetWindowPos pForm.hwnd, hWndMode, 0, 0, 10, 10, SWP_NOMOVE Or SWP_NOSIZE
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ��ǻ�͸� �̸��� ��� �Լ�
'   ��ȯ : ��ǻ�� �̸�
'-----------------------------------------------------------------------------'
Public Function mGetComNm() As String
   Dim sBuffer$
   Dim nSize    As Long
   Dim rtn      As Long
   
   sBuffer = String(256, Chr(0))
   rtn = GetComputerName(sBuffer$, Len(sBuffer))
   mGetComNm = sBuffer
End Function

'-----------------------------------------------------------------------------'
'   ��� : INI���Ͽ��� �����б�
'   �μ� :
'       1.pFileNm   : INI���ϸ�(�������)
'       2.pSection  : INI������ Section
'       3.pKey      : INI������ Key
'       4.pDefault  : �ش� Ű���� ������ Default��
'-----------------------------------------------------------------------------'
Public Function mReadINI(ByVal pFileNm As String, ByVal pSection As String, _
                         ByVal pKey As String, Optional ByVal pDefault As String = "") As String
    Dim p As String
    
    p = Space$(256)
    GetPrivateProfileString pSection, pKey, pDefault, p, 256, pFileNm
    mReadINI = Mid(Trim(p), 1, Len(Trim(p)) - 1)
End Function

'-----------------------------------------------------------------------------'
'   ��� : INI���Ͽ��� �������� - 1.2.1: (2005-09-13)
'   �μ� :
'       1.pFileNm   : INI���ϸ�(�������)
'       2.pSection  : INI������ Section
'       3.pKey      : INI������ Key
'       4.pValue    : INI������ Key Value
'-----------------------------------------------------------------------------'
Public Sub mWriteINI(ByVal pFileNm, ByVal pSection As String, ByVal pKey As String, _
                     ByVal pValue As String)

    WritePrivateProfileString pSection, pKey, pValue, pFileNm
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڿ��� �����ڸ� �̿��� ������ ������ ��ġ�� ���ڿ��� ����
'   �μ� :
'       1.pText      : �����ڷ� ������ ���ڿ�
'       2.pPosiion   : ��ġ
'       3.pDelimiter : ������
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition �μ��� 1�� ��� For�� Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '�ش� �÷�
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function

'-----------------------------------------------------------------------------'
'   ��� : �޺��ڽ����� �ش繮�ڿ��� ã��1
'   �μ� :
'       1.pCombo    : ComboBox ��Ʈ��
'       2.pSerach   : ã�����ڿ�
'   ��ȯ : -1(���ڿ��� ã�� ���Ѱ��), �̿ܿ��� ã�� ���ڿ��� ListIndex
'-----------------------------------------------------------------------------'
Public Function mFindCombo(ByRef pCombo As Object, ByVal pSearch As String) As Long
    Dim i As Long
    
    If pSearch = "" Then
        mFindCombo = -1
        Exit Function
    End If
    
    With pCombo
        For i = 0 To .ListCount - 1
            If .List(i) Like (pSearch & "*") Then
                mFindCombo = i
                Exit Function
            End If
        Next i
    End With
    mFindCombo = -1
End Function

'-----------------------------------------------------------------------------'
'   ��� : �޺��ڽ����� �ش繮�ڿ��� ã��2
'   �μ� :
'       1.pCombo    : ComboBox ��Ʈ��
'       2.pSerach   : ã�����ڿ�
'   ��ȯ : -1(���ڿ��� ã�� ���Ѱ��), �̿ܿ��� ã�� ���ڿ��� ListIndex
'-----------------------------------------------------------------------------'
Public Function mFindComboX(ByRef pCombo As Object, ByVal pSearch As String) As Long
    Dim i As Long
    
    If pSearch = "" Then
        mFindComboX = -1
        Exit Function
    End If
    
    With pCombo
        For i = 0 To .ListCount - 1
            If Trim(mGetP(.List(i), 2, DIV)) = pSearch Then
                mFindComboX = i
                Exit Function
            End If
        Next i
    End With
    mFindComboX = -1
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� Spread Sheet�� Data�� Clear
'   �μ� :
'       - pSpread : ��� Spread Sheet
'       - pCol    : Column Header�� Clear�� ������
'       - pRow    : Row Header�� Clear�� ������
'-----------------------------------------------------------------------------'
Public Sub mTblClear(ByVal pSpread As Object, Optional ByVal pCol As Boolean, _
                     Optional ByVal pRow As Boolean)
Dim ii As Integer
    
    If pSpread.MaxRows > 0 Then
        pSpread.Col = 1
        pSpread.Col2 = pSpread.MaxCols
        pSpread.Row = 1
        pSpread.Row2 = pSpread.MaxRows
        pSpread.BlockMode = True
        pSpread.Action = 12    '## ActionClearText
        pSpread.BlockMode = False
    End If
    
    If IsMissing(pCol) Then pCol = False
    If IsMissing(pRow) Then pCol = False
    
    '## Column Header Clear
    If pCol = True Then
        For ii = 1 To pSpread.MaxCols
            pSpread.Row = 0
            pSpread.Col = ii
            pSpread.Value = " "
        Next ii
    End If
    '## Row Header Clear
    If pRow = True Then
        For ii = 1 To pSpread.MaxRows
            pSpread.Col = 0
            pSpread.Row = ii
            pSpread.Value = " "
        Next ii
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : SetParent API�� �����Լ�
'   �μ� :
'       - hWndChild  : �ڽĿ������� �ڵ�
'       - hWndParent : �θ�������� �ڵ�
'-----------------------------------------------------------------------------'
Public Function mSetParent(ByVal hWndChild As Long, ByVal hWndParent As Long)
    Call SetParent(hWndChild, hWndParent)
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش�ð����� �ð��� ����
'   �μ� :
'       - pInterval : �����ð�(1000���� 1��)
'-----------------------------------------------------------------------------'
Public Sub mSleep(ByVal pInterval As Long)
    Call Sleep(pInterval)
End Sub

'-----------------------------------------------------------------------------'
'   ��� : Workarea, AccDt, AccSeq�� �̿��� ������ȣ�� ����(Workarea-AccDt-AccSeq)
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'   ��ȯ : Workarea-AccDt-AccSeq ������ ������ȣ
'-----------------------------------------------------------------------------'
Public Function mGetAccNo(ByVal pWorkarea As String, ByVal pAccDt As String, ByVal pAccSeq As Long) As String
    mGetAccNo = pWorkarea & "-" & pAccDt & "-" & CStr(pAccSeq)
End Function

'-----------------------------------------------------------------------------'
'   ��� : ����ġ�� From, To ���� �̿��� From-To �������� ��ȯ
'   �μ� :
'       - pRefFrVal : Reference From Value
'       - pRefToVal : Reference From Value
'   ��ȯ : "From-To"������ ����ġ
'-----------------------------------------------------------------------------'
Public Function mGetRef(ByVal pRefFrVal As Single, ByVal pRefToVal As Single) As String
    If pRefFrVal = -1 And pRefToVal = -1 Then
        mGetRef = ""
    Else
        mGetRef = CStr(pRefFrVal) & " - " & CStr(pRefToVal)
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ֹι�ȣ ���ڸ��� �̿��� ���̰�� (By �赿��)
'   �μ� :
'       - pBirth : �ֹι�ȣ ���ڸ�
'   ��ȯ : ����
'-----------------------------------------------------------------------------'
Public Function mGetAge(ByVal pBirth As String) As String
    Dim strYear         As String
    Dim strMonth        As String
    Dim strDay          As String
    Dim strPtYear       As String   'ȯ�� ��������� �⵵
    Dim strPtMonth      As String   'ȯ�� ��������� ��
    Dim strPtDay        As String   'ȯ�� ��������� ����
    Dim lngYear         As Long
    Dim lngMonth        As Long
    Dim lngThisYear     As Long
    Dim lngThisMonth    As Long
    Dim lngThisDay      As Long
    
    If Trim$(pBirth) = "" Then
        mGetAge = ""
        Exit Function
    End If
    
    strPtYear = Mid$(pBirth, 1, 2)
    strPtMonth = Mid$(pBirth, 3, 2)
    strPtDay = Mid$(pBirth, 5, 2)
    
    lngThisYear = CLng(Format$(Now, "YY"))
    If (CLng(strPtYear) > lngThisYear) Then lngThisYear = lngThisYear + 100
    
    lngThisMonth = CLng(Format(Now, "MM"))
    If (CLng(strPtMonth) > lngThisMonth) Then lngThisMonth = lngThisMonth + 12
    
    strYear = CStr(lngThisYear - CLng(strPtYear))
    strMonth = CStr(lngThisMonth - CLng(strPtMonth))
    
    '## 1.0.2: (2004-12-14)
    '   - ������Ͽ� "700000"������ �����Ҷ� DateDiff�Լ����� ��Ÿ�ӿ��� �߻��� �����ϱ� ����
    '     ������ ��,���� ����!
    '## 1.1.5: (2005-03-25)
    '   - ��������� ��,�Ͽ� �ش簪�� ���� 13,32 �̻��� ��� ������ ��,���� "01"�� ����
    If strPtMonth = "00" Or CLng(strPtMonth) > 12 Then strPtMonth = "01"
    If strPtDay = "00" Or CLng(strPtDay) > 31 Then strPtDay = "01"
    strDay = DateDiff("d", strPtYear & "-" & strPtMonth & "-" & strPtDay, Now)
    
    If Trim$(strYear) <> "0" Then
        mGetAge = strYear & "Y"
    ElseIf Trim$(strMonth) > "1" Then
        mGetAge = strMonth & "M"
    Else
        mGetAge = strDay & "D"
    End If
End Function
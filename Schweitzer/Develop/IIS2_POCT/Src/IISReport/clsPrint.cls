VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISPrint.cls
'   �ۼ���  : �̻��
'   ��  ��  : Print Ŭ����
'   �ۼ���  : 2004-12-17
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

'## Orientation Enum
Public Enum OrientationEnum
    ccPortrait = 1      '�������
    ccLandscape = 2     '�������
End Enum

'## Width Align Enum
Public Enum WAlignEnum
    ccLeft = 1          '��������
    ccCenter = 2        '�������
    ccRight = 3         '����������
End Enum

'## Height Align Enum
Public Enum HAlignEnum
    ccTop = 1           '��������
    ccCenter = 2        '�������
    ccBottom = 3        '�Ʒ�����
End Enum

Private mFont       As String       '������Ʈ
Private mFontSize   As Long         '������Ʈũ��
Private mStartXPos  As Single       'Stat X Position
Private mCurYPos    As Single       'Current Y Position
Private mLastYPos   As Single       'Last Y Position
Private mScaleWidth As Single       '������ ���ο���
Private mSpace      As Single       '���λ����� ����
Private mPageCnt    As Long         'Page Count

Private mTitle      As String       '����Ʈ Ÿ��Ʋ
Private mTFont      As String       'Ÿ��Ʋ ��Ʈ
Private mTFontSize  As Long         'Ÿ��Ʋ ��Ʈ ������
Private mTWAlign    As WAlignEnum   'Ÿ��Ʋ�� ��������

Private mHeader     As String       '���� Ÿ��Ʋ
Private mColumns    As String       '�÷� Ÿ��Ʋ
Private mContent    As String       '����Ʈ ����

'-----------------------------------------------------------------------------'
'   ��� : ���
'   �μ� :
'       - pMode     : ��¹���
'       - pFont     : ��Ʈ
'       - pFontSize : ��Ʈũ��
'       - pStartX   : Start X Position
'       - pStartY   : Start Y Position
'       - pSpace    : �ٰ���
'-----------------------------------------------------------------------------'
Public Sub PrintDoc(ByVal pMode As OrientationEnum, ByVal pFont As String, _
                    ByVal pFontSize As Long, Optional ByVal pStartX As Single = 5, _
                    Optional ByVal pStartY As Single = 10, Optional ByVal pSpace As Single = 6)

    '## 1.������ �ʱ�ȭ
    mFont = pFont: mFontSize = pFontSize
    mStartXPos = pStartX: mCurYPos = pStartY: mSpace = pSpace
    
    With Printer
        .Font = mFont
        .FontSize = mFontSize
        .Orientation = pMode
        .ScaleMode = vbMillimeters
        
        mScaleWidth = .ScaleWidth
        mLastYPos = .ScaleHeight
    End With
    
    '## 2.Title �μ�
    Call PrintTitle
    
    '## 3.Sub Title �μ�
    
    '## 4.Column Header �μ�
    
    '## 5.Content �μ�
    
    Printer.EndDoc
End Sub

'-----------------------------------------------------------------------------'
'   ��� : Title �μ�
'-----------------------------------------------------------------------------'
Private Sub PrintTitle()
    '## Ÿ��Ʋ, ��Ʈ, ��������
End Sub

'-----------------------------------------------------------------------------'
'   ��� : Title ����
'   �μ� :
'       -
'-----------------------------------------------------------------------------'
Public Sub SetTitle(ByVal pTitle As String, Optional pFont As String = "", _
                    Optional pFontSize As Long = -1, _
                    Optional pWidthAlign As WAlignEnum = WAlignEnum.ccCenter)
    
    '## Ÿ��Ʋ ���� - �⺻ ����ü, 20Size, �������
    mTitle = pTitle
    mTFont = IIf(pFont = "", "����ü", pFont)
    mTFontSize = IIf(pFontSize = -1, 20, pFontSize)
    mTWAlign = pWidthAlign
End Sub

Private Sub RowPrint(ByVal pContent As String, ByVal pX As Single, ByVal pY As String, _
                     ByVal pWidth As Single, Optional ByVal pWAlign As WAlignEnum = WAlignEnum.ccLeft, _
                     Optional ByVal pHAlign As HAlignEnum = HAlignEnum.ccCenter, _
                     Optional ByVal pAddLine As Boolean = False)

    '## ��������
    Select Case pWAlign
        Case WAlignEnum.ccLeft   '��������
            Printer.CurrentX = pX + 0.5
        Case WAlignEnum.ccCenter '�������
            Printer.CurrentX = pX + (pWidth - Printer.TextWidth(pContent)) / 2
        Case WAlignEnum.ccRight  '����������
            Printer.CurrentX = pX + (pWidth - Printer.TextWidth(pContent)) - 0.5
    End Select
    
    '## ��������
    Select Case pHAlign
        Case HAlignEnum.ccTop
            
        Case HAlignEnum.ccCenter
        Case HAlignEnum.ccBottom
    End Select
End Sub
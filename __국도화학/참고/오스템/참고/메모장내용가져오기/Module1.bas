Attribute VB_Name = "Module1"
Option Explicit

 

 

Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

 

Public Const WM_SETTEXT = &HC

Public Const WM_GETTEXT = &HD

Public Const WM_GETTEXTLENGTH = &HE

 

 

 

 

Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long

 

    Dim length As Long

    Dim result As Long

    Dim strtmp As String

 

 

    length = SendMessage(hwnd, WM_GETTEXTLENGTH, ByVal 0, ByVal 0) + 1

    strtmp = Space(length - 1)

    result = SendMessage(hwnd, WM_GETTEXT, ByVal length, ByVal strtmp)

    result = SendMessage(Form1.Picture1.hwnd, WM_GETTEXT, ByVal length, ByVal strtmp)
 

    



    Debug.Print strtmp

    

    EnumChildProc = 1

 

End Function

'''
'''
'''
''''Clipboard ��ü
''''�ý��� Ŭ�����忡 �׼����մϴ�.
''''����
''''Clipboard
''''����
''''Clipboard ��ü�� Ŭ�����忡 �ִ� ���ڿ� �׷����� ������ �� ����մϴ�. �� ��ü�� ����ϸ� ���� ���α׷��� ���ڿ� �׷����� ����, �ڸ���, �ٿ��ֱ� ���� �۾��� ������ �� �ֽ��ϴ�. Clipboard ��ü�� ��Ҹ� �����ϱ� ���� Clipboard.Clear ���� Clear �޼��带 �����Ͽ� ������ ������ �մϴ�.
''''Clipboard ��ü�� Windows ���� ���α׷��� �����ǵ��� �����ϸ� ����ڰ� �ٸ� ���� ���α׷����� ��ȯ�� ������ ������ ����˴ϴ�.
''''������ �����Ͱ� �ٸ� ������ ��� Clipboard ��ü�� ���� ���� �����͸� ������ �� �ֽ��ϴ�. ���� ��� vbCFDIB �������� SetData �޼��带 ����Ͽ� Clipboard�� ��Ʈ���� ���� �� �ְ� vbCFText �������� SetText �޼��带 ����Ͽ� Clipboard�� �ؽ�Ʈ�� ���� �� �ֽ��ϴ�. GetText �޼��带 ����Ͽ� �ؽ�Ʈ�� �����ϰ� GetData �޼��带 ����Ͽ� �׷����� ������ �� �ֽ��ϴ�.������ ������ �ٸ� ������ ������ �ڵ峪 �޴� ��ɾ ���� Clipboard ���� ���̸� Clipboard�� �����ʹ� �ҽǵ˴ϴ�.
''''
'''
'''    Clipboard.Clear
'''    Clipboard.GetFormat (vbCFText)
'''    Clipboard.SetText "�޸��ʵ� ����", vbCFText
'''
'''
'''
'''    '�޸��忡 ��Ŀ�� �ְ�...
'''
'''    Clipboard.GetText (vbCFText)
'''
'''
'''
'''    '�Ǵ� SendKeys "^{V}"
'''
'''

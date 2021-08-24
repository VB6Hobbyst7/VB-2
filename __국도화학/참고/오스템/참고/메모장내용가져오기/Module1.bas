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
''''Clipboard 개체
''''시스템 클립보드에 액세스합니다.
''''구문
''''Clipboard
''''참고
''''Clipboard 개체는 클립보드에 있는 문자와 그래픽을 조작할 때 사용합니다. 이 개체를 사용하면 응용 프로그램에 문자와 그래픽의 복사, 자르기, 붙여넣기 등의 작업을 수행할 수 있습니다. Clipboard 개체에 요소를 복사하기 전에 Clipboard.Clear 같은 Clear 메서드를 수행하여 내용을 지워야 합니다.
''''Clipboard 개체를 Windows 응용 프로그램과 공유되도록 주의하면 사용자가 다른 응용 프로그램으로 전환할 때마다 내용이 변경됩니다.
''''각각의 데이터가 다른 유형인 경우 Clipboard 개체는 여러 개의 데이터를 포함할 수 있습니다. 예를 들어 vbCFDIB 유형으로 SetData 메서드를 사용하여 Clipboard에 비트맵을 넣을 수 있고 vbCFText 유형으로 SetText 메서드를 사용하여 Clipboard에 텍스트를 넣을 수 있습니다. GetText 메서드를 사용하여 텍스트를 복구하고 GetData 메서드를 사용하여 그래픽을 복구할 수 있습니다.동일한 유형의 다른 데이터 집합이 코드나 메뉴 명령어를 통해 Clipboard 위에 놓이면 Clipboard의 데이터는 소실됩니다.
''''
'''
'''    Clipboard.Clear
'''    Clipboard.GetFormat (vbCFText)
'''    Clipboard.SetText "메모필드 내용", vbCFText
'''
'''
'''
'''    '메모장에 포커스 주고...
'''
'''    Clipboard.GetText (vbCFText)
'''
'''
'''
'''    '또는 SendKeys "^{V}"
'''
'''

Attribute VB_Name = "UrlEncoder"
Option Explicit

Private Const API_NULL As Long = 0

Private Const S_OK As Long = 0
Private Const E_POINTER As Long = &H80004003

Private Const CP_UTF8 As Long = 65001

Private Enum UrlEscapeFlags
    URL_DONT_ESCAPE_EXTRA_INFO = &H2000000
    URL_ESCAPE_SPACES_ONLY = &H4000000
    URL_ESCAPE_PERCENT = &H1000&
    URL_ESCAPE_SEGMENT_ONLY = &H2000&
    URL_ESCAPE_AS_UTF8 = &H40000    'Win7 or later.
End Enum

Private Enum UrlParts
    URL_PART_NONE = 0
    URL_PART_SCHEME = 1
    URL_PART_HOSTNAME
    URL_PART_USERNAME
    URL_PART_PASSWORD
    URL_PART_PORT
    URL_PART_QUERY
End Enum

Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeW" ( _
    ByVal pszURL As Long, _
    ByVal pszEscaped As Long, _
    ByRef cchEscaped As Long, _
    ByVal dwFlags As UrlEscapeFlags) As Long

Private Declare Function UrlGetPart Lib "shlwapi" Alias "UrlGetPartW" ( _
    ByVal pszIn As Long, _
    ByVal pszOut As Long, _
    ByRef cchOut As Long, _
    ByVal dwPart As UrlParts, _
    ByVal dwFlags As Long) As Long

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long
    
Private Function Utf8Escape(ByVal Unicode As String) As String
    Dim I As Integer
    Dim UnicodeCh As String
    Dim Utf8Ch() As Byte
    Dim cchMultiByte As Long
    Dim J As Long
    Dim Escaped As String
    
    ReDim Utf8Ch(6) 'Up to 6 bytes plus 1 for NUL byte.
    For I = 1 To Len(Unicode)
        UnicodeCh = Mid$(Unicode, I, 1)
        cchMultiByte = 7
        cchMultiByte = WideCharToMultiByte(CP_UTF8, _
                                           0, _
                                           StrPtr(UnicodeCh), _
                                           1, _
                                           VarPtr(Utf8Ch(0)), _
                                           cchMultiByte, _
                                           0, _
                                           0)
        If cchMultiByte = 1 Then
            Utf8Escape = Utf8Escape & UnicodeCh
        Else
            Escaped = ""
            For J = 0 To cchMultiByte - 1
                Escaped = Escaped & "%" & Right$("0" & Hex$(Utf8Ch(J)), 2)
            Next
            Utf8Escape = Utf8Escape & Escaped
        End If
    Next
End Function

Public Function UrlQueryStringEncode(ByVal Url As String) As String
    'To be entirely correct, a first pass should be done on Url calling
    'UrlEscape() with Flags = URL_ESCAPE_PERCENT Or URL_ESCAPE_AS_UTF8
    'in order to escape the part before the query string.
    '
    'That would also have to use the pre-Win7 hack of calling the
    'Utf8Escape() function above.
    '
    'However for many purposes what we have here should be good enough,
    'and as far as I know Utf8Escape() does an accurate job on pre-Win7
    'systems with the query string part.
    Dim Parts As UrlParts
    Dim cchOut As Long
    Dim Query As String
    Dim HResult As Long
    Dim Flags As UrlEscapeFlags
    Dim EscapedQuery As String
    Dim cchEscaped As Long
    Dim OSVersion As Long
    Dim BeforeWin7 As Boolean
    
    Parts = URL_PART_QUERY
    cchOut = 1
    Query = ""
    HResult = UrlGetPart(StrPtr(Url), StrPtr(Query), cchOut, Parts, 0)
    If HResult = E_POINTER Then
        Query = Space$(cchOut - 1)
        HResult = UrlGetPart(StrPtr(Url), StrPtr(Query), cchOut, Parts, 0)
        If HResult = S_OK Then
            Flags = URL_ESCAPE_PERCENT Or URL_ESCAPE_SEGMENT_ONLY Or URL_ESCAPE_AS_UTF8
            cchEscaped = 1
            EscapedQuery = ""
            HResult = UrlEscape(StrPtr(Query), StrPtr(EscapedQuery), cchEscaped, Flags)
            If HResult = E_POINTER Then
                EscapedQuery = Space$(cchEscaped - 1)
                HResult = UrlEscape(StrPtr(Query), StrPtr(EscapedQuery), cchEscaped, Flags)
                If HResult = S_OK Then
                    OSVersion = GetVersion()
                    BeforeWin7 = CCur(OSVersion And &HFF&) _
                               + CCur((OSVersion \ &H100&) And &HFF&) / 10000@ < 6.0001@
                    If BeforeWin7 Then EscapedQuery = Utf8Escape(EscapedQuery)
                    UrlQueryStringEncode = LEFT$(Url, Len(Url) - Len(Query)) & EscapedQuery
                Else
                    Err.Raise HResult, "UrlEscape", "System error " & Hex$(HResult)
                End If
            Else
                Err.Raise HResult, "UrlEscape", "System error " & Hex$(HResult)
            End If
        Else
            Err.Raise HResult, "UrlGetPart", "System error " & Hex$(HResult)
        End If
    Else
        Err.Raise HResult, "UrlGetPart", "System error " & Hex$(HResult)
    End If
End Function




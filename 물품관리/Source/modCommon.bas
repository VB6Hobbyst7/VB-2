Attribute VB_Name = "modCommon"
Option Explicit

Public gDbCn As ADODB.Connection, gOraCn As ADODB.Connection, gSql As String, cDb As clsDbConnect
Public gUserId As String

Public Const gReqStatusStr As String = "��û|����|����|�԰�"                                    ' ��û�� ����
Public Const gOperFlagStr As String = "����|�ϴ���|�ִ���|������|�б����|�ݱ����|�����"      ' ��� �ֱ�
Public Const gOutFlagStr As String = "�Ϲ�|����|�"                                           ' ����ڷ� ����
Public Const gEndFlagStr As String = "���|����"                                                ' �����˻� ��������
Public Const gReagentFlagStr As String = "�Ϲ�|�þ�"                                            ' ��ǰ�з� ����
Public Const gBuyTypeStr As String = "�Ϲݰ��|���ΰ��"                                        ' ��ǰ���� ����
Public Const gRmdTypeStr As String = "������|��������"                                        ' ��ǰ������ ����
Public Const gUserLevelStr As String = "�Ϲ�|����"                                              ' ����ڱ���
Public Const gOrderTypeStr As String = "�Ϲ�|��û"                                              ' ���ֱ���
Public Const gOrderStatStr As String = "���|�԰�|�Ϸ�"                                         ' ���ֻ���
Public Const gBuyIoTypeStr As String = "�Ϲ�|����"                                              ' ���ű���

Public gReqStatus() As String, gOperFlag() As String, gOutFlag() As String, gEndFlag() As String, gReagentFlag() As String
Public gBuyType() As String, gRmdType() As String, gUserLevel() As String, gOrderType() As String, gOrderStat() As String, gBuyIoType() As String

Public Const gAllData As Byte = 9, gDelNo As Byte = 0, gDelYes As Byte = 1

Public Const gEndWrite As Byte = 0, gEndDay As Byte = 1
Public Const gOutNormal As Byte = 0, gOutDayEnd As Byte = 1, gOutMachOper As Byte = 2
Public Const gReqStatWrt As Byte = 0, gReqStatHold As Byte = 1, gReqStatOrder As Byte = 2, gReqStatBuy As Byte = 3
Public Const gOrderNormal As Byte = 0, gOrderReq As Byte = 1
Public Const gOrderStsWrt As Byte = 0, gOrderStsBuy As Byte = 1, gOrderStsEnd As Byte = 2
Public Const gBuyIoNormal As Byte = 0, gBuyIoOrder As Byte = 1

'----------- API Popup ----------------------
Public Const MF_CHECKED = &H8&
Public Const MF_APPEND = &H100&
Public Const TPM_LEFTALIGN = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_GRAYED = &H1&
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const TPM_RETURNCMD = &H100&
Public Const TPM_RIGHTBUTTON = &H2&
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal HWnd As Long, ByVal lptpm As Any) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public hMenu As Long

Sub Main()

    If App.PrevInstance Then
        MsgBox App.Title & " �� ���α׷��� �̹� �������Դϴ�.!", vbInformation
        End
    End If

    frm����ȭ��.Show
    
    Set cDb = New clsDbConnect
    Do While Not cDb.cfConnect
        frm�����ͺ��̽�.Show vbModal

        If MsgBox("�����ͺ��̽��� ������ �� �����ϴ�. �ٽ� �õ��Ͻðڽ��ϱ� ?", vbYesNo + vbQuestion) <> vbYes Then
            End
        End If
    Loop

    gReqStatus = Split(gReqStatusStr, "|")
    gOperFlag = Split(gOperFlagStr, "|")
    gOutFlag = Split(gOutFlagStr, "|")
    gEndFlag = Split(gEndFlagStr, "|")
    gReagentFlag = Split(gReagentFlagStr, "|")
    gBuyType = Split(gBuyTypeStr, "|")
    gRmdType = Split(gRmdTypeStr, "|")
    gUserLevel = Split(gUserLevelStr, "|")
    gOrderType = Split(gOrderTypeStr, "|")
    gOrderStat = Split(gOrderStatStr, "|")
    gBuyIoType = Split(gBuyIoTypeStr, "|")

    frm����ȭ��.Caption = frm����ȭ��.Caption & " (ver " & App.Major & "." & App.Minor & "." & App.Revision & " / " & gDbCn.Properties("Server Name").Value & ")"
    frm����ȭ��.Enabled = False
    
    Call gsRegisterApply
    
    frm�α���.Show vbModal

End Sub

Public Sub gsRegisterApply()
Dim cReg As clsRegister

    Set cReg = New clsRegister

    frm����ȭ��.stsBar.Panels(1).Text = cReg.username

End Sub

Public Function gfSystemDate() As Date
' �����ý����� ��¥/�ð�
Dim sDate As Date

    gSql = "select GETDATE() as sysdt"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            sDate = .Fields("sysdt").Value
            .Close
        Else
            sDate = Now
        End If
    End With
    
    gfSystemDate = sDate
'    gfSystemDate = "2012-05-15 " & Format(Now, "Hh:Nn:Ss")

End Function


Public Function HLeft(ByVal vString As String, ByVal vLen As Long) As String
' �ѱ����Ե� ���忡�� Left�Լ� ���
    HLeft = StrConv(LeftB(StrConv(vString, vbFromUnicode), vLen), vbUnicode)

End Function

Public Function HRight(ByVal vString As String, ByVal vLen As Long) As String
' �ѱ����Ե� ���忡�� right�Լ� ���

    HRight = StrConv(RightB(StrConv(vString, vbFromUnicode), vLen), vbUnicode)

End Function

Public Function HMid(ByVal vString As String, ByVal vLenF As Long, ByVal vLenT As Long) As String
' �ѱ����Ե� ���忡�� mid�Լ� ���

    HMid = StrConv(MidB(StrConv(vString, vbFromUnicode), vLenF, vLenT), vbUnicode)

End Function

Public Function HLen(ByVal vString As String) As Long
' �ѱ����Ե� ���忡�� len�Լ� ���

    HLen = LenB(StrConv(vString, vbFromUnicode))

End Function

Public Sub gsEnterEsc_KeyPress(ByVal brForm As Object, ByVal brKeyAscii As Integer, ByVal brCount As Integer)
Dim NextTabIndex As Integer, I As Integer
    
    On Error Resume Next
    If brKeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        GoTo Rtn_Exit
    End If

    If brKeyAscii = Asc("'") Then
        MsgBox "���α׷� Ư���� ����� �� ���� ���� �Դϴ�.", vbCritical, "���Ұ�����"
        SendKeys "{BS}"
        GoTo Rtn_Exit
    End If
    
    If Left$(brForm.ActiveControl.Name, 3) = "txt" Then
        If HLen(brForm.ActiveControl.Text) > brForm.ActiveControl.MaxLength And brForm.ActiveControl.MaxLength > 0 Then
            SendKeys "{BS}"
        End If
    End If
    
Rtn_Exit:
End Sub

Public Sub gsSpreadClear(ByVal brSpread As Object, Optional ByVal brRow As Long = 1000, Optional ByVal brColor As Boolean = False, Optional ByVal brHeight As Integer = 0, _
                         Optional ByVal brRowAdd As Boolean = False)
' �������� Clear
    
   On Error GoTo gsSpreadClear_ERROR

    With brSpread
        .UserResize = UserResizeNone
        
        .MaxRows = brRow
        .RetainSelBlock = True
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
        If brRowAdd = False Then
            .Action = ActionClearText
        End If
        If brHeight = 0 Then
            .RowHeight(-1) = .FontSize * 1.5
        Else
            .RowHeight(-1) = brHeight
        End If
        .BlockMode = False
        
        If brColor Then
            .SetOddEvenRowColor vbWhite, vbBlack, &HF1F1F1, vbBlack
            .SetCellBorder 1, 1, .MaxCols, .MaxRows, 13, &HDEDEDE, CellBorderStyleSolid
        End If
    End With

   Exit Sub
gsSpreadClear_ERROR:
   MsgBox Err.Numbe, vbCritical

End Sub

Public Sub gsFieldClear(ByVal brForm As Object)
Dim ii, sName As String
' Control Field Clear

    On Error Resume Next
    For ii = 0 To brForm.Count - 1
        sName = Left$(brForm.Controls(ii).Name, 3)
        If brForm.Controls(ii).Enabled = True Then
            Select Case LCase(sName)
                Case "num":     brForm.Controls(ii).Text = ""
                Case "txt":     brForm.Controls(ii).Text = ""
                Case "cbo":     brForm.Controls(ii).Text = ""
                Case "lbl":     brForm.Controls(ii).Caption = ""
                Case "gtm":     brForm.Controls(ii).Value = 0
                Case "dtp"
                        brForm.Controls(ii).Value = Now
                        If brForm.Controls(ii).CheckBox Then brForm.Controls(ii).Value = ""
                Case "chk":     brForm.Controls(ii).Value = 0
            End Select
        End If
    Next ii

End Sub

Public Sub gsSetStkTree(ByVal brTree As SSTree, Optional ByVal brType As Byte = gAllData)
Dim sKeyStr As String, sChildStr As String

    brTree.ImageList = frm����ȭ��.imgTree
    
    brTree.Nodes.Clear
    gSql = "select * from mstSTKG "
    If brType <> gAllData Then
        gSql = gSql & " where reagentfg = 1 "
    End If
    gSql = gSql & " order by kindnm "
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF)
                sKeyStr = "" & .Fields("kindcd").Value
                brTree.Nodes.Add , , sKeyStr, .Fields("kindnm").Value, "close", "open", .Fields("kindcd").Value
                
                gSql = "select stkcd, stknm, stkspec from mstSTK where kindcd = " & .Fields("kindcd").Value & " and delfg = " & gDelNo & " order by stknm"
                With cDb.cfRecordSet(gSql)
                    If .State = adStateOpen Then
                        While (Not .EOF)
                            sChildStr = "" & .Fields("stknm").Value
                            If Len("" & .Fields("stkspec").Value) > 0 Then
                                sChildStr = sChildStr & "(" & .Fields("stkspec").Value & ")"
                            End If
                            brTree.Nodes.Add sKeyStr, tvwChild, "" & .Fields("stkcd").Value, sChildStr, "", "choice", .Fields("stkcd").Value
                            
                            .MoveNext
                        Wend
                        .Close
                    End If
                End With
                
                .MoveNext
            Wend
            .Close
        End If
    End With

End Sub

Public Sub gsSetTestTree(ByVal brTree As SSTree, Optional ByVal brType As Byte = gAllData)
Dim sKeyStr As String, sChildStr As String, sPart As String

    brTree.ImageList = frm����ȭ��.imgTree
    
    brTree.Nodes.Clear
    If cDb.cfOraConnect Then
        gSql = "select itemcode, lpartcode, itemhnm from TWMED_ITEM where visible = 1 order by lpartcode,itemhnm"
        With cDb.cfOraRecordSet(gSql)
            If .State = adStateOpen Then
                While (Not .EOF)
                    If sPart <> ("" & .Fields("lpartcode").Value) Then
                        sPart = "" & .Fields("lpartcode").Value
                        If Len(sPart) = 0 Then
                            sKeyStr = "Etc"
                        Else
                            sKeyStr = "" & .Fields("lpartcode").Value
                        End If
                        
                        brTree.Nodes.Add , , sKeyStr, sKeyStr, "close", "open", sKeyStr
                    End If
                                
                    sChildStr = "[" & .Fields("itemcode").Value & "] " & .Fields("itemhnm").Value
                    brTree.Nodes.Add sKeyStr, tvwChild, "" & .Fields("itemcode").Value, sChildStr, "", "choice", .Fields("itemcode").Value
                    
                    .MoveNext
                Wend
                .Close
            End If
        End With
    End If
    
End Sub


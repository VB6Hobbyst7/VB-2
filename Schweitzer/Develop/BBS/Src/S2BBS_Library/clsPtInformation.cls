VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPtInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarPtId As String
Private mvarPtNm As String
Private mvarDoB As String
Private mvarBedIndt As String
Private mvarDeptCd As String
Private mvarDeptNm As String
Private mvarMajdoct As String
Private mvarDoctNm As String
Private mvarWardID As String
Private mvarSex As String
Private mvarAge As String
Private mvarPtDiv As String
Private mvarHosilID As String
Private mvarBedID As String

Public Property Get BedID() As String
    BedID = mvarBedID
End Property
Public Property Get HosilID() As String
    HosilID = mvarHosilID
End Property
Public Property Get PtDiv() As String
    PtDiv = mvarPtDiv
End Property

Public Property Get PtNm() As String
    PtNm = mvarPtNm
End Property

Public Property Get PtId() As String
    PtId = mvarPtId
End Property

Public Property Get dob() As String
    dob = mvarDoB
End Property

Public Property Get BedIndt() As String
    BedIndt = mvarBedIndt
End Property

Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property

Public Property Get DeptNm() As String
    DeptNm = mvarDeptNm
End Property

Public Property Get MajDoct() As String
    MajDoct = mvarMajdoct
End Property

Public Property Get DoctNm() As String
    DoctNm = mvarDoctNm
End Property
Public Property Get wardid() As String
    wardid = mvarWardID
End Property
Public Property Get sex() As String
    sex = mvarSex
End Property
Public Property Get Age() As String
    Age = mvarAge
End Property

'Public Sub setDbConn(ByRef pDbConn As DrDatabase)
'    Set DbConn = pDbConn
'End Sub

Public Function GetPtInfo(ByVal Searchkey As String, ByVal blnSearch As Boolean, ByVal orddt As String, Optional Bussdiv As String = "", Optional DoneFg As String = "") As String
    '���ȯ���� ����(donefg=""),ä��������� ����(donefg="0")
    
    Dim stsORDER   As String
    Dim stsCOLLECT As String
    
    If TRANS_REQUIRE_USED Then
        stsORDER = BBSOrdStatus.stsORDER
        stsCOLLECT = BBSOrdStatus.stsCOLLECT
    Else
        stsORDER = BBSOrderStatus.stsORDER
        stsCOLLECT = BBSOrderStatus.stsCOLLECT
    End If
    
    If DoneFg = "" Then                 'ó���Ͽ��� ���ȯ�� ��ȸ
        GetPtInfo = " SELECT " & _
                                 F_PTID & " as ptid," & _
                                 F_PTNM & " as ptnm," & _
                                 F_SSN & " as ssn" & _
                    " FROM " & _
                                 T_HIS001 & " a" & _
                    " WHERE "
    Else
        If Bussdiv = "" Then            '��ȣ��ä������ ����ȯ�ڸ� ��ȸ
            GetPtInfo = " SELECT distinct a." & F_PTID & " as ptid,a." & F_PTNM & " as ptnm," & F_SSN2("a") & " as ssn " & _
                        " FROM " & _
                          T_HIS001 & " a," & T_LAB101 & " b" & _
                        " WHERE a." & F_PTID & " = b.ptid " & _
                        " AND bussdiv in(" & _
                                        DBV("bussdiv", BBSBUSSDIV.stsNotBed, 1) & _
                                        DBV("bussdiv", BBSBUSSDIV.stsBed) & ")" & _
                        " AND " & DBW("b.donefg", stsORDER, 2) & _
                        " AND " & DBW("b.orddt", orddt, 2) & _
                        " AND " & DBW("b.orddiv", C_WORKAREA, 2) & " AND "
        ElseIf Bussdiv = "�ܷ�ä��" Then    '�ܷ�����(ó�����,ä������ �Ѵ� ��ȸ)
            GetPtInfo = " SELECT distinct a." & F_PTID & " as ptid,a." & F_PTNM & " as ptnm," & F_SSN2("a") & " as ssn " & _
                        " FROM " & _
                          T_HIS001 & " a," & T_LAB101 & " b" & _
                        " WHERE a." & F_PTID & " = b.ptid " & _
                        " AND bussdiv in (" & _
                                        DBV("bussdiv", BBSBUSSDIV.stsNotBed, 1) & _
                                        DBV("bussdiv", BBSBUSSDIV.stsEr) & ")" & _
                        " AND b.donefg in(" & _
                                      DBV("bussdiv", stsORDER, 1) & _
                                      DBV("bussdiv", stsCOLLECT) & ") " & _
                        " AND " & DBW("b.orddt", orddt, 2) & _
                        " AND " & DBW("b.orddiv", C_WORKAREA, 2) & " AND"
        Else                                '�ܷ�ä��(ó����¸� ��ȸ)..
            GetPtInfo = " SELECT distinct a." & F_PTID & " as ptid,a." & F_PTNM & " as ptnm," & F_SSN2("a") & " as ssn " & _
                        " FROM " & T_HIS001 & " a," & T_LAB101 & " b" & _
                        " WHERE a." & F_PTID & " = b.ptid " & _
                        " AND bussdiv in(" & _
                                        DBV("bussdiv", BBSBUSSDIV.stsNotBed, 1) & _
                                        DBV("bussdiv", BBSBUSSDIV.stsEr) & ")" & _
                        " AND " & DBW("b.donefg", stsORDER, 2) & _
                        " AND " & DBW("b.orddt", orddt, 2) & "" & _
                        " AND " & DBW("b.orddiv", C_WORKAREA, 2) & " AND"
        End If
    End If
   'ȯ��ID�� ���翩�θ� Ȯ���Ѵ�.
    If blnSearch = True Then
        GetPtInfo = GetPtInfo & " a." & F_PTID & " between " & DBV(F_PTID, Format(Searchkey, "00000000")) & " AND " & DBV(F_PTID, Format((Searchkey + 1000), "00000000")) & " "
    Else
        GetPtInfo = GetPtInfo & " a." & F_PTNM & " like " & "'" & Searchkey & "%" & "'"
    End If
             
End Function

Public Function Get_Ptid(ByVal PtId As String) As String
    Get_Ptid = " SELECT " & F_PTID & _
               " FROM   " & T_HIS001 & _
               " WHERE  " & _
                            F_PTID & DBV(F_PTID, PtId, 2)
End Function

Public Function BedPt_Chk(ByVal strPtid As String, ByVal strDate As String) As String
    '�Կ�ȯ�� ��ȸ�� �Ѵ�.(���� �Կ��ߴ� ȯ�ڵ� ������)
    Dim sSql As String
    Dim Rs   As Recordset
    
    Set Rs = New Recordset
    

    '
    '��������Ϳ��� ����ϴ� ȯ��ID Alias�� F_INPTID�̴�.....
    '
    sSql = "SELECT " & F_INPTID & " as ptid FROM " & T_HIS002 & _
           " WHERE " & DBW(F_INPTID, strPtid, 2)

    Call Rs.Open(sSql, DBConn)
    
'    If Rs.DBerror = True Then
'        DBCONN.DisplayErrors
'        Set Rs = Nothing
'        Exit Function
'    End If
    If Not Rs.EOF = True Then
        '�Կ�ȯ���̴�.(����� ȯ���ϼ����ִ�.)
        
        BedPt_Info strPtid, strDate
    Else
        '�Կ�ȯ�ھƴϴ�.
        
        GetBedPtInfo strPtid, False
    End If
    Set Rs = Nothing

End Function

Private Function BedPt_Info(ByVal strPtid As String, ByVal strBedinDt As String)
    '�Կ�ȯ�ڸ� ��ȸ(���� ���� �Կ��� ȯ�ڸ� ��ȸ)
    Dim Rs      As Recordset
    Dim sSql    As String
    
    sSql = " SELECT " & F_INPTID & " as ptid FROM " & T_HIS002 & _
             " WHERE " & DBW(F_INPTID, strPtid, 2) & " " & _
             " AND " & F_BEDINDT & " <=to_date(" & strBedinDt & ",'YYYYMMDD')"
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
'    If Rs.DBerror = True Then
'        dbconn.DisplayErrors
'        Set Rs = Nothing
'        Exit Function
'    End If
        
    If Not Rs.EOF = True Then
        '�Կ�ȯ��(���� ���絵 �Կ��ߴ����� ������)
        BedOut_chk strPtid, strBedinDt
    End If
    
    Set Rs = Nothing
End Function

Private Function BedOut_chk(ByVal strPtid As String, ByVal BadOut As String)
    Dim Rs   As Recordset
    Dim sSql As String
    
    sSql = " SELECT " & F_BEDOUTDT & " as bedoutdt FROM " & T_HIS002 & _
           " WHERE " & DBW(F_INPTID, strPtid, 2) & _
           " AND " & F_BEDINDT & " =to_date(" & BadOut & ",'YYYYMMDD')"
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
'    If Rs.DBerror = True Then
'        dbconn.DisplayErrors
'        Rs.RsClose
'        Set Rs = Nothing
'        Exit Function
'    End If
    If Not Rs.EOF = True Then
        
       '���ſ� �Կ��ߴ� ȯ��.
        GetBedPtInfo strPtid, False
    Else
        'Ȯ���� �Կ�ȯ��
        
        GetBedPtInfo strPtid, True
    End If
    Set Rs = Nothing
End Function

Private Function GetBedPtInfo(ByVal strPtid As String, blnBedIn As Boolean)
    '����ȯ���� ������ ������ �´�.
    'True=�Կ�ȯ�� : False=�ܷ�ȯ��
    Dim strSQL As String
    Dim Rs  As New Recordset
    Dim strTmp As String
    Dim lngsex As Long
    
    mvarPtId = strPtid
    
    If blnBedIn = True Then
        strSQL = "SELECT " & F_SSN2("a") & " as ssn," & _
                       " a." & F_PTNM & " as ptnm," & _
                               F_BEDINDT2("b") & " as bedindt," & _
                       " b." & F_PTDEPTCD & " as deptcd," & _
                       " c." & F_DEPTNM & " as deptnm," & _
                       " b." & F_MAJDOCT & " as majdoct, " & _
                       " e." & F_DOCTNM & " as doctnm," & _
                       " b." & F_PTWARDID & " as wardid," & _
                       " b." & F_PTROOMID & " as hosilid " & _
                " FROM " & T_HIS001 & " a," & T_HIS002 & " b," & T_HIS003 & " c," & T_HIS005 & " e " & _
                " WHERE " & _
                " " & DBW("b." & F_INPTID, strPtid, 2) & " " & _
                " AND b." & F_INPTID & " =a." & F_PTID & _
                " AND " & DBJ("b." & F_PTDEPTCD & "*=c." & F_DEPTCD) & _
                " AND " & DBJ("b." & F_MAJDOCT & "*=e." & F_DOCTID)
    Else
        strSQL = " SELECT " & F_PTID & " as ptid," & F_SSN2 & " as ssn," & F_PTNM & " as ptnm" & _
                 " FROM " & T_HIS001 & _
                 " WHERE " & DBW(F_PTID, strPtid, 2)
    End If
    
    Call Rs.Open(strSQL, DBConn)
    
'    If Rs.DBerror = True Then
'        dbconn.DisplayErrors
'        Rs.RsClose
'        Set Rs = Nothing
'        Exit Function
'    End If

    If blnBedIn = True Then
        If Not Rs.EOF = True Then
            mvarPtDiv = "BED" '����ȯ��
           ' mvarDoB = Format(Mid(Rs.Fields("ssn").Value, 1, 8), "####-##-##")
            mvarPtNm = Rs.Fields("ptnm").Value & ""
            mvarBedIndt = Format(Rs.Fields("BedIndt").Value & "", "####-##-##")
            mvarDeptCd = Rs.Fields("deptcd").Value & ""
            mvarDeptCd = IIf(mvarDeptCd = "NULL", "", mvarDeptCd)
            mvarDeptNm = Rs.Fields("DeptNm").Value & ""
            mvarMajdoct = IIf(Rs.Fields("Majdoct").Value & "" = 0, "", Rs.Fields("majdoct").Value & "")
            mvarDoctNm = Rs.Fields("DoctNm").Value & ""
            mvarWardID = Trim(Rs.Fields("WardId").Value & "")
            mvarHosilID = Trim(Rs.Fields("hosilid").Value & "")
            mvarBedID = ""
            
            Call SexCheck(Rs.Fields("ssn").Value & "")
            
        Else
            mvarPtDiv = "NOBED" '�ܷ�ȯ��
            mvarPtNm = Rs.Fields("ptnm").Value & ""
            mvarBedIndt = ""
            Call SexCheck(Rs.Fields("ssn").Value & "")
        End If
    Else
        mvarPtDiv = "NOBED" '�ܷ�ȯ��
        mvarPtNm = Rs.Fields("ptnm").Value & ""
        mvarBedIndt = ""
        Call SexCheck(Rs.Fields("ssn").Value & "")
    End If
    
    Set Rs = Nothing
    
End Function

Private Sub SexCheck(ByVal ssn As String)
    Dim strTmp  As String
    Dim strYY   As String
    Dim strMM   As String
    Dim strDD   As String
    
    strYY = Trim(Mid(ssn, 1, 2))
    strMM = Trim(Mid(ssn, 3, 2))
    strDD = Trim(Mid(ssn, 5, 2))
    
    On Error Resume Next
    
    If Val(strMM) < 1 Then strMM = "01"
    If Val(strMM) > 12 Then strMM = "12"
    If Val(strDD) < 1 Then strDD = "01"
    If Val(strDD) > 31 Then strDD = "31"
    
    If ssn <> "" Then
        strTmp = Mid(ssn, 7, 1)
        
        Select Case strTmp
            Case "0": mvarSex = "��": mvarDoB = "18" & strYY & "-" & strMM & "-" & strDD
            Case "1": mvarSex = "��": mvarDoB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "2": mvarSex = "��": mvarDoB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "3": mvarSex = "��": mvarDoB = "20" & strYY & "-" & strMM & "-" & strDD
            Case "4": mvarSex = "��": mvarDoB = "20" & strYY & "-" & strMM & "-" & strDD
            Case Else: mvarSex = "��": mvarDoB = "19" & strYY & "-" & strMM & "-" & strDD
        End Select
        
        
        If strTmp Mod 2 = 0 Then
            mvarSex = "��"
        Else
            mvarSex = "��"
        End If
        If Len(ssn) = 13 Then
            If mvarDoB <> "" Then
                mvarAge = medFindAge(Replace(mvarDoB, "-", ""), "Y")
            End If
        End If
    End If


End Sub
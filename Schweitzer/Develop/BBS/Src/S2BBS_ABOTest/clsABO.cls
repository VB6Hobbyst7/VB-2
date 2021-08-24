VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsABO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarPtId As String

Private mvarABO As String
Private mvarRH As String
Private mvarABOSub As String
Private mvarRHSub As String

'Private CODE_ABOFRONT As String
'Private CODE_ABOBACK As String
'Private CODE_RH As String
'Private CODE_ABOSUB As String
'Private CODE_RHSUB As String

Public Property Get ptid() As String
    ptid = mvarPtId
End Property

Public Property Let ptid(ByVal vData As String)
    mvarPtId = vData
    mvarABO = "": mvarRH = ""
End Property

Public Property Get ABO() As String
    ABO = mvarABO
End Property
    
Public Property Get Rh() As String
    Rh = mvarRH
End Property

Public Property Get ABOSub() As String
    ABOSub = mvarABOSub
End Property

Public Property Get RhSub() As String
    RhSub = mvarRHSub
End Property

Private Function Get_ABO(ABO As String) As Boolean
    Dim strSQL     As String
    Dim Rs     As Recordset
    Dim CellABO  As String
    Dim SerumABO As String
    
    strSQL = "SELECT testcd,rstval,rstcd,vfydt,vfytm,vfyid " & _
           "FROM " & T_LAB302 & " " & _
           "WHERE " & DBW("ptid=", mvarPtId) & " " & _
           "AND testcd in ('" & CODE_ABOFRONT & "','" & CODE_ABOBACK & "') " & _
           "AND vfydt=(SELECT max(vfydt) FROM " & T_LAB302 & " " & _
                      "WHERE " & DBW("ptid=", mvarPtId) & " " & _
                      "AND " & DBW("testcd=", CODE_ABOFRONT) & " " & _
                      "AND (vfydt is not null or " & DBW("vfydt<>", "") & " ))"
                      
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
'        dbconn.DisplayErrors
        Get_ABO = False
        ABO = ""
    Else
        With Rs
            If .RecordCount < 1 Then
                Get_ABO = False
                ABO = ""
            Else
                Do Until Rs.EOF
                    If CODE_ABOFRONT = Rs.Fields("testcd").Value & "" Then
                        CellABO = Rs.Fields("rstcd").Value & ""
                    ElseIf CODE_ABOBACK = Rs.Fields("testcd").Value & "" Then
                        SerumABO = Rs.Fields("rstcd").Value & ""
                    End If
                    Rs.MoveNext
                Loop
                If CellABO <> "" Then
                    ABO = CellABO
                    If SerumABO <> "" Then ABO = ABO & "(" & SerumABO & ")"
                Else
                    If SerumABO <> "" Then ABO = "(" & ABO & ")"
                End If
                Get_ABO = True
'                ABO = .Fields("rstcd").Value & ""
            End If
        End With
        Set Rs = Nothing
    End If
    If ABO = "" Then
        strSQL = " SELECT * FROM " & T_BBS902 & " WHERE " & DBW("ptid=", mvarPtId)
        Set Rs = New Recordset
        Rs.Open strSQL, DBConn
        If Not Rs.EOF Then
            ABO = Rs.Fields("abo1").Value & ""
            If ABO <> "" Then
                If Rs.Fields("abo2").Value & "" <> "" Then ABO = ABO & "(" & Rs.Fields("abo2").Value & "" & ")"
            Else
                If Rs.Fields("abo2").Value & "" <> "" Then ABO = Rs.Fields("abo2").Value & ""
            End If
        End If
    End If
    If ABO <> "" Then Get_ABO = True
    Set Rs = Nothing
End Function

Private Function Get_RH(Rh As String) As Boolean
    Dim strSQL As String
    Dim Rs As Recordset
    
    strSQL = "SELECT testcd,rstval,rstcd,vfydt,vfytm,vfyid " & _
           "FROM " & T_LAB302 & " " & _
           "WHERE " & DBW("ptid=", mvarPtId) & " " & _
           "AND " & DBW("testcd=", CODE_RH) & " " & _
           "AND vfydt=(SELECT max(vfydt) FROM " & T_LAB302 & " " & _
                      "WHERE " & DBW("ptid=", mvarPtId) & " " & _
                      "AND " & DBW("testcd=", CODE_RH) & " " & _
                      "AND (vfydt is not null or " & DBW("vfydt<>", "") & "))"
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
'        dbconn.DisplayErrors
        Get_RH = False
        Rh = ""
    Else
        With Rs
            If .RecordCount < 1 Then
                Get_RH = False
                Rh = ""
            Else
                Get_RH = True
                Rh = .Fields("rstcd").Value & ""
            End If
        End With
        Set Rs = Nothing
    End If
    If Rh = "" Then
        strSQL = " SELECT * FROM " & T_BBS902 & " WHERE " & DBW("ptid=", mvarPtId)
        Set Rs = New Recordset
        Rs.Open strSQL, DBConn
        If Not Rs.EOF Then
            Rh = Rs.Fields("rh").Value & ""
        End If
    End If
    If Rh <> "" Then Get_RH = True
    Set Rs = Nothing
End Function

Private Function Get_ABOSUB(ABOSub As String) As Boolean
    Dim strSQL As String
    Dim Rs As Recordset
    
    strSQL = "SELECT testcd,rstval,rstcd,vfydt,vfytm,vfyid " & _
           "FROM " & T_LAB302 & " " & _
           "WHERE " & DBW("ptid=", mvarPtId) & " " & _
           "AND " & DBW("testcd=", CODE_ABOSUB) & " " & _
           "AND vfydt=(SELECT max(vfydt) FROM " & T_LAB302 & " " & _
                      "WHERE " & DBW("ptid=", mvarPtId) & " " & _
                      "AND " & DBW("testcd=", CODE_ABOSUB) & " " & _
                      "AND vfydt is not null " & _
                      "AND " & DBW("vfydt<>", "") & ")"
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
'        dbconn.DisplayErrors
        Get_ABOSUB = False
        ABOSub = ""
    Else
        With Rs
            If .RecordCount < 1 Then
                Get_ABOSUB = False
                ABOSub = ""
            Else
                Get_ABOSUB = True
                ABOSub = .Fields("rstcd").Value & ""
            End If
        End With
        Set Rs = Nothing
    End If
End Function

Private Function Get_RHSUB(RhSub As String) As Boolean
    Dim strSQL As String
    Dim Rs As Recordset
    
    strSQL = "SELECT testcd,rstval,rstcd,vfydt,vfytm,vfyid " & _
           "FROM " & T_LAB302 & " " & _
           "WHERE " & DBW("ptid=", mvarPtId) & " " & _
           "AND " & DBW("testcd=", CODE_RHSUB) & " " & _
           "AND vfydt=(SELECT max(vfydt) FROM " & T_LAB302 & " " & _
                      "WHERE " & DBW("ptid=", mvarPtId) & " " & _
                      "AND " & DBW("testcd=", CODE_RHSUB) & " " & _
                      "AND vfydt is not null " & _
                      "AND " & DBW("vfydt<>", "") & ")"
    Set Rs = New Recordset
    Call Rs.Open(strSQL, DBConn)
    If Rs.EOF Then
'        dbconn.DisplayErrors
        Get_RHSUB = False
        RhSub = ""
    Else
        With Rs
            If .RecordCount < 1 Then
                Get_RHSUB = False
                RhSub = ""
            Else
                Get_RHSUB = True
                RhSub = .Fields("rstcd").Value & ""
            End If
        End With
        Set Rs = Nothing
    End If
End Function

Public Function GetABO() As Boolean
    Dim ABO As String
    Dim Rh As String
    Dim ABOSub As String
    Dim RhSub As String
    
    GetABO = False
    
    If mvarPtId = "" Then
        MsgBox "환자ID가 없습니다.", vbCritical, "오류"
        Exit Function
    End If
    
    If Get_ABO(ABO) = False Then Exit Function
    If Get_RH(Rh) = False Then Exit Function
'    If Get_ABOSUB(abosub) = False Then Exit Function
'    If Get_RHSUB(rhsub) = False Then Exit Function
    
    GetABO = True
    
    mvarABO = ABO
    mvarRH = Rh
    mvarABOSub = ABOSub
    mvarRHSub = RhSub
    
End Function

Private Sub Class_Initialize()
'    Dim objcom003 As clsCom003
'    Dim Rs As Recordset
'
'    Set objcom003 = New clsCom003
'    Set Rs = objcom003.OpenRecordSetDay(BC2_ABO_TEST)
'    Set objcom003 = Nothing
'
'    If Not (Rs Is Nothing) Then
'        With Rs
'            If .RecordCount > 0 Then
'                CODE_ABOFRONT = .Fields("field1").Value & ""
'                CODE_ABOBACK = .Fields("field2").Value & ""
'                CODE_RH = .Fields("field3").Value & ""
'                CODE_ABOSUB = .Fields("field4").Value & ""
'                CODE_RHSUB = .Fields("text1").Value & ""
'            Else
'                CODE_ABOFRONT = ""
'                CODE_ABOBACK = ""
'                CODE_RH = ""
'                CODE_ABOSUB = ""
'                CODE_RHSUB = ""
'            End If
'        End With
'        Set Rs = Nothing
'    End If
    mvarABO = "": mvarRH = ""
    Call LoadTestCd
End Sub

Public Function GetLastRst(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As String
    Dim strSQL As String
    Dim strTmp As String
    Dim strLstVfyDtTmId As String
    
    Dim Rs   As Recordset
    
    'ABOFRONT
    strSQL = " SELECT * " & _
           " FROM " & T_LAB302 & " " & _
           " WHERE " & DBW("workarea=", vWorkarea) & " " & _
           " AND   " & DBW("accdt   =", vAccdt) & " " & _
           " AND   " & DBW("accseq  =", vAccseq) & " " & _
           " AND   " & DBW("testcd=", CODE_ABOFRONT)
           
    Set Rs = Nothing
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Not Rs.EOF Then
        strTmp = Rs.Fields("lastrst").Value & ""
        If Rs.Fields("lastrst").Value & "" <> "" Then
            strLstVfyDtTmId = "  ABO Front : " & _
                              Format(Rs.Fields("lastrst").Value & "", "!" & String(5, "@")) & " " & _
                              Format(Rs.Fields("lastvfydt").Value & "", "0###/0#/0#") & " " & _
                              Format(Rs.Fields("lastvfytm").Value & "", "0#:0#:0#") & " " & _
                              GetEmpNm(Rs.Fields("lastvfyid").Value & "") & COL_DIV
        End If
    End If
    Set Rs = Nothing
    
    'ABOBACK
    strSQL = " SELECT * " & _
           " FROM " & T_LAB302 & " " & _
           " WHERE " & DBW("workarea=", vWorkarea) & " " & _
           " AND   " & DBW("accdt   =", vAccdt) & " " & _
           " AND   " & DBW("accseq  =", vAccseq) & " " & _
           " AND   " & DBW("testcd=", CODE_ABOBACK)
           
    Set Rs = Nothing
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Not Rs.EOF Then
        If strTmp = "" Then
            strTmp = strTmp & Rs.Fields("lastrst").Value & ""
        Else
            strTmp = strTmp & "(" & Rs.Fields("lastrst").Value & "" & ")"
        End If
        
        If Rs.Fields("lastrst").Value & "" <> "" Then
            strLstVfyDtTmId = strLstVfyDtTmId & _
                              "  ABO Back  : " & _
                              Format(Rs.Fields("lastrst").Value & "", "!" & String(5, "@")) & " " & _
                              Format(Rs.Fields("lastvfydt").Value & "", "0###/0#/0#") & " " & _
                              Format(Rs.Fields("lastvfytm").Value & "", "0#:0#:0#") & " " & _
                              GetEmpNm(Rs.Fields("lastvfyid").Value & "") & COL_DIV
        End If
    End If
    Set Rs = Nothing
    
    
    strSQL = " SELECT * " & _
           " FROM " & T_LAB302 & " " & _
           " WHERE " & DBW("workarea=", vWorkarea) & " " & _
           " AND   " & DBW("accdt   =", vAccdt) & " " & _
           " AND   " & DBW("accseq  =", vAccseq) & " " & _
           " AND   " & DBW("testcd=", CODE_RH)
    Set Rs = Nothing
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Not Rs.EOF Then
        strTmp = strTmp & Rs.Fields("lastrst").Value & ""
        
        If Rs.Fields("lastrst").Value & "" <> "" Then
            strLstVfyDtTmId = strLstVfyDtTmId & _
                              "  Rh        : " & _
                              Format(Rs.Fields("lastrst").Value & "", "!" & String(5, "@")) & " " & _
                              Format(Rs.Fields("lastvfydt").Value & "", "0###/0#/0#") & " " & _
                              Format(Rs.Fields("lastvfytm").Value & "", "0#:0#:0#") & " " & _
                              GetEmpNm(Rs.Fields("lastvfyid").Value & "") & COL_DIV
        End If
    End If
    Set Rs = Nothing
    
    'ABOSUB는 툴팁으로 보여줄 최근결과만 표시
    strSQL = " SELECT * " & _
           " FROM " & T_LAB302 & " " & _
           " WHERE " & DBW("workarea=", vWorkarea) & " " & _
           " AND   " & DBW("accdt   =", vAccdt) & " " & _
           " AND   " & DBW("accseq  =", vAccseq) & " " & _
           " AND   " & DBW("testcd=", CODE_ABOSUB)
    Set Rs = Nothing
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Not Rs.EOF Then
        If Rs.Fields("lastrst").Value & "" <> "" Then
            strLstVfyDtTmId = strLstVfyDtTmId & _
                              "  ABOSub    : " & _
                              Format(Rs.Fields("lastrst").Value & "", "!" & String(5, "@")) & " " & _
                              Format(Rs.Fields("lastvfydt").Value & "", "0###/0#/0#") & " " & _
                              Format(Rs.Fields("lastvfytm").Value & "", "0#:0#:0#") & " " & _
                              GetEmpNm(Rs.Fields("lastvfyid").Value & "") & COL_DIV
        End If
    End If
    Set Rs = Nothing
    
    'RHSUB
    strSQL = " SELECT * " & _
           " FROM " & T_LAB302 & " " & _
           " WHERE " & DBW("workarea=", vWorkarea) & " " & _
           " AND   " & DBW("accdt   =", vAccdt) & " " & _
           " AND   " & DBW("accseq  =", vAccseq) & " " & _
           " AND   " & DBW("testcd=", CODE_RHSUB)
    Set Rs = Nothing
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Not Rs.EOF Then
        If Rs.Fields("lastrst").Value & "" <> "" Then
            strLstVfyDtTmId = strLstVfyDtTmId & _
                              "  RhSub     : " & _
                              Format(Rs.Fields("lastrst").Value & "", "!" & String(5, "@")) & " " & _
                              Format(Rs.Fields("lastvfydt").Value & "", "0###/0#/0#") & " " & _
                              Format(Rs.Fields("lastvfytm").Value & "", "0#:0#:0#") & " " & _
                              GetEmpNm(Rs.Fields("lastvfyid").Value & "") & COL_DIV
        End If
    End If
    Set Rs = Nothing
    
    GetLastRst = strTmp & LINE_DIV & strLstVfyDtTmId
End Function

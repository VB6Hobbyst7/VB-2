VERSION 5.00
Begin VB.Form frmMDBSync 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDB Sync"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "맑은 고딕"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6075
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox chkOrder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "작업지시서 포함"
      Height          =   225
      Left            =   3900
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   2
      Top             =   990
      Width           =   5445
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "DataBase Sync"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3900
      TabIndex        =   0
      Top             =   420
      Width           =   1845
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  '투명
      Caption         =   "로컬 데이터베이스에 접속되었습니다."
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   1
      Top             =   510
      Width           =   3585
   End
End
Attribute VB_Name = "frmMDBSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim AdoCn_Local                As ADODB.Connection
'Dim AdoRs_Local                As ADODB.Recordset

Private Sub Form_Load()

    '-- 로컬 DB 접속
    If Not DbConnect_Local_UpDate Then
        MsgBox "로컬 데이터베이스 접속실패", vbCritical
        Unload Me
    Else
        lblStatus.Caption = "로컬 데이터베이스에 접속되었습니다."
        cn_Server_Flag = True
    End If
    
End Sub


Private Sub cmdStart_Click()

On Error GoTo ErrorRoutine
    
    If MsgBox("데이터베이스 Sync 작업을 진행하시겠습니까?" & vbNewLine & " 이 작업은 로컬 데이터베이스를 변경합니다." & vbNewLine & "  (트래킹 데이터는 제외됩니다) ", vbInformation + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    
    txtStatus.Text = ""
    
    'Master
    Call GetUserSync
    txtStatus.Text = "User Master SYNC 성공"
    
    Call GetCompSync
    txtStatus.Text = txtStatus.Text & vbCrLf & "Comp Master SYNC 성공"
    
    Call GetTempSync
    txtStatus.Text = txtStatus.Text & vbCrLf & "Temp Master SYNC 성공"
    
    Call GetMateSync
    txtStatus.Text = txtStatus.Text & vbCrLf & "Material Master SYNC 성공"
    
    Call GetPackSync
    txtStatus.Text = txtStatus.Text & vbCrLf & "Pack Master SYNC 성공"
    
    Call GetProdSync
    txtStatus.Text = txtStatus.Text & vbCrLf & "Prod Master SYNC 성공"
    
    'Information
    Call GetLabelMasterSync
    txtStatus.Text = txtStatus.Text & vbCrLf & "라벨 Master SYNC 성공"

    Call GetLabelMasterDetailSync
    txtStatus.Text = txtStatus.Text & vbCrLf & "라벨 Detail SYNC 성공"

    Call GetBarMasterSync
    txtStatus.Text = txtStatus.Text & vbCrLf & "바코드 Master SYNC 성공"

    Call GetBarMasterDetailSync
    txtStatus.Text = txtStatus.Text & vbCrLf & "바코드 Detail SYNC 성공"

    If chkOrder.Value = "1" Then
        '오더정보
        Call GetPackSync
        txtStatus.Text = txtStatus.Text & vbCrLf & "작업지시서 Master SYNC 성공"
        
        Call GetProdSlittingSync
        txtStatus.Text = txtStatus.Text & vbCrLf & "작업지시서 Detail SYNC 성공"
        
'        Call GetTrackSync
'        txtStatus.Text = txtStatus.Text & vbCrLf & "트래킹 Data SYNC 성공"
        
        Call GetMaxNoSync
        txtStatus.Text = txtStatus.Text & vbCrLf & "MaxNo Data SYNC 성공"
        
        'LBL_MAX_NO
        'LBL_PACK_TRACK
        'LBL_PROD_ORDER
        'LBL_SLITING_ORDER
    End If
    
Exit Sub

ErrorRoutine:
    Call DBErrorSet(AdoCn, SQL, "DB Sync")
    
End Sub

Private Sub GetUserSync()

    Set AdoRs = Get_UserList("")
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        '1. 사용자 마스터
        SQL = "DELETE FROM LBL_M_USER "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_M_USER     " & vbCrLf
            SQL = SQL & " ( USER_CD                 " & vbCrLf
            SQL = SQL & " , USER_NAME               " & vbCrLf
            SQL = SQL & " , USER_PW                 " & vbCrLf
            SQL = SQL & " , USER_DEPART             " & vbCrLf
            SQL = SQL & " , USER_COMP               " & vbCrLf
            SQL = SQL & " , USED_YN                 " & vbCrLf
            SQL = SQL & " , REGIST_ID               " & vbCrLf
            SQL = SQL & " , REGIST_DT               "
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & " , MODIFY_ID               " & vbCrLf
                SQL = SQL & " , MODIFY_DT               "
            End If
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("USER_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USER_NAME").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USER_PW").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USER_DEPART").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USER_COMP").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USED_YN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID").Value & "'" & vbCrLf
            SQL = SQL & ", '" & Format(AdoRs.Fields("REGIST_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & ", '" & AdoRs.Fields("MODIFY_ID").Value & "'" & vbCrLf
                SQL = SQL & ", '" & Format(AdoRs.Fields("MODIFY_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub GetCompSync()

    Set AdoRs = Get_CompList("")
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_M_COMP "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_M_COMP     " & vbCrLf
            SQL = SQL & " ( COMP_CD                 " & vbCrLf
            SQL = SQL & " , COMP_NAME               " & vbCrLf
            SQL = SQL & " , COMP_LINE               " & vbCrLf
            SQL = SQL & " , COMP_VIEW               " & vbCrLf
            SQL = SQL & " , COMP_DIS_NO             " & vbCrLf
            SQL = SQL & " , USED_YN                 " & vbCrLf
            SQL = SQL & " , REGIST_ID               " & vbCrLf
            SQL = SQL & " , REGIST_DT               "
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & " , MODIFY_ID               " & vbCrLf
                SQL = SQL & " , MODIFY_DT               "
            End If
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("COMP_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("COMP_NAME").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("COMP_LINE").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("COMP_VIEW").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("COMP_DIS_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USED_YN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID").Value & "'" & vbCrLf
            SQL = SQL & ", '" & Format(AdoRs.Fields("REGIST_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & ", '" & AdoRs.Fields("MODIFY_ID").Value & "'" & vbCrLf
                SQL = SQL & ", '" & Format(AdoRs.Fields("MODIFY_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub GetPackSync()

    Set AdoRs = Get_PackList("")
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_M_PACK "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_M_PACK     " & vbCrLf
            SQL = SQL & " ( PACK_CD                 " & vbCrLf
            SQL = SQL & " , PACK_NAME               " & vbCrLf
            SQL = SQL & " , PACK_CORE               " & vbCrLf
            SQL = SQL & " , PACK_DIA                " & vbCrLf
            SQL = SQL & " , PACK_DIS_NO             " & vbCrLf
            SQL = SQL & " , PACK_CAT_WIDTH          " & vbCrLf
            SQL = SQL & " , PACK_PRO_WIDTH          " & vbCrLf
            SQL = SQL & " , PACK_PRO_LENGTH         " & vbCrLf
            SQL = SQL & " , PACK_CAT_GU             " & vbCrLf
            SQL = SQL & " , USED_YN                 " & vbCrLf
            SQL = SQL & " , REGIST_ID               " & vbCrLf
            SQL = SQL & " , REGIST_DT               "
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & " , MODIFY_ID               " & vbCrLf
                SQL = SQL & " , MODIFY_DT               "
            End If
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("PACK_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PACK_NAME").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PACK_CORE").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PACK_DIA").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PACK_DIS_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PACK_CAT_WIDTH").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PACK_PRO_WIDTH").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PACK_PRO_LENGTH").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PACK_CAT_GU").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USED_YN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID").Value & "'" & vbCrLf
            SQL = SQL & ", '" & Format(AdoRs.Fields("REGIST_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & ", '" & AdoRs.Fields("MODIFY_ID").Value & "'" & vbCrLf
                SQL = SQL & ", '" & Format(AdoRs.Fields("MODIFY_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub GetTrackSync()

    Set AdoRs = Get_TrackList()
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_PACK_TRACK "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_PACK_TRACK     " & vbCrLf
            SQL = SQL & " ( PROD_ORDER_DT               " & vbCrLf
            SQL = SQL & " , PROD_CD                     " & vbCrLf
            SQL = SQL & " , PROD_REEL_BAR               " & vbCrLf
            SQL = SQL & " , PROD_PP_BAR                 " & vbCrLf
            SQL = SQL & " , PROD_ICE_BAR                " & vbCrLf
            SQL = SQL & " , PROD_PP_BAR_IN              " & vbCrLf
            SQL = SQL & " , PROD_ICE_BAR_IN             " & vbCrLf
            SQL = SQL & " , PROD_LOT_NO                 " & vbCrLf
            SQL = SQL & " , REGIST_ID_R                 " & vbCrLf
            SQL = SQL & " , REGIST_DT_R                 " & vbCrLf
            SQL = SQL & " , REGIST_ID_P                 " & vbCrLf
            SQL = SQL & " , REGIST_DT_P                 " & vbCrLf
            SQL = SQL & " , REGIST_ID_I                 " & vbCrLf
            SQL = SQL & " , REGIST_DT_I                 " & vbCrLf
            SQL = SQL & " , REEL_PRT_VAL                " & vbCrLf
            SQL = SQL & " , PP_PRT_VAL                  " & vbCrLf
            SQL = SQL & " , ICE_PRT_VAL                 "
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                       " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("PROD_ORDER_DT").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_REEL_BAR").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_PP_BAR").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_ICE_BAR").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_PP_BAR_IN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_ICE_BAR_IN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_LOT_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID_R").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_DT_R").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID_P").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_DT_P").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID_I").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_DT_I").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REEL_PRT_VAL").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PP_PRT_VAL").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("ICE_PRT_VAL").Value & "'" & vbCrLf
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub GetMaxNoSync()

    Set AdoRs = Get_MaxNo()
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_MAX_NO "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_MAX_NO     " & vbCrLf
            SQL = SQL & " ( PROD_ORDER_DT           " & vbCrLf
            SQL = SQL & " , PROD_CD                 " & vbCrLf
            SQL = SQL & " , BOX_GU                  " & vbCrLf
            SQL = SQL & " , MAX_NO                  " & vbCrLf
            SQL = SQL & " , LOT_NO                  " & vbCrLf
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("PROD_ORDER_DT").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("BOX_GU").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("MAX_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LOT_NO").Value & "'" & vbCrLf
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub GetProdOrderSync()

    Set AdoRs = Get_ProdOrderList("")
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_PROD_ORDER "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_PROD_ORDER " & vbCrLf
            SQL = SQL & " ( PROD_ORDER_DT           " & vbCrLf
            SQL = SQL & " , PROD_CD                 " & vbCrLf
            SQL = SQL & " , SLITING_NO              " & vbCrLf
            SQL = SQL & " , COMP_CD                 " & vbCrLf
            SQL = SQL & " , PROD_NAME               " & vbCrLf
            SQL = SQL & " , PACK_CD                 " & vbCrLf
            SQL = SQL & " , REEL_QTY                " & vbCrLf
            SQL = SQL & " , JOB_INFO                " & vbCrLf
            SQL = SQL & " , ORDER_MEMO              " & vbCrLf
            SQL = SQL & " , LOT_NO                  " & vbCrLf
            SQL = SQL & " , CLOSE_YN                " & vbCrLf
            SQL = SQL & " , REGIST_ID               " & vbCrLf
            SQL = SQL & " , REGIST_DT               "
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & " , MODIFY_ID               " & vbCrLf
                SQL = SQL & " , MODIFY_DT               "
            End If
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("PROD_ORDER_DT").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("SLITING_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("COMP_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_NAME").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PACK_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REEL_QTY").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("JOB_INFO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("ORDER_MEMO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LOT_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("CLOSE_YN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID").Value & "'" & vbCrLf
            SQL = SQL & ", '" & Format(AdoRs.Fields("REGIST_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & ", '" & AdoRs.Fields("MODIFY_ID").Value & "'" & vbCrLf
                SQL = SQL & ", '" & Format(AdoRs.Fields("MODIFY_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub GetProdSlittingSync()

    Set AdoRs = Get_ProdSlittingList()
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_SLITING_ORDER "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_SLITING_ORDER " & vbCrLf
            SQL = SQL & " ( PROD_ORDER_DT           " & vbCrLf
            SQL = SQL & " , PROD_CD                 " & vbCrLf
            SQL = SQL & " , SLITING_NO              " & vbCrLf
            SQL = SQL & " , SEQ_NO                 " & vbCrLf
            SQL = SQL & " , SLITING_INFO               " & vbCrLf
            SQL = SQL & " , P_NO_F                 " & vbCrLf
            SQL = SQL & " , P_NO_T                " & vbCrLf
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("PROD_ORDER_DT").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("SLITING_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("SEQ_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("SLITING_INFO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("P_NO_F").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("P_NO_T").Value & "'" & vbCrLf
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub GetProdSync()

    Set AdoRs = Get_ProdList("")
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_M_PROD "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_M_PROD     " & vbCrLf
            SQL = SQL & " ( PROD_CD                 " & vbCrLf
            SQL = SQL & " , PROD_NAME               " & vbCrLf
            SQL = SQL & " , PROD_PRT_NAME               " & vbCrLf
            SQL = SQL & " , COMP_CD                " & vbCrLf
            SQL = SQL & " , PROD_LENGTH             " & vbCrLf
            SQL = SQL & " , PROD_MATERIAL_CD          " & vbCrLf
            SQL = SQL & " , EXPIR_MONTH          " & vbCrLf
            SQL = SQL & " , PROD_STOR_TEMP         " & vbCrLf
            SQL = SQL & " , PROD_SIZE             " & vbCrLf
            SQL = SQL & " , PROD_CHIMEI_PN                " & vbCrLf
            SQL = SQL & " , VENDER_CD             " & vbCrLf
            SQL = SQL & " , PROD_LINE_FA          " & vbCrLf
            SQL = SQL & " , PROD_SLIT_FA          " & vbCrLf
            SQL = SQL & " , PROD_CONTROL_YN         " & vbCrLf
            SQL = SQL & " , PROD_PCN_NO             " & vbCrLf
            SQL = SQL & " , ITEM_BARCODE             " & vbCrLf
            SQL = SQL & " , USED_YN                 " & vbCrLf
            SQL = SQL & " , REGIST_ID               " & vbCrLf
            SQL = SQL & " , REGIST_DT               "
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & " , MODIFY_ID               " & vbCrLf
                SQL = SQL & " , MODIFY_DT               "
            End If
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("PROD_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_NAME").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_PRT_NAME").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("COMP_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_LENGTH").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_MATERIAL_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("EXPIR_MONTH").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_STOR_TEMP").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_SIZE").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_CHIMEI_PN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("VENDER_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_LINE_FA").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_SLIT_FA").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_CONTROL_YN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_PCN_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("ITEM_BARCODE").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USED_YN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID").Value & "'" & vbCrLf
            SQL = SQL & ", '" & Format(AdoRs.Fields("REGIST_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & ", '" & AdoRs.Fields("MODIFY_ID").Value & "'" & vbCrLf
                SQL = SQL & ", '" & Format(AdoRs.Fields("MODIFY_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub GetLabelMasterSync()

    Set AdoRs = Get_LabelMaster("")
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_LABEL_MASTER "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_LABEL_MASTER   " & vbCrLf
            SQL = SQL & " ( PROD_LABEL_CD               " & vbCrLf
            SQL = SQL & " , PROD_CD                     " & vbCrLf
            SQL = SQL & " , COMP_CD                     " & vbCrLf
            SQL = SQL & " , PROD_LABEL_TYPE             " & vbCrLf
            SQL = SQL & " , LABEL_PRT_NO                " & vbCrLf
            SQL = SQL & " , LABEL_PRT_SIDE              " & vbCrLf
            SQL = SQL & " , LABEL_BAR_SIDE01_TYPE       " & vbCrLf
            SQL = SQL & " , LABEL_BAR_SIDE02_TYPE       " & vbCrLf
            SQL = SQL & " , PROD_MAX_TOT                " & vbCrLf
            SQL = SQL & " , USED_YN                     " & vbCrLf
            SQL = SQL & " , REGIST_ID                   " & vbCrLf
            SQL = SQL & " , REGIST_DT                   "
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & " , MODIFY_ID               " & vbCrLf
                SQL = SQL & " , MODIFY_DT               "
            End If
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("PROD_LABEL_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("COMP_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_LABEL_TYPE").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_PRT_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_PRT_SIDE").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_BAR_SIDE01_TYPE").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_BAR_SIDE02_TYPE").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_MAX_TOT").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USED_YN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID").Value & "'" & vbCrLf
            SQL = SQL & ", '" & Format(AdoRs.Fields("REGIST_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & ", '" & AdoRs.Fields("MODIFY_ID").Value & "'" & vbCrLf
                SQL = SQL & ", '" & Format(AdoRs.Fields("MODIFY_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub GetBarMasterSync()

    Set AdoRs = Get_BarMaster("")
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_BAR_MASTER "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_BAR_MASTER     " & vbCrLf
            SQL = SQL & " ( BAR_CD                      " & vbCrLf
            SQL = SQL & " , PROD_CD                     " & vbCrLf
            SQL = SQL & " , COMP_CD                     " & vbCrLf
            SQL = SQL & " , BAR_TYPE                    " & vbCrLf
            SQL = SQL & " , BAR_GU                      " & vbCrLf
            SQL = SQL & " , TEMP_MASTER_GU              " & vbCrLf
            SQL = SQL & " , USED_YN                     " & vbCrLf
            SQL = SQL & " , REGIST_ID                   " & vbCrLf
            SQL = SQL & " , REGIST_DT                   "
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & " , MODIFY_ID               " & vbCrLf
                SQL = SQL & " , MODIFY_DT               "
            End If
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("BAR_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("PROD_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("COMP_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("BAR_TYPE").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("BAR_GU").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("TEMP_MASTER_GU").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USED_YN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID").Value & "'" & vbCrLf
            SQL = SQL & ", '" & Format(AdoRs.Fields("REGIST_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & ", '" & AdoRs.Fields("MODIFY_ID").Value & "'" & vbCrLf
                SQL = SQL & ", '" & Format(AdoRs.Fields("MODIFY_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub


Private Sub GetLabelMasterDetailSync()

    Set AdoRs = Get_LabelMasterDetail()
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_LABEL_DETAIL "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_LABEL_DETAIL   " & vbCrLf
            SQL = SQL & " ( PROD_LABEL_CD               " & vbCrLf
            SQL = SQL & " , LABEL_ITEM_NO                     " & vbCrLf
            SQL = SQL & " , LABEL_ITEM_SEQ                     " & vbCrLf
            SQL = SQL & " , LABEL_ITEM_NAME             " & vbCrLf
            SQL = SQL & " , LABEL_ITEM_NAME_PRT              " & vbCrLf
            SQL = SQL & " , BAR_CD                " & vbCrLf
            SQL = SQL & " , LABEL_ITEM_X_COORD              " & vbCrLf
            SQL = SQL & " , LABEL_ITEM_Y_COORD       " & vbCrLf
            SQL = SQL & " , LABEL_ITEM_GU       " & vbCrLf
            SQL = SQL & " , LABEL_ITEM_FONT                " & vbCrLf
            SQL = SQL & " , LABEL_ITEM_ROT                " & vbCrLf
            SQL = SQL & " , USED_YN                     " & vbCrLf
            SQL = SQL & " , REGIST_ID                   " & vbCrLf
            SQL = SQL & " , REGIST_DT                   "
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & " , MODIFY_ID               " & vbCrLf
                SQL = SQL & " , MODIFY_DT               "
            End If
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("PROD_LABEL_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_ITEM_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_ITEM_SEQ").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_ITEM_NAME").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("BAR_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_ITEM_GU").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_ITEM_FONT").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_ITEM_ROT").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USED_YN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID").Value & "'" & vbCrLf
            SQL = SQL & ", '" & Format(AdoRs.Fields("REGIST_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & ", '" & AdoRs.Fields("MODIFY_ID").Value & "'" & vbCrLf
                SQL = SQL & ", '" & Format(AdoRs.Fields("MODIFY_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub GetBarMasterDetailSync()

    Set AdoRs = Get_BarMasterDetail()
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_BAR_DETAIL "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_BAR_DETAIL     " & vbCrLf
            SQL = SQL & " ( BAR_CD                      " & vbCrLf
            SQL = SQL & " , BAR_ITEM_NO                 " & vbCrLf
            SQL = SQL & " , BAR_ITEM_SEQ                " & vbCrLf
            SQL = SQL & " , BAR_ITEM_NAME               " & vbCrLf
            SQL = SQL & " , BAR_CHR_NUM                 " & vbCrLf
            SQL = SQL & " , LABEL_ITEM_TYPE             " & vbCrLf
            SQL = SQL & " , USED_YN                     " & vbCrLf
            SQL = SQL & " , REGIST_ID                   " & vbCrLf
            SQL = SQL & " , REGIST_DT                   "
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & " , MODIFY_ID               " & vbCrLf
                SQL = SQL & " , MODIFY_DT               "
            End If
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("BAR_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("BAR_ITEM_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("BAR_ITEM_SEQ").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("BAR_ITEM_NAME").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("BAR_CHR_NUM").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("LABEL_ITEM_TYPE").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USED_YN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID").Value & "'" & vbCrLf
            SQL = SQL & ", '" & Format(AdoRs.Fields("REGIST_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & ", '" & AdoRs.Fields("MODIFY_ID").Value & "'" & vbCrLf
                SQL = SQL & ", '" & Format(AdoRs.Fields("MODIFY_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub


Private Sub GetMateSync()

    Set AdoRs = Get_Material()
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM LBL_M_MATERIAL "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_M_MATERIAL     " & vbCrLf
            SQL = SQL & " ( MAT_CD                 " & vbCrLf
            SQL = SQL & " , MAT_NAME               " & vbCrLf
            SQL = SQL & " , MAT_DIS_NO               " & vbCrLf
            SQL = SQL & " , USED_YN                 " & vbCrLf
            SQL = SQL & " , REGIST_ID               " & vbCrLf
            SQL = SQL & " , REGIST_DT               "
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & " , MODIFY_ID               " & vbCrLf
                SQL = SQL & " , MODIFY_DT               "
            End If
            SQL = SQL & ")"
            SQL = SQL & "  VALUES                   " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("MAT_CD").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("MAT_NAME").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("MAT_DIS_NO").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("USED_YN").Value & "'" & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("REGIST_ID").Value & "'" & vbCrLf
            SQL = SQL & ", '" & Format(AdoRs.Fields("REGIST_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            If AdoRs.Fields("MODIFY_ID").Value & "" <> "" Then
                SQL = SQL & ", '" & AdoRs.Fields("MODIFY_ID").Value & "'" & vbCrLf
                SQL = SQL & ", '" & Format(AdoRs.Fields("MODIFY_DT").Value, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub


Private Sub GetTempSync()

    Set AdoRs = Get_TempList("")
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        SQL = "DELETE FROM TEMP_MASTER "
        Call DBExec(AdoCn_Local, SQL)
        
        Do Until AdoRs.EOF
            SQL = ""
            SQL = SQL & "INSERT INTO TEMP_MASTER    " & vbCrLf
            SQL = SQL & " ( GUBUN_CD                " & vbCrLf
            SQL = SQL & " , SEQNO               " & vbCrLf
            SQL = SQL & " , CODE1               " & vbCrLf
            SQL = SQL & " , CODE2               " & vbCrLf
            SQL = SQL & " , CODE3               " & vbCrLf
            SQL = SQL & " , NAME1               " & vbCrLf
            SQL = SQL & " , NAME2               " & vbCrLf
            SQL = SQL & " , NAME3               " & vbCrLf
            SQL = SQL & " , GUBUN_MEMO          " & vbCrLf
            SQL = SQL & ")"
            SQL = SQL & "  VALUES               " & vbCrLf
            SQL = SQL & "( '" & AdoRs.Fields("GUBUN_CD").Value & "' " & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("SEQNO").Value & "'    " & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("CODE1").Value & "'    " & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("CODE2").Value & "'    " & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("CODE3").Value & "'    " & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("NAME1").Value & "'    " & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("NAME2").Value & "'    " & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("NAME3").Value & "'    " & vbCrLf
            SQL = SQL & ", '" & AdoRs.Fields("GUBUN_MEMO").Value & "'" & vbCrLf
            SQL = SQL & ")"
            
            AdoRs.MoveNext
            Call DBExec(AdoCn_Local, SQL)
        Loop
        AdoRs.Close
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    AdoCn_Local.Close
    
End Sub

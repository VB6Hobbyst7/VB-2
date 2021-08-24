Attribute VB_Name = "sl_d_60_c"
Function sql_prepare_sl_d_60&(db$, login$, pwd$)
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "db", db)
        rv& = dce_push_string(Socket, "login", login)
        rv& = dce_push_string(Socket, "pwd", pwd)
        dce_table = dce_submit("sl_d_60", "sql_prepare_sl_d_60", Socket)
    End If
    sql_prepare_sl_d_60 = dce_pop_long(dce_table, "dce_result")
    Call dce_table_destroy(dce_table)
End Function


Function sql_rows_sl_d_60&(maxrows&)
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        Call dce_push_long(Socket, "maxrows", maxrows)
        dce_table = dce_submit("sl_d_60", "sql_rows_sl_d_60", Socket)
    End If
    sql_rows_sl_d_60 = dce_pop_long(dce_table, "dce_result")
    Call dce_table_destroy(dce_table)
End Function


Function sql_set_max_rows_sl_d_60&(maxrows&)
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        Call dce_push_long(Socket, "maxrows", maxrows)
        dce_table = dce_submit("sl_d_60", "sql_set_max_rows_sl_d_60", Socket)
    End If
    sql_set_max_rows_sl_d_60 = dce_pop_long(dce_table, "dce_result")
    Call dce_table_destroy(dce_table)
End Function


Function sl_d_60_worklist2&(spc_no$(), pt_no$(), gnl_item_cd$(), tst_frct_cd$(), tst_cd$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        dce_table = dce_submit("sl_d_60", "sl_d_60_worklist2", Socket)
    End If
    sl_d_60_worklist2 = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "spc_no", spc_no())
    rv& = dce_pop_array(dce_table, "pt_no", pt_no())
    rv& = dce_pop_array(dce_table, "gnl_item_cd", gnl_item_cd())
    rv& = dce_pop_array(dce_table, "tst_frct_cd", tst_frct_cd())
    rv& = dce_pop_array(dce_table, "tst_cd", tst_cd())
    Call dce_table_destroy(dce_table)
End Function


Function sl_d_60_worklist&(idates1$, idates2$, islip_no$, ihospital$, spc_no$(), pt_no$(), acpt_no$(), tst_frct_cd$(), tst_cd$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "idates1", idates1)
        rv& = dce_push_string(Socket, "idates2", idates2)
        rv& = dce_push_string(Socket, "islip_no", islip_no)
        rv& = dce_push_string(Socket, "ihospital", ihospital)
        dce_table = dce_submit("sl_d_60", "sl_d_60_worklist", Socket)
    End If
    sl_d_60_worklist = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "spc_no", spc_no())
    rv& = dce_pop_array(dce_table, "pt_no", pt_no())
    rv& = dce_pop_array(dce_table, "acpt_no", acpt_no())
    rv& = dce_pop_array(dce_table, "tst_frct_cd", tst_frct_cd())
    rv& = dce_pop_array(dce_table, "tst_cd", tst_cd())
    Call dce_table_destroy(dce_table)
End Function


Function sl_d_60_select&(spcno$, tst_cd$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "spcno", spcno)
        dce_table = dce_submit("sl_d_60", "sl_d_60_select", Socket)
    End If
    sl_d_60_select = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "tst_cd", tst_cd())
    Call dce_table_destroy(dce_table)
End Function


Function sel_order_total_select&(spc_no$(), tst_cd$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        dce_table = dce_submit("sl_d_60", "sel_order_total_select", Socket)
    End If
    sel_order_total_select = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "spc_no", spc_no())
    rv& = dce_pop_array(dce_table, "tst_cd", tst_cd())
    Call dce_table_destroy(dce_table)
End Function


Function sl_spcid_tstcd_select&(spcno$, tst_cd$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "spcno", spcno)
        dce_table = dce_submit("sl_d_60", "sl_spcid_tstcd_select", Socket)
    End If
    sl_spcid_tstcd_select = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "tst_cd", tst_cd())
    Call dce_table_destroy(dce_table)
End Function


Function sl_spcid_tstcd_select_an&(spcno$, tst_cd$(), an$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "spcno", spcno)
        dce_table = dce_submit("sl_d_60", "sl_spcid_tstcd_select_an", Socket)
    End If
    sl_spcid_tstcd_select_an = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "tst_cd", tst_cd())
    rv& = dce_pop_array(dce_table, "an", an())
    Call dce_table_destroy(dce_table)
End Function


Function sl_spcid_tstcd_select_an_day&(spcno$, pt_no$(), patname$(), tst_cd$(), an$(), day_yn$(), ord_site$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "spcno", spcno)
        dce_table = dce_submit("sl_d_60", "sl_spcid_tstcd_select_an_day", Socket)
    End If
    sl_spcid_tstcd_select_an_day = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "pt_no", pt_no())
    rv& = dce_pop_array(dce_table, "patname", patname())
    rv& = dce_pop_array(dce_table, "tst_cd", tst_cd())
    rv& = dce_pop_array(dce_table, "an", an())
    rv& = dce_pop_array(dce_table, "day_yn", day_yn())
    rv& = dce_pop_array(dce_table, "ord_site", ord_site())
    Call dce_table_destroy(dce_table)
End Function


Function sl_spcid_tstcd_select2&(i_equip_cd$, i_spc_no$, vtst_cd$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
        rv& = dce_push_string(Socket, "i_spc_no", i_spc_no)
        dce_table = dce_submit("sl_d_60", "sl_spcid_tstcd_select2", Socket)
    End If
    sl_spcid_tstcd_select2 = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "vtst_cd", vtst_cd())
    Call dce_table_destroy(dce_table)
End Function


Function sel_order_total_select2&(spc_no$(), tst_cd$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        dce_table = dce_submit("sl_d_60", "sel_order_total_select2", Socket)
    End If
    sel_order_total_select2 = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "spc_no", spc_no())
    rv& = dce_pop_array(dce_table, "tst_cd", tst_cd())
    Call dce_table_destroy(dce_table)
End Function


Function sl_d_60_user_info&(vuser$, vpasswd$, wkpers$(), nm$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "vuser", vuser)
        rv& = dce_push_string(Socket, "vpasswd", vpasswd)
        dce_table = dce_submit("sl_d_60", "sl_d_60_user_info", Socket)
    End If
    sl_d_60_user_info = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "wkpers", wkpers())
    rv& = dce_pop_array(dce_table, "nm", nm())
    Call dce_table_destroy(dce_table)
End Function


Function sl_d_60_sel_spcid_tstcd&(spcno$, pt_no$(), patname$(), tst_cd$(), an$(), day_yn$(), ord_site$(), tst_dte$(), SPC_CD_1$(), SPC_NM1$(), SPC_CD_2$(), SPC_NM2$(), tst_stat$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_60")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "spcno", spcno)
        dce_table = dce_submit("sl_d_60", "sl_d_60_sel_spcid_tstcd", Socket)
    End If
    sl_d_60_sel_spcid_tstcd = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "pt_no", pt_no())
    rv& = dce_pop_array(dce_table, "patname", patname())
    rv& = dce_pop_array(dce_table, "tst_cd", tst_cd())
    rv& = dce_pop_array(dce_table, "an", an())
    rv& = dce_pop_array(dce_table, "day_yn", day_yn())
    rv& = dce_pop_array(dce_table, "ord_site", ord_site())
    rv& = dce_pop_array(dce_table, "tst_dte", tst_dte())
    rv& = dce_pop_array(dce_table, "SPC_CD_1", SPC_CD_1())
    rv& = dce_pop_array(dce_table, "SPC_NM1", SPC_NM1())
    rv& = dce_pop_array(dce_table, "SPC_CD_2", SPC_CD_2())
    rv& = dce_pop_array(dce_table, "SPC_NM2", SPC_NM2())
    rv& = dce_pop_array(dce_table, "tst_stat", tst_stat())
    Call dce_table_destroy(dce_table)
End Function




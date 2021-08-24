Attribute VB_Name = "sl_d_m01_c"
Option Explicit

Dim rv&

Public Function sql_prepare_sl_d_m01&(db$, login$, pwd$)
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_m01")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "db", db)
        rv& = dce_push_string(Socket, "login", login)
        rv& = dce_push_string(Socket, "pwd", pwd)
        dce_table = dce_submit("sl_d_m01", "sql_prepare_sl_d_m01", Socket)
    End If
    sql_prepare_sl_d_m01 = dce_pop_long(dce_table, "dce_result")
    Call dce_table_destroy(dce_table)
End Function


Public Function sql_rows_sl_d_m01&(maxrows&)
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_m01")
    If (Socket > -1) Then
        Call dce_push_long(Socket, "maxrows", maxrows)
        dce_table = dce_submit("sl_d_m01", "sql_rows_sl_d_m01", Socket)
    End If
    sql_rows_sl_d_m01 = dce_pop_long(dce_table, "dce_result")
    Call dce_table_destroy(dce_table)
End Function


Public Function sql_set_max_rows_sl_d_m01&(maxrows&)
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_m01")
    If (Socket > -1) Then
        Call dce_push_long(Socket, "maxrows", maxrows)
        dce_table = dce_submit("sl_d_m01", "sql_set_max_rows_sl_d_m01", Socket)
    End If
    sql_set_max_rows_sl_d_m01 = dce_pop_long(dce_table, "dce_result")
    Call dce_table_destroy(dce_table)
End Function


Public Function sl_sysdate_select&(v_date$(), v_date_8$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_m01")
    If (Socket > -1) Then
        dce_table = dce_submit("sl_d_m01", "sl_sysdate_select", Socket)
    End If
    sl_sysdate_select = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "v_date", v_date())
    rv& = dce_pop_array(dce_table, "v_date_8", v_date_8())
    Call dce_table_destroy(dce_table)
End Function


Public Function sl_sel_spcno_tstcd_all&(i_spc_no$, i_equip_cd$, v_spc_no$(), v_pt_no$(), v_pt_nm$(), v_tst_cd$(), v_tst_nm$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_m01")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "i_spc_no", i_spc_no)
        rv& = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
        dce_table = dce_submit("sl_d_m01", "sl_sel_spcno_tstcd_all", Socket)
    End If
    sl_sel_spcno_tstcd_all = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "v_spc_no", v_spc_no())
    rv& = dce_pop_array(dce_table, "v_pt_no", v_pt_no())
    rv& = dce_pop_array(dce_table, "v_pt_nm", v_pt_nm())
    rv& = dce_pop_array(dce_table, "v_tst_cd", v_tst_cd())
    rv& = dce_pop_array(dce_table, "v_tst_nm", v_tst_nm())
    Call dce_table_destroy(dce_table)
End Function


Public Function sl_sel_spcno_tstcd_all_sub&(i_spc_no$, i_equip_cd$, v_spc_no$(), v_pt_no$(), v_pt_nm$(), v_tst_frct_cd$(), v_tst_frct_nm$(), v_acpt_dte$(), v_acpt_no$(), v_sex$(), v_age$(), v_spc_cd$(), v_spc_nm$(), v_tst_cd$(), v_tst_nm$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_m01")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "i_spc_no", i_spc_no)
        rv& = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
        dce_table = dce_submit("sl_d_m01", "sl_sel_spcno_tstcd_all_sub", Socket)
    End If
    sl_sel_spcno_tstcd_all_sub = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "v_spc_no", v_spc_no())
    rv& = dce_pop_array(dce_table, "v_pt_no", v_pt_no())
    rv& = dce_pop_array(dce_table, "v_pt_nm", v_pt_nm())
    rv& = dce_pop_array(dce_table, "v_tst_frct_cd", v_tst_frct_cd())
    rv& = dce_pop_array(dce_table, "v_tst_frct_nm", v_tst_frct_nm())
    rv& = dce_pop_array(dce_table, "v_acpt_dte", v_acpt_dte())
    rv& = dce_pop_array(dce_table, "v_acpt_no", v_acpt_no())
    rv& = dce_pop_array(dce_table, "v_sex", v_sex())
    rv& = dce_pop_array(dce_table, "v_age", v_age())
    rv& = dce_pop_array(dce_table, "v_spc_cd", v_spc_cd())
    rv& = dce_pop_array(dce_table, "v_spc_nm", v_spc_nm())
    rv& = dce_pop_array(dce_table, "v_tst_cd", v_tst_cd())
    rv& = dce_pop_array(dce_table, "v_tst_nm", v_tst_nm())
    Call dce_table_destroy(dce_table)
End Function


Public Function sl_sel_spcno_tstcd_unin&(i_spc_no$, i_equip_cd$, v_spc_no$(), v_pt_no$(), v_pt_nm$(), v_tst_cd$(), v_tst_nm$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_m01")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "i_spc_no", i_spc_no)
        rv& = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
        dce_table = dce_submit("sl_d_m01", "sl_sel_spcno_tstcd_unin", Socket)
    End If
    sl_sel_spcno_tstcd_unin = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "v_spc_no", v_spc_no())
    rv& = dce_pop_array(dce_table, "v_pt_no", v_pt_no())
    rv& = dce_pop_array(dce_table, "v_pt_nm", v_pt_nm())
    rv& = dce_pop_array(dce_table, "v_tst_cd", v_tst_cd())
    rv& = dce_pop_array(dce_table, "v_tst_nm", v_tst_nm())
    Call dce_table_destroy(dce_table)
End Function


Public Function sl_sel_spcno_tstcd_unin_sub&(i_spc_no$, i_equip_cd$, v_spc_no$(), v_pt_no$(), v_pt_nm$(), v_tst_frct_cd$(), v_tst_frct_nm$(), v_acpt_dte$(), v_acpt_no$(), v_sex$(), v_age$(), v_spc_cd$(), v_spc_nm$(), v_tst_cd$(), v_tst_nm$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_m01")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "i_spc_no", i_spc_no)
        rv& = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
        dce_table = dce_submit("sl_d_m01", "sl_sel_spcno_tstcd_unin_sub", Socket)
    End If
    sl_sel_spcno_tstcd_unin_sub = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "v_spc_no", v_spc_no())
    rv& = dce_pop_array(dce_table, "v_pt_no", v_pt_no())
    rv& = dce_pop_array(dce_table, "v_pt_nm", v_pt_nm())
    rv& = dce_pop_array(dce_table, "v_tst_frct_cd", v_tst_frct_cd())
    rv& = dce_pop_array(dce_table, "v_tst_frct_nm", v_tst_frct_nm())
    rv& = dce_pop_array(dce_table, "v_acpt_dte", v_acpt_dte())
    rv& = dce_pop_array(dce_table, "v_acpt_no", v_acpt_no())
    rv& = dce_pop_array(dce_table, "v_sex", v_sex())
    rv& = dce_pop_array(dce_table, "v_age", v_age())
    rv& = dce_pop_array(dce_table, "v_spc_cd", v_spc_cd())
    rv& = dce_pop_array(dce_table, "v_spc_nm", v_spc_nm())
    rv& = dce_pop_array(dce_table, "v_tst_cd", v_tst_cd())
    rv& = dce_pop_array(dce_table, "v_tst_nm", v_tst_nm())
    Call dce_table_destroy(dce_table)
End Function


Public Function sl_sel_acptno_tstcd_vitek&(i_tst_frct_cd$, i_acpt_dte$, i_acpt_no$, i_equip_cd$, i_spc_no$, v_spc_no$(), v_pt_no$(), v_pt_nm$(), v_tst_frct_cd$(), v_tst_frct_nm$(), v_acpt_dte$(), v_acpt_no$(), v_sex$(), v_age$(), v_birthday$(), v_ord_site$(), v_dept$(), v_spc_cd$(), v_spc_nm$(), v_tst_cd$(), v_tst_nm$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_m01")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "i_tst_frct_cd", i_tst_frct_cd)
        rv& = dce_push_string(Socket, "i_acpt_dte", i_acpt_dte)
        rv& = dce_push_string(Socket, "i_acpt_no", i_acpt_no)
        rv& = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
        rv& = dce_push_string(Socket, "i_spc_no", i_spc_no)
        dce_table = dce_submit("sl_d_m01", "sl_sel_acptno_tstcd_vitek", Socket)
    End If
    sl_sel_acptno_tstcd_vitek = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "v_spc_no", v_spc_no())
    rv& = dce_pop_array(dce_table, "v_pt_no", v_pt_no())
    rv& = dce_pop_array(dce_table, "v_pt_nm", v_pt_nm())
    rv& = dce_pop_array(dce_table, "v_tst_frct_cd", v_tst_frct_cd())
    rv& = dce_pop_array(dce_table, "v_tst_frct_nm", v_tst_frct_nm())
    rv& = dce_pop_array(dce_table, "v_acpt_dte", v_acpt_dte())
    rv& = dce_pop_array(dce_table, "v_acpt_no", v_acpt_no())
    rv& = dce_pop_array(dce_table, "v_sex", v_sex())
    rv& = dce_pop_array(dce_table, "v_age", v_age())
    rv& = dce_pop_array(dce_table, "v_birthday", v_birthday())
    rv& = dce_pop_array(dce_table, "v_ord_site", v_ord_site())
    rv& = dce_pop_array(dce_table, "v_dept", v_dept())
    rv& = dce_pop_array(dce_table, "v_spc_cd", v_spc_cd())
    rv& = dce_pop_array(dce_table, "v_spc_nm", v_spc_nm())
    rv& = dce_pop_array(dce_table, "v_tst_cd", v_tst_cd())
    rv& = dce_pop_array(dce_table, "v_tst_nm", v_tst_nm())
    Call dce_table_destroy(dce_table)
End Function


Public Function sl_sel_machine_id&(i_equip_cd$, machine_id$(), equip_cd$(), equip_nm$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_d_m01")
    If (Socket > -1) Then
        rv& = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
        dce_table = dce_submit("sl_d_m01", "sl_sel_machine_id", Socket)
    End If
    sl_sel_machine_id = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_array(dce_table, "machine_id", machine_id())
    rv& = dce_pop_array(dce_table, "equip_cd", equip_cd())
    rv& = dce_pop_array(dce_table, "equip_nm", equip_nm())
    Call dce_table_destroy(dce_table)
End Function




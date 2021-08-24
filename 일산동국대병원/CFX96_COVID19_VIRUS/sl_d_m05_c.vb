Attribute VB_Name = "sl_d_m05"
Function sql_prepare_sl_d_m05&(db$, login$, pwd$)
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "db", db)
                rv% = dce_push_string(Socket, "login", login)
                rv% = dce_push_string(Socket, "pwd", pwd)
                dce_table = dce_submit("sl_d_m05", "sql_prepare_sl_d_m05", Socket)
        End If
        sql_prepare_sl_d_m05 = dce_pop_long(dce_table, "dce_result")
        Call dce_table_destroy(dce_table)
End Function


Function sql_rows_sl_d_m05&(maxrows&)
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                Call dce_push_long(Socket, "maxrows", maxrows)
                dce_table = dce_submit("sl_d_m05", "sql_rows_sl_d_m05", Socket)
        End If
        sql_rows_sl_d_m05 = dce_pop_long(dce_table, "dce_result")
        Call dce_table_destroy(dce_table)
End Function


Function sql_set_max_rows_sl_d_m05&(maxrows&)
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                Call dce_push_long(Socket, "maxrows", maxrows)
                dce_table = dce_submit("sl_d_m05", "sql_set_max_rows_sl_d_m05", Socket)
        End If
        sql_set_max_rows_sl_d_m05 = dce_pop_long(dce_table, "dce_result")
        Call dce_table_destroy(dce_table)
End Function


Function sl_spcid_workdate_select&(iacptdate$, ihospital$, islipcode$, imachine$, pt_no$(), ptnm$(), gnl_item_cd$(), spc_no$(), tst_cd$(), eqcode$(), examname$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "iacptdate", iacptdate)
                rv% = dce_push_string(Socket, "ihospital", ihospital)
                rv% = dce_push_string(Socket, "islipcode", islipcode)
                rv% = dce_push_string(Socket, "imachine", imachine)
                dce_table = dce_submit("sl_d_m05", "sl_spcid_workdate_select", Socket)
        End If
        sl_spcid_workdate_select = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "pt_no", pt_no())
        rv% = dce_pop_array(dce_table, "ptnm", ptnm())
        rv% = dce_pop_array(dce_table, "gnl_item_cd", gnl_item_cd())
        rv% = dce_pop_array(dce_table, "spc_no", spc_no())
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        rv% = dce_pop_array(dce_table, "eqcode", eqcode())
        rv% = dce_pop_array(dce_table, "examname", examname())
        Call dce_table_destroy(dce_table)
End Function


Function sl_spcid_workdate_code&(iacptdate$, ihospital$, itstcd$, imachine$, pt_no$(), ptnm$(), gnl_item_cd$(), spc_no$(), tst_cd$(), eqcode$(), examname$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "iacptdate", iacptdate)
                rv% = dce_push_string(Socket, "ihospital", ihospital)
                rv% = dce_push_string(Socket, "itstcd", itstcd)
                rv% = dce_push_string(Socket, "imachine", imachine)
                dce_table = dce_submit("sl_d_m05", "sl_spcid_workdate_code", Socket)
        End If
        sl_spcid_workdate_code = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "pt_no", pt_no())
        rv% = dce_pop_array(dce_table, "ptnm", ptnm())
        rv% = dce_pop_array(dce_table, "gnl_item_cd", gnl_item_cd())
        rv% = dce_pop_array(dce_table, "spc_no", spc_no())
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        rv% = dce_pop_array(dce_table, "eqcode", eqcode())
        rv% = dce_pop_array(dce_table, "examname", examname())
        Call dce_table_destroy(dce_table)
End Function


Function sl_spcid_workdate_InOut&(iacptdate$, ihospital$, itstcd$, imachine$, ipatsect$, pt_no$(), ptnm$(), gnl_item_cd$(), spc_no$(), tst_cd$(), eqcode$(), examname$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "iacptdate", iacptdate)
                rv% = dce_push_string(Socket, "ihospital", ihospital)
                rv% = dce_push_string(Socket, "itstcd", itstcd)
                rv% = dce_push_string(Socket, "imachine", imachine)
                rv% = dce_push_string(Socket, "ipatsect", ipatsect)
                dce_table = dce_submit("sl_d_m05", "sl_spcid_workdate_InOut", Socket)
        End If
        sl_spcid_workdate_InOut = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "pt_no", pt_no())
        rv% = dce_pop_array(dce_table, "ptnm", ptnm())
        rv% = dce_pop_array(dce_table, "gnl_item_cd", gnl_item_cd())
        rv% = dce_pop_array(dce_table, "spc_no", spc_no())
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        rv% = dce_pop_array(dce_table, "eqcode", eqcode())
        rv% = dce_pop_array(dce_table, "examname", examname())
        Call dce_table_destroy(dce_table)
End Function


Function sl_spcid_tstcd_select_new&(SpcNo$, eqcode$(), examname$(), tst_cd$(), pt_no$(), ptnm$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "spcno", SpcNo)
                dce_table = dce_submit("sl_d_m05", "sl_spcid_tstcd_select_new", Socket)
        End If
        sl_spcid_tstcd_select = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "eqcode", eqcode())
        rv% = dce_pop_array(dce_table, "examname", examname())
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        rv% = dce_pop_array(dce_table, "pt_no", pt_no())
        rv% = dce_pop_array(dce_table, "ptnm", ptnm())
        Call dce_table_destroy(dce_table)
End Function


Function sel_order_Total_select&(SpcNo$, vmachcode$, eqcode$(), examname$(), tst_cd$(), pt_no$(), ptnm$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "spcno", SpcNo)
                rv% = dce_push_string(Socket, "vmachcode", vmachcode)
                dce_table = dce_submit("sl_d_m05", "sel_order_Total_select", Socket)
        End If
        sel_order_Total_select = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "eqcode", eqcode())
        rv% = dce_pop_array(dce_table, "examname", examname())
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        rv% = dce_pop_array(dce_table, "pt_no", pt_no())
        rv% = dce_pop_array(dce_table, "ptnm", ptnm())
        Call dce_table_destroy(dce_table)
End Function


Function sl_spcid_select&(iacptdate$, imachine$, pt_no$(), ptnm$(), spc_no$(), gnl_item_cd$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "iacptdate", iacptdate)
                rv% = dce_push_string(Socket, "imachine", imachine)
                dce_table = dce_submit("sl_d_m05", "sl_spcid_select", Socket)
        End If
        sl_spcid_select = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "pt_no", pt_no())
        rv% = dce_pop_array(dce_table, "ptnm", ptnm())
        rv% = dce_pop_array(dce_table, "spc_no", spc_no())
        rv% = dce_pop_array(dce_table, "gnl_item_cd", gnl_item_cd())
        Call dce_table_destroy(dce_table)
End Function


Function sl_areano_tstcd_select&(iacptdate$, islipcode$, ihospital$, iareano$, imachine$, pt_no$(), ptnm$(), gnl_item_cd$(), spc_no$(), tst_cd$(), eqcode$(), examname$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "iacptdate", iacptdate)
                rv% = dce_push_string(Socket, "islipcode", islipcode)
                rv% = dce_push_string(Socket, "ihospital", ihospital)
                rv% = dce_push_string(Socket, "iareano", iareano)
                rv% = dce_push_string(Socket, "imachine", imachine)
                dce_table = dce_submit("sl_d_m05", "sl_areano_tstcd_select", Socket)
        End If
        sl_areano_tstcd_select = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "pt_no", pt_no())
        rv% = dce_pop_array(dce_table, "ptnm", ptnm())
        rv% = dce_pop_array(dce_table, "gnl_item_cd", gnl_item_cd())
        rv% = dce_pop_array(dce_table, "spc_no", spc_no())
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        rv% = dce_pop_array(dce_table, "eqcode", eqcode())
        rv% = dce_pop_array(dce_table, "examname", examname())
        Call dce_table_destroy(dce_table)
End Function

'-- 원본
'Function sl_spcid_tstcd_select_qc&(SpcNo$, i_equip_cd$, tst_cd$())
'        Dim dce_table As Long, Socket As Integer
'
'        Call dce_checkver(2, 0)
'        Socket = dce_findserver("sl_d_m05")
'        If (Socket > -1) Then
'                rv% = dce_push_string(Socket, "spcno", SpcNo)
'                rv% = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
'                dce_table = dce_submit("sl_d_m05", "sl_spcid_tstcd_select_qc", Socket)
'        End If
'        sl_spcid_tstcd_select_qc = dce_pop_long(dce_table, "dce_result")
'        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
'        Call dce_table_destroy(dce_table)
'End Function

'-- 수정1
'Function sl_examdata_select&(ispc_no$, i_equip_cd$, pt_no$(), acpt_no$(), tst_cd$(), spc_cd$())
'        Dim dce_table As Long, Socket As Integer
'
'        Call dce_checkver(2, 0)
'        Socket = dce_findserver("sl_d_m05")
'        If (Socket > -1) Then
'                rv% = dce_push_string(Socket, "ispc_no", ispc_no)
'                rv% = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
'                dce_table = dce_submit("sl_d_m05", "sl_examdata_select", Socket)
'        End If
'        sl_examdata_select = dce_pop_long(dce_table, "dce_result")
'        rv% = dce_pop_array(dce_table, "pt_no", pt_no())
'        rv% = dce_pop_array(dce_table, "acpt_no", acpt_no())
'        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
'        rv% = dce_pop_array(dce_table, "spc_cd", spc_cd())
'        Call dce_table_destroy(dce_table)
'End Function

'-- 최종
Function sl_examdata_select&(ispc_no$, i_equip_cd$, eqcode$(), examname$(), tst_cd$(), pt_no$(), ptnm$(), acpt_no$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "ispc_no", ispc_no)
                rv% = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
                dce_table = dce_submit("sl_d_m05", "sl_examdata_select", Socket)
        End If
        sl_examdata_select = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "eqcode", eqcode())
        rv% = dce_pop_array(dce_table, "examname", examname())
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        rv% = dce_pop_array(dce_table, "pt_no", pt_no())
        rv% = dce_pop_array(dce_table, "ptnm", ptnm())
        rv% = dce_pop_array(dce_table, "acpt_no", acpt_no())
        Call dce_table_destroy(dce_table)
End Function

Function sl_spcid_workdate_s_InOut&(iacptdate$, itstcd1$, itstcd2$, itstcd3$, itstcd4$, imachine$, acpt_dte$(), spc_no$(), tst_cd$(), eqcode$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "iacptdate", iacptdate)
                rv% = dce_push_string(Socket, "itstcd1", itstcd1)
                rv% = dce_push_string(Socket, "itstcd2", itstcd2)
                rv% = dce_push_string(Socket, "itstcd3", itstcd3)
                rv% = dce_push_string(Socket, "itstcd4", itstcd4)
                rv% = dce_push_string(Socket, "imachine", imachine)
                dce_table = dce_submit("sl_d_m05", "sl_spcid_workdate_s_InOut", Socket)
        End If
        sl_spcid_workdate_s_InOut = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "acpt_dte", acpt_dte())
        rv% = dce_pop_array(dce_table, "spc_no", spc_no())
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        rv% = dce_pop_array(dce_table, "eqcode", eqcode())
        Call dce_table_destroy(dce_table)
End Function


Function sl_areano_search_spcid&(iacptdate$, islipcode$, ihospital$, iareano$, imachine$, spc_no$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "iacptdate", iacptdate)
                rv% = dce_push_string(Socket, "islipcode", islipcode)
                rv% = dce_push_string(Socket, "ihospital", ihospital)
                rv% = dce_push_string(Socket, "iareano", iareano)
                rv% = dce_push_string(Socket, "imachine", imachine)
                dce_table = dce_submit("sl_d_m05", "sl_areano_search_spcid", Socket)
        End If
        sl_areano_search_spcid = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "SPC_NO", spc_no())
        Call dce_table_destroy(dce_table)
End Function




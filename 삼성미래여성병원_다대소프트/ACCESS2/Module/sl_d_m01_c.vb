Attribute VB_Name = "sl_d_m01"
Function sql_prepare_sl_d_m01&(db$, login$, pwd$)
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m01")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "db", db)
                rv% = dce_push_string(Socket, "login", login)
                rv% = dce_push_string(Socket, "pwd", pwd)
                dce_table = dce_submit("sl_d_m01", "sql_prepare_sl_d_m01", Socket)
        End If
        sql_prepare_sl_d_m01 = dce_pop_long(dce_table, "dce_result")
        Call dce_table_destroy(dce_table)
End Function


Function sql_rows_sl_d_m01&(maxrows&)
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


Function sql_set_max_rows_sl_d_m01&(maxrows&)
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


'Function sl_spcid_tstcd_select&(SpcNo$, tst_cd$())
'        Dim dce_table As Long, Socket As Integer
'
'        Call dce_checkver(2, 0)
'        Socket = dce_findserver("sl_d_m01")
'        If (Socket > -1) Then
'                rv% = dce_push_string(Socket, "spcno", SpcNo)
'                dce_table = dce_submit("sl_d_m01", "sl_spcid_tstcd_select", Socket)
'        End If
'        sl_spcid_tstcd_select = dce_pop_long(dce_table, "dce_result")
'        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
'        Call dce_table_destroy(dce_table)
'End Function


Function sl_areano_tstcd_select&(iacptdate$, islipcode$, ihospital$, iareano$, tst_cd$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m01")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "iacptdate", iacptdate)
                rv% = dce_push_string(Socket, "islipcode", islipcode)
                rv% = dce_push_string(Socket, "ihospital", ihospital)
                rv% = dce_push_string(Socket, "iareano", iareano)
                dce_table = dce_submit("sl_d_m01", "sl_areano_tstcd_select", Socket)
        End If
        sl_areano_tstcd_select = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        Call dce_table_destroy(dce_table)
End Function


Function sl_areano_pt_inf_select&(in_spc_no$, pt_no$(), patname$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m01")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "in_spc_no", in_spc_no)
                dce_table = dce_submit("sl_d_m01", "sl_areano_pt_inf_select", Socket)
        End If
        sl_areano_pt_inf_select = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "pt_no", pt_no())
        rv% = dce_pop_array(dce_table, "patname", patname())
        Call dce_table_destroy(dce_table)
End Function


Function sl_spcid_tstcd_select_qc&(i_equip_cd$, SpcNo$, tst_cd$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m01")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
                rv% = dce_push_string(Socket, "spcno", SpcNo)
                dce_table = dce_submit("sl_d_m01", "sl_spcid_tstcd_select_qc", Socket)
        End If
        sl_spcid_tstcd_select_qc = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        Call dce_table_destroy(dce_table)
End Function


Function sl_spcid_tstcd_select_qc1&(i_equip_cd$, i_level$, spc_no$(), tst_cd$(), level_cd$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m01")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
                rv% = dce_push_string(Socket, "i_level", i_level)
                dce_table = dce_submit("sl_d_m01", "sl_spcid_tstcd_select_qc1", Socket)
        End If
        sl_spcid_tstcd_select_qc1 = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "spc_no", spc_no())
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        rv% = dce_pop_array(dce_table, "level_cd", level_cd())
        Call dce_table_destroy(dce_table)
End Function


Function sl_xe2100_examdata_select&(ispc_no$, i_equip_cd$, pt_no$(), acpt_no$(), tst_cd$(), spc_cd$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m01")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "ispc_no", ispc_no)
                rv% = dce_push_string(Socket, "i_equip_cd", i_equip_cd)
                dce_table = dce_submit("sl_d_m01", "sl_xe2100_examdata_select", Socket)
        End If
        sl_xe2100_examdata_select = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "pt_no", pt_no())
        rv% = dce_pop_array(dce_table, "acpt_no", acpt_no())
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        rv% = dce_pop_array(dce_table, "spc_cd", spc_cd())
        Call dce_table_destroy(dce_table)
End Function


Function sl_spc_Elecsys2010_channel&(eqcode$(), examcode$(), examname$(), normal1$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m01")
        If (Socket > -1) Then
                dce_table = dce_submit("sl_d_m01", "sl_spc_Elecsys2010_channel", Socket)
        End If
        sl_spc_Elecsys2010_channel = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "eqcode", eqcode())
        rv% = dce_pop_array(dce_table, "examcode", examcode())
        rv% = dce_pop_array(dce_table, "examname", examname())
        rv% = dce_pop_array(dce_table, "normal1", normal1())
        Call dce_table_destroy(dce_table)
End Function



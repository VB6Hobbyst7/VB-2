Function sql_prepare_sl_d_m05& (db$,login$,pwd$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_m05")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"db",db)
		rv% = dce_push_string(Socket,"login",login)
		rv% = dce_push_string(Socket,"pwd",pwd)
		dce_table = dce_submit("sl_d_m05","sql_prepare_sl_d_m05",Socket)
	End If
	sql_prepare_sl_d_m05 = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function


Function sql_rows_sl_d_m05& (maxrows&)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_m05")
	If (Socket > -1) Then
		Call dce_push_long(Socket,"maxrows",maxrows)
		dce_table = dce_submit("sl_d_m05","sql_rows_sl_d_m05",Socket)
	End If
	sql_rows_sl_d_m05 = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function


Function sql_set_max_rows_sl_d_m05& (maxrows&)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_m05")
	If (Socket > -1) Then
		Call dce_push_long(Socket,"maxrows",maxrows)
		dce_table = dce_submit("sl_d_m05","sql_set_max_rows_sl_d_m05",Socket)
	End If
	sql_set_max_rows_sl_d_m05 = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function


Function sl_spcid_tstcd_select& (spcno$,tst_cd$(),pt_no$(),patname$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_m05")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"spcno",spcno)
		dce_table = dce_submit("sl_d_m05","sl_spcid_tstcd_select",Socket)
	End If
	sl_spcid_tstcd_select = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"patname",patname())
	Call dce_table_destroy(dce_table)
End function


Function sl_tstcd_spcid_select& (acptdte_in$,tstcd_in$,spc_no$(),pt_no$(),patname$(),tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_m05")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"acptdte_in",acptdte_in)
		rv% = dce_push_string(Socket,"tstcd_in",tstcd_in)
		dce_table = dce_submit("sl_d_m05","sl_tstcd_spcid_select",Socket)
	End If
	sl_tstcd_spcid_select = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"patname",patname())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sel_order_Total_select& (spc_no$(),tst_cd$(),pt_no$(),patname$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_m05")
	If (Socket > -1) Then
		dce_table = dce_submit("sl_d_m05","sel_order_Total_select",Socket)
	End If
	sel_order_Total_select = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"patname",patname())
	Call dce_table_destroy(dce_table)
End function


Function sl_list_select& (in_acpt_dte$,in_tst_cd$,a_pt_no$(),b_patname$(),a_tst_frct_cd$(),a_gnl_item_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_m05")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"in_acpt_dte",in_acpt_dte)
		rv% = dce_push_string(Socket,"in_tst_cd",in_tst_cd)
		dce_table = dce_submit("sl_d_m05","sl_list_select",Socket)
	End If
	sl_list_select = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"a_pt_no",a_pt_no())
	rv% = dce_pop_array(dce_table,"b_patname",b_patname())
	rv% = dce_pop_array(dce_table,"a_tst_frct_cd",a_tst_frct_cd())
	rv% = dce_pop_array(dce_table,"a_gnl_item_cd",a_gnl_item_cd())
	Call dce_table_destroy(dce_table)
End function



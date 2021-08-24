Function sql_prepare_sl_d_60& (db$,login$,pwd$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"db",db)
		rv% = dce_push_string(Socket,"login",login)
		rv% = dce_push_string(Socket,"pwd",pwd)
		dce_table = dce_submit("sl_d_60","sql_prepare_sl_d_60",Socket)
	End If
	sql_prepare_sl_d_60 = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function


Function sql_rows_sl_d_60& (maxrows&)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		Call dce_push_long(Socket,"maxrows",maxrows)
		dce_table = dce_submit("sl_d_60","sql_rows_sl_d_60",Socket)
	End If
	sql_rows_sl_d_60 = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function


Function sql_set_max_rows_sl_d_60& (maxrows&)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		Call dce_push_long(Socket,"maxrows",maxrows)
		dce_table = dce_submit("sl_d_60","sql_set_max_rows_sl_d_60",Socket)
	End If
	sql_set_max_rows_sl_d_60 = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_worklist2& (spc_no$(),pt_no$(),gnl_item_cd$(),tst_frct_cd$(),tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		dce_table = dce_submit("sl_d_60","sl_d_60_worklist2",Socket)
	End If
	sl_d_60_worklist2 = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"gnl_item_cd",gnl_item_cd())
	rv% = dce_pop_array(dce_table,"tst_frct_cd",tst_frct_cd())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_worklist& (idates1$,idates2$,islip_no$,ihospital$,spc_no$(),pt_no$(),acpt_no$(),tst_frct_cd$(),tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"idates1",idates1)
		rv% = dce_push_string(Socket,"idates2",idates2)
		rv% = dce_push_string(Socket,"islip_no",islip_no)
		rv% = dce_push_string(Socket,"ihospital",ihospital)
		dce_table = dce_submit("sl_d_60","sl_d_60_worklist",Socket)
	End If
	sl_d_60_worklist = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"acpt_no",acpt_no())
	rv% = dce_pop_array(dce_table,"tst_frct_cd",tst_frct_cd())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_select& (spcno$,tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"spcno",spcno)
		dce_table = dce_submit("sl_d_60","sl_d_60_select",Socket)
	End If
	sl_d_60_select = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sel_order_total_select& (spc_no$(),tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		dce_table = dce_submit("sl_d_60","sel_order_total_select",Socket)
	End If
	sel_order_total_select = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_spcid_tstcd_select& (spcno$,tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"spcno",spcno)
		dce_table = dce_submit("sl_d_60","sl_spcid_tstcd_select",Socket)
	End If
	sl_spcid_tstcd_select = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_sel_examcode& (idates1$,idates2$,iexamcode$,pt_no$(),patname$(),sex$(),age$(),spc_no$(),gnl_item_cd$(),bl_gth_dte$(),dept$(),wd_no$(),tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"idates1",idates1)
		rv% = dce_push_string(Socket,"idates2",idates2)
		rv% = dce_push_string(Socket,"iexamcode",iexamcode)
		dce_table = dce_submit("sl_d_60","sl_d_60_sel_examcode",Socket)
	End If
	sl_d_60_sel_examcode = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"patname",patname())
	rv% = dce_pop_array(dce_table,"sex",sex())
	rv% = dce_pop_array(dce_table,"age",age())
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"gnl_item_cd",gnl_item_cd())
	rv% = dce_pop_array(dce_table,"bl_gth_dte",bl_gth_dte())
	rv% = dce_pop_array(dce_table,"dept",dept())
	rv% = dce_pop_array(dce_table,"wd_no",wd_no())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_sel_examcode1& (idates1$,idates2$,iexamcode$,pt_no$(),patname$(),sex$(),age$(),spc_no$(),gnl_item_cd$(),bl_gth_dte$(),dept$(),wd_no$(),tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"idates1",idates1)
		rv% = dce_push_string(Socket,"idates2",idates2)
		rv% = dce_push_string(Socket,"iexamcode",iexamcode)
		dce_table = dce_submit("sl_d_60","sl_d_60_sel_examcode1",Socket)
	End If
	sl_d_60_sel_examcode1 = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"patname",patname())
	rv% = dce_pop_array(dce_table,"sex",sex())
	rv% = dce_pop_array(dce_table,"age",age())
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"gnl_item_cd",gnl_item_cd())
	rv% = dce_pop_array(dce_table,"bl_gth_dte",bl_gth_dte())
	rv% = dce_pop_array(dce_table,"dept",dept())
	rv% = dce_pop_array(dce_table,"wd_no",wd_no())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_sel_examcode_rsd800& (idates1$,idates2$,pt_no$(),patname$(),sex$(),age$(),spc_no$(),bl_gth_dte$(),dept$(),wd_no$(),infect$(),day_yn$(),er_yn$(),tst_cd$(),frct_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"idates1",idates1)
		rv% = dce_push_string(Socket,"idates2",idates2)
		dce_table = dce_submit("sl_d_60","sl_d_60_sel_examcode_rsd800",Socket)
	End If
	sl_d_60_sel_examcode_rsd800 = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"patname",patname())
	rv% = dce_pop_array(dce_table,"sex",sex())
	rv% = dce_pop_array(dce_table,"age",age())
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"bl_gth_dte",bl_gth_dte())
	rv% = dce_pop_array(dce_table,"dept",dept())
	rv% = dce_pop_array(dce_table,"wd_no",wd_no())
	rv% = dce_pop_array(dce_table,"infect",infect())
	rv% = dce_pop_array(dce_table,"day_yn",day_yn())
	rv% = dce_pop_array(dce_table,"er_yn",er_yn())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	rv% = dce_pop_array(dce_table,"frct_cd",frct_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_sel_spcno& (ispcno$,pt_no$(),patname$(),sex$(),age$(),gnl_item_cd$(),bl_gth_dte$(),dept$(),wd_no$(),tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"ispcno",ispcno)
		dce_table = dce_submit("sl_d_60","sl_d_60_sel_spcno",Socket)
	End If
	sl_d_60_sel_spcno = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"patname",patname())
	rv% = dce_pop_array(dce_table,"sex",sex())
	rv% = dce_pop_array(dce_table,"age",age())
	rv% = dce_pop_array(dce_table,"gnl_item_cd",gnl_item_cd())
	rv% = dce_pop_array(dce_table,"bl_gth_dte",bl_gth_dte())
	rv% = dce_pop_array(dce_table,"dept",dept())
	rv% = dce_pop_array(dce_table,"wd_no",wd_no())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_sel_spcno_new& (ispcno$,itststat$,pt_no$(),patname$(),sex$(),age$(),gnl_item_cd$(),bl_gth_dte$(),dept$(),wd_no$(),tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"ispcno",ispcno)
		rv% = dce_push_string(Socket,"itststat",itststat)
		dce_table = dce_submit("sl_d_60","sl_d_60_sel_spcno_new",Socket)
	End If
	sl_d_60_sel_spcno_new = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"patname",patname())
	rv% = dce_pop_array(dce_table,"sex",sex())
	rv% = dce_pop_array(dce_table,"age",age())
	rv% = dce_pop_array(dce_table,"gnl_item_cd",gnl_item_cd())
	rv% = dce_pop_array(dce_table,"bl_gth_dte",bl_gth_dte())
	rv% = dce_pop_array(dce_table,"dept",dept())
	rv% = dce_pop_array(dce_table,"wd_no",wd_no())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_sel_examcode_new& (idates1$,idates2$,iexamcode$,itststat$,pt_no$(),patname$(),sex$(),age$(),spc_no$(),gnl_item_cd$(),Acpt_dtetm$(),dept$(),wd_no$(),tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"idates1",idates1)
		rv% = dce_push_string(Socket,"idates2",idates2)
		rv% = dce_push_string(Socket,"iexamcode",iexamcode)
		rv% = dce_push_string(Socket,"itststat",itststat)
		dce_table = dce_submit("sl_d_60","sl_d_60_sel_examcode_new",Socket)
	End If
	sl_d_60_sel_examcode_new = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"patname",patname())
	rv% = dce_pop_array(dce_table,"sex",sex())
	rv% = dce_pop_array(dce_table,"age",age())
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"gnl_item_cd",gnl_item_cd())
	rv% = dce_pop_array(dce_table,"Acpt_dtetm",Acpt_dtetm())
	rv% = dce_pop_array(dce_table,"dept",dept())
	rv% = dce_pop_array(dce_table,"wd_no",wd_no())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_sel_examcode_new_micro& (idates1$,idates2$,iexamcode$,itststat$,pt_no$(),patname$(),sex$(),age$(),spc_no$(),gnl_item_cd$(),Acpt_dtetm$(),dept$(),wd_no$(),tst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"idates1",idates1)
		rv% = dce_push_string(Socket,"idates2",idates2)
		rv% = dce_push_string(Socket,"iexamcode",iexamcode)
		rv% = dce_push_string(Socket,"itststat",itststat)
		dce_table = dce_submit("sl_d_60","sl_d_60_sel_examcode_new_micro",Socket)
	End If
	sl_d_60_sel_examcode_new_micro = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"pt_no",pt_no())
	rv% = dce_pop_array(dce_table,"patname",patname())
	rv% = dce_pop_array(dce_table,"sex",sex())
	rv% = dce_pop_array(dce_table,"age",age())
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"gnl_item_cd",gnl_item_cd())
	rv% = dce_pop_array(dce_table,"Acpt_dtetm",Acpt_dtetm())
	rv% = dce_pop_array(dce_table,"dept",dept())
	rv% = dce_pop_array(dce_table,"wd_no",wd_no())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_sel_slxworkt_rslt& (iexamcode$,iptno$,idates$,spc_no$(),tst_dte$(),tst_cd$(),tst_rslt$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"iexamcode",iexamcode)
		rv% = dce_push_string(Socket,"iptno",iptno)
		rv% = dce_push_string(Socket,"idates",idates)
		dce_table = dce_submit("sl_d_60","sl_d_60_sel_slxworkt_rslt",Socket)
	End If
	sl_d_60_sel_slxworkt_rslt = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"tst_dte",tst_dte())
	rv% = dce_pop_array(dce_table,"tst_cd",tst_cd())
	rv% = dce_pop_array(dce_table,"tst_rslt",tst_rslt())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_ins_slxwrkep& (ispcno$,igrape$,islide$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"ispcno",ispcno)
		rv% = dce_push_string(Socket,"igrape",igrape)
		rv% = dce_push_string(Socket,"islide",islide)
		dce_table = dce_submit("sl_d_60","sl_d_60_ins_slxwrkep",Socket)
	End If
	sl_d_60_ins_slxwrkep = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_del_slxwrkep& (ispcno$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"ispcno",ispcno)
		dce_table = dce_submit("sl_d_60","sl_d_60_del_slxwrkep",Socket)
	End If
	sl_d_60_del_slxwrkep = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_sel_slxwrkep& (ispcno$,spc_no$(),tst_grape$(),tst_slide$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"ispcno",ispcno)
		dce_table = dce_submit("sl_d_60","sl_d_60_sel_slxwrkep",Socket)
	End If
	sl_d_60_sel_slxwrkep = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spc_no",spc_no())
	rv% = dce_pop_array(dce_table,"tst_grape",tst_grape())
	rv% = dce_pop_array(dce_table,"tst_slide",tst_slide())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_sel_spcno_qc& (i_equip_cd$,i_spc_no$,vtst_cd$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"i_equip_cd",i_equip_cd)
		rv% = dce_push_string(Socket,"i_spc_no",i_spc_no)
		dce_table = dce_submit("sl_d_60","sl_d_60_sel_spcno_qc",Socket)
	End If
	sl_d_60_sel_spcno_qc = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"vtst_cd",vtst_cd())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_60_ins_slxeqrdt& (in_date$,in_tst_cd$,in_tst_rslt$,in_EQUIP_CD$,in_SPC_NO$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_60")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"in_date",in_date)
		rv% = dce_push_string(Socket,"in_tst_cd",in_tst_cd)
		rv% = dce_push_string(Socket,"in_tst_rslt",in_tst_rslt)
		rv% = dce_push_string(Socket,"in_EQUIP_CD",in_EQUIP_CD)
		rv% = dce_push_string(Socket,"in_SPC_NO",in_SPC_NO)
		dce_table = dce_submit("sl_d_60","sl_d_60_ins_slxeqrdt",Socket)
	End If
	sl_d_60_ins_slxeqrdt = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function



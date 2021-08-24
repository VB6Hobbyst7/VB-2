unit smi_p_021_c;

interface 

uses SysUtils, Classes, ODET3020, Variants;

function smi_p_021_POCT_order_DC(sz_iEQUIP_CD: String;
	sz_iPT_NO: String;
	sz_iTEST_DTETM: String;
	var msg: String): LongInt;
function smi_p_021_POCT_insert_order(sz_iEQUIP_CD: String;
	sz_iPT_NO: String;
	sz_iTEST_DTETM: String;
	sz_iORD_SITE: String;
	sz_iORDCD: String;
	sz_iEDIT_ID: String;
	sz_iRESULT: String;
	var msg: String): LongInt;
function smi_p_021_update_ptno(sz_iEQUIP_CD: String;
	sz_iTEST_DTETM: String;
	sz_iNEW_PT_NO: String;
	var msg: String): LongInt;
function smi_p_021_POCT_insert_order_set(sz_iEQUIP_CD: String;
	sz_iPT_NO: String;
	sz_iTEST_DTETM: String;
	sz_iORD_SITE: String;
	sz_iORDCD: String;
	sz_iEDIT_ID: String;
	i_RSLT_CNT: Longint;
	sz_iRSLT_CDS: Variant;
	sz_iRESULTS: Variant;
	var msg: String): LongInt;


implementation

function smi_p_021_POCT_insert_order(sz_iEQUIP_CD: String;
	sz_iPT_NO: String;
	sz_iTEST_DTETM: String;
	sz_iORD_SITE: String;
	sz_iORDCD: String;
	sz_iEDIT_ID: String;
	sz_iRESULT: String;
	var msg: String): LongInt;
var
	dce_table : PTable;
	socket : Integer;
	rv     : Integer;
begin
	dce_table := nil;
	dce_checkver(2,0);
	socket := dce_findserver('smi_p_021');
	if (socket > -1) then begin
		dce_push_char_NStr(socket,'sz_iEQUIP_CD',sz_iEQUIP_CD);
		dce_push_char_NStr(socket,'sz_iPT_NO',sz_iPT_NO);
		dce_push_char_NStr(socket,'sz_iTEST_DTETM',sz_iTEST_DTETM);
		dce_push_char_NStr(socket,'sz_iORD_SITE',sz_iORD_SITE);
		dce_push_char_NStr(socket,'sz_iORDCD',sz_iORDCD);
		dce_push_char_NStr(socket,'sz_iEDIT_ID',sz_iEDIT_ID);
		dce_push_char_NStr(socket,'sz_iRESULT',sz_iRESULT);
		dce_table := dce_submit('smi_p_021','smi_p_021_POCT_insert_order',socket);
	end;
	smi_p_021_POCT_insert_order := dce_pop_long(dce_table,'dce_result');
	dce_pop_char_NStr(dce_table,'msg',msg);
	dce_table_destroy(dce_table);
end;

function smi_p_021_POCT_order_DC(sz_iEQUIP_CD: String;
	sz_iPT_NO: String;
	sz_iTEST_DTETM: String;
	var msg: String): LongInt;
var
	dce_table : PTable;
	socket : Integer;
	rv     : Integer;
begin
	dce_table := nil;
	dce_checkver(2,0);
	socket := dce_findserver('smi_p_021');
	if (socket > -1) then begin
		dce_push_char_NStr(socket,'sz_iEQUIP_CD',sz_iEQUIP_CD);
		dce_push_char_NStr(socket,'sz_iPT_NO',sz_iPT_NO);
		dce_push_char_NStr(socket,'sz_iTEST_DTETM',sz_iTEST_DTETM);
		dce_table := dce_submit('smi_p_021','smi_p_021_POCT_order_DC',socket);
	end;
	smi_p_021_POCT_order_DC := dce_pop_long(dce_table,'dce_result');
	dce_pop_char_NStr(dce_table,'msg',msg);
	dce_table_destroy(dce_table);
end;

function smi_p_021_update_ptno(sz_iEQUIP_CD: String;
	sz_iTEST_DTETM: String;
	sz_iNEW_PT_NO: String;
	var msg: String): LongInt;
var
	dce_table : PTable;
	socket : Integer;
	rv     : Integer;
begin
	dce_table := nil;
	dce_checkver(2,0);
	socket := dce_findserver('smi_p_021');
	if (socket > -1) then begin
		dce_push_char_NStr(socket,'sz_iEQUIP_CD',sz_iEQUIP_CD);
		dce_push_char_NStr(socket,'sz_iTEST_DTETM',sz_iTEST_DTETM);
		dce_push_char_NStr(socket,'sz_iNEW_PT_NO',sz_iNEW_PT_NO);
		dce_table := dce_submit('smi_p_021','smi_p_021_update_ptno',socket);
	end;
	smi_p_021_update_ptno := dce_pop_long(dce_table,'dce_result');
	dce_pop_char_NStr(dce_table,'msg',msg);
	dce_table_destroy(dce_table);
end;

function smi_p_021_POCT_insert_order_set(sz_iEQUIP_CD: String;
	sz_iPT_NO: String;
	sz_iTEST_DTETM: String;
	sz_iORD_SITE: String;
	sz_iORDCD: String;
	sz_iEDIT_ID: String;
	i_RSLT_CNT: Longint;
	sz_iRSLT_CDS: Variant;
	sz_iRESULTS: Variant;
	var msg: String): LongInt;
var
	dce_table : PTable;
	socket : Integer;
	rv     : Integer;
begin
	dce_table := nil;
	dce_checkver(2,0);
	socket := dce_findserver('smi_p_021');
	if (socket > -1) then begin
		dce_push_char_NStr(socket,'sz_iEQUIP_CD',sz_iEQUIP_CD);
		dce_push_char_NStr(socket,'sz_iPT_NO',sz_iPT_NO);
		dce_push_char_NStr(socket,'sz_iTEST_DTETM',sz_iTEST_DTETM);
		dce_push_char_NStr(socket,'sz_iORD_SITE',sz_iORD_SITE);
		dce_push_char_NStr(socket,'sz_iORDCD',sz_iORDCD);
		dce_push_char_NStr(socket,'sz_iEDIT_ID',sz_iEDIT_ID);
		dce_push_long(socket,'i_RSLT_CNT',i_RSLT_CNT);
		rv := dce_push_Narr(socket,'sz_iRSLT_CDS',sz_iRSLT_CDS);
		rv := dce_push_Narr(socket,'sz_iRESULTS',sz_iRESULTS);
		dce_table := dce_submit('smi_p_021','smi_p_021_POCT_insert_order_set',socket);
	end;
	smi_p_021_POCT_insert_order_set := dce_pop_long(dce_table,'dce_result');
	dce_pop_char_NStr(dce_table,'msg',msg);
	dce_table_destroy(dce_table);
end;

end.


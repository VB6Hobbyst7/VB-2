  select sLabws_momu, slabws_scnt 
  From AspsLabws                  
  where sLabws_cnt = 55      
    and sLabws_date = '20120423' 
    and slabws_momu in ('C2200','C2210','B2570','HT AST','HT ALT','B2580','B2602','B2710','HT R-GTP','C2411','HT CHOL','HT GLU','C3711','C3720','C3730','HTU-PRO','C3750','C3780','B2590','C3721','C2443','HT TG','C2420','HT HDL-C','B2630','B2611')
    and sLabws_hid  = 150000      

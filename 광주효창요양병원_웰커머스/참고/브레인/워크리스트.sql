 Select distinct a.slabw_date                              
               , a.slabw_cnt                               
               , a.slabw_cham                              
               , c.cham_whanja                             
 from AspCham C, AspsLabw a, AspsLabws  b                  
 Where b.slabws_date BETWEEN  '20120423' AND '20120423'    
   and b.slabws_momu in ('C2200','C2210','B2570','HT AST','HT ALT','B2580','B2602','B2710','HT R-GTP','C2411','HT CHOL','HT GLU','C3711','C3720','C3730','HTU-PRO','C3750','C3780','B2590','C3721','C2443','HT TG','C2420','HT HDL-C','B2630','B2611')
 and a.slabw_status in (0,1) 
   and a.slabw_hid = 150000                                
   and a.slabw_hid  = b.slabws_hid                         
   and a.slabw_date = b.slabws_date                        
   and a.slabw_cnt  = b.slabws_cnt                         
   and c.cham_Key   = a.slabw_cham                         
   Order by a.slabw_date , a.slabw_cnt                     

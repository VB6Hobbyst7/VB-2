 Select distinct a.slabw_date                              
               , a.slabw_cnt                               
               , a.slabw_cham                              
               , c.cham_whanja                             
 from AspCham C, AspsLabw a, AspsLabws  b                  
 Where b.slabws_date BETWEEN  '20120421' AND '20120423'    
   and b.slabws_momu in ('B1050     ','B1040','B1010','B1020','B1060','GR','LYMPHO','MONO','XLA0019','B1220','XLA0017')
 and a.slabw_status = 1 
   and a.slabw_hid = 150000                                
   and a.slabw_hid  = b.slabws_hid                         
   and a.slabw_date = b.slabws_date                        
   and a.slabw_cnt  = b.slabws_cnt                         
   and c.cham_Key   = a.slabw_cham                         
   Order by a.slabw_date , a.slabw_cnt                     



  select *
  From AspsLabws                  
  where sLabws_cnt = 37      
    and sLabws_date = '20120423' 
   -- and slabws_momu in ('B1050','B1040','B1010','B1020','B1060','GR','LYMPHO','MONO','XLA0019','B1220','XLA0017')
    and sLabws_hid  = 150000      

 Select sLabws_result From AspsLabws    
 where (sLabws_result = 0) 
    and sLabws_date = '20120423'    
    and sLabws_cnt  = 15        
    and sLabws_Hid  = 150000         

select * from aspslab
where slab_key = 'CBC/DIFF '

select * from aspslabs
where slabs_key = 'HT JOB   '

select * from aspslabs
where slabs_momu = 'HT AST'
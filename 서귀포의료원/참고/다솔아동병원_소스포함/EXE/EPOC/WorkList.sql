 SELECT distinct H141_TSAMPLENO 샘플번호    
 　　 , H141_TAKEDAT 수거일자      
      , H141_TAKETM 수거시간       
      , H141_TAKESEQ 수거순번      
      , H141_CHARTNO 차트번호      
      , FN_PATIENT_INFO(H141_CHARTNO) 환자성명              
      , FN_PATIENT_INFO(H141_CHARTNO, 'B') 생년월일       
      , FN_SEXAGE(H141_CHARTNO) 성별나이                    
      , H141_VISTDAT 방문일자                               
      , H141_ODRDAT 처방일자                                
      , H141_ODRNO 처방번호                                 
      , H141_ODRSEQ 처방서브번호                            
      , H141_RSLTYN 결과유무                                
      , H141_NOTYYN 통보유무                                
      , H141_SPECCD                                         
 FROM TB_H141_LISTAKEBODY                                   
    , TB_H131_SPPRESULT                                     
 WHERE H141_TAKEDAT between '20151004'                          And '20151005'    
   AND H141_SUGACD in ('LSC381501','LSC381502','LSC381503','LSC381507','LSC381510','LSC381505','LSC381506','LSC381512','C3791-1','LSC381513','C3792-1','LSC381514','LSC381515','C3710','C3850','C3796','LSC381509')
   AND NVL(H141_RSLTYN,' ') IN('N', 'T')                         
   AND H141_SEQNO = H131_SEQNO                                         
   AND (TRIM(H131_RESULT) IS NULL OR TRIM(H131_RESULT) = '결과대기') 
   AND H131_SPPTYPE = 'L010'                            
 Order by 수거일자, 수거시간 

 SELECT distinct H141_TSAMPLENO ���ù�ȣ    
 ���� , H141_TAKEDAT ��������      
      , H141_TAKETM ���Žð�       
      , H141_TAKESEQ ���ż���      
      , H141_CHARTNO ��Ʈ��ȣ      
      , FN_PATIENT_INFO(H141_CHARTNO) ȯ�ڼ���              
      , FN_PATIENT_INFO(H141_CHARTNO, 'B') �������       
      , FN_SEXAGE(H141_CHARTNO) ��������                    
      , H141_VISTDAT �湮����                               
      , H141_ODRDAT ó������                                
      , H141_ODRNO ó���ȣ                                 
      , H141_ODRSEQ ó�漭���ȣ                            
      , H141_RSLTYN �������                                
      , H141_NOTYYN �뺸����                                
      , H141_SPECCD                                         
 FROM TB_H141_LISTAKEBODY                                   
    , TB_H131_SPPRESULT                                     
 WHERE H141_TAKEDAT between '20151004'                          And '20151005'    
   AND H141_SUGACD in ('LSC381501','LSC381502','LSC381503','LSC381507','LSC381510','LSC381505','LSC381506','LSC381512','C3791-1','LSC381513','C3792-1','LSC381514','LSC381515','C3710','C3850','C3796','LSC381509')
   AND NVL(H141_RSLTYN,' ') IN('N', 'T')                         
   AND H141_SEQNO = H131_SEQNO                                         
   AND (TRIM(H131_RESULT) IS NULL OR TRIM(H131_RESULT) = '������') 
   AND H131_SPPTYPE = 'L010'                            
 Order by ��������, ���Žð� 

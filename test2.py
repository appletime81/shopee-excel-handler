row = 2
a = '=IF(R2="作業","障礙", IF(R2="障礙","障礙", IF(R2="抗爭","抗爭", IF(R2="40055重大障礙","40055重大障礙", IF(R2="非TWM問題的障礙","非TWM問題的障礙", IF(U2=35806,"非TWM問題的障礙", IF( OR(AND(AJ2<>"",AJ2>0,AJ2<0.7), AND(AK2<>"",AK2>0,AK2<0.7), AND(AL2<>"",AL2>0,AL2<0.7), AND(AP2<>"",AP2>0,AP2<0.7), AND(AQ2<>"",AQ2>0,AQ2<0.7), AND(AR2<>"",AR2>0,AR2<0.7), AND(AS2<>"",AS2>0,AS2<0.7), AND(AT2<>"",AT2>0,AT2<0.7), AND(AU2<>"",AU2>0,AU2<0.7)),"障礙", IF( OR(AND(BD2<>"",BD2>0,BD2<0.7), AND(BE2<>"",BE2>0,BE2<0.7), AND(BF2<>"",BF2>0,BF2<0.7), AND(BJ2<>"",BJ2>0,BJ2<0.7), AND(BK2<>"",BK2>0,BK2<0.7), AND(BL2<>"",BL2>0,BL2<0.7), AND(BM2<>"",BM2>0,BM2<0.7), AND(BN2<>"",BN2>0,BN2<0.7), AND(BO2<>"",BO2>0,BO2<0.7)),"障礙", IF( OR(AND(BX2<>"",BX2>0,BX2<0.7), AND(BY2<>"",BY2>0,BY2<0.7), AND(BZ2<>"",BZ2>0,BZ2<0.7), AND(CD2<>"",CD2>0,CD2<0.7), AND(CE2<>"",CE2>0,CE2<0.7), AND(CF2<>"",CF2>0,CF2<0.7), AND(CG2<>"",CG2>0,CG2<0.7), AND(CH2<>"",CH2>0,CH2<0.7), AND(CI2<>"",CI2>0,CI2<0.7)),"障礙", IF(OR(CJ2="住抗",CJ2="暫時移除設備"),"抗爭", IF(CJ2<>"","障礙", IF(DM2>2,"干擾", IF(Q2=6,"CC6", IF( OR(AND(DD2<>"",DD2>0.8),AND(DE2<>"",DE2>0.8),AND(DF2<>"",DF2>0.8)),"PRB>80", IF(AND(DC2>-106,DC2<-30),"RSRP優於-106", IF(DC2<=-106,"RSRP劣於-106", ""))))))))))))))))'


b = '=IF(R2="作業","障礙",\
IF(R2="障礙","障礙",\
IF(R2="抗爭","抗爭",\
IF(R2="40055重大障礙","40055重大障礙",\
IF(R2="非TWM問題的障礙","非TWM問題的障礙",\
IF(U2=35806,"非TWM問題的障礙",\
IF( OR(AND(AJ2<>"",AJ2>0,AJ2<0.7),\
       AND(AK2<>"",AK2>0,AK2<0.7),\
       AND(AL2<>"",AL2>0,AL2<0.7),\
       AND(AP2<>"",AP2>0,AP2<0.7),\
       AND(AQ2<>"",AQ2>0,AQ2<0.7),\
       AND(AR2<>"",AR2>0,AR2<0.7),\
       AND(AS2<>"",AS2>0,AS2<0.7),\
       AND(AT2<>"",AT2>0,AT2<0.7),\
       AND(AU2<>"",AU2>0,AU2<0.7)),"障礙",\
IF( OR(AND(BD2<>"",BD2>0,BD2<0.7),\
       AND(BE2<>"",BE2>0,BE2<0.7),\
       AND(BF2<>"",BF2>0,BF2<0.7),\
       AND(BJ2<>"",BJ2>0,BJ2<0.7),\
       AND(BK2<>"",BK2>0,BK2<0.7),\
       AND(BL2<>"",BL2>0,BL2<0.7),\
       AND(BM2<>"",BM2>0,BM2<0.7),\
       AND(BN2<>"",BN2>0,BN2<0.7),\
       AND(BO2<>"",BO2>0,BO2<0.7)),"障礙",\
IF( OR(AND(BX2<>"",BX2>0,BX2<0.7),\
       AND(BY2<>"",BY2>0,BY2<0.7),\
       AND(BZ2<>"",BZ2>0,BZ2<0.7),\
       AND(CD2<>"",CD2>0,CD2<0.7),\
       AND(CE2<>"",CE2>0,CE2<0.7),\
       AND(CF2<>"",CF2>0,CF2<0.7),\
       AND(CG2<>"",CG2>0,CG2<0.7),\
       AND(CH2<>"",CH2>0,CH2<0.7),\
       AND(CI2<>"",CI2>0,CI2<0.7)),"障礙",\
IF(OR(CJ2="住抗",CJ2="暫時移除設備"),"抗爭",\
IF(CJ2<>"","障礙",\
IF(DM2>2,"干擾",\
IF(Q2=6,"CC6",\
IF( OR(AND(DD2<>"",DD2>0.8),AND(DE2<>"",DE2>0.8),AND(DF2<>"",DF2>0.8)),"PRB>80",\
IF(AND(DC2>-106,DC2<-30),"RSRP優於-106",\
IF(DC2<=-106,"RSRP劣於-106",\
""))))))))))))))))'
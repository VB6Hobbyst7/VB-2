Attribute VB_Name = "Module1"
Function ConV(KK As String) As String
Dim Dcod(25) As String
Dim LL As Integer
Dim PP As Integer
Dim Check As Integer
Dim Temp As String

Dcod(0) = "�a�b�e�h�i�j�k�q�s�t�u�v�w�x�y�{�|�}���������������������������������ňɈ����������������������A�E�I�Q�S�U�V�W�a�b�c�e�h�i�q�s�u�v�w�{���������������������������������������ŉɉ�"
Dcod(1) = "�щӉՉ׉��������A�B�E�I�Q�S�U�W�a�e�i�s�u�����������������������������������������Պ��������A�E�I�a�b�e�h�i�j�q�s�u�w���������������������������a�b�c�e�i�k�q�s�u�v�w�{����������"
Dcod(2) = "����������������������������A�B�E�Q�U�W�a�e�i�u�v�{���������������������������ɍ֍׍����A�E�I�Q�S�W�a�������������������������������ŎɎюӎ֎�����A�a�b�e�g�i�k�p�q�s�u�w"
Dcod(3) = "�{�����������������a�b�c�e�h�i�j�k�q�s�u�v�w�x�y�{�}������������������������������������������������A�B�E�I�Q�S�U�V�W�a�b�e�i�q�s�v�w�z�����������������������������őɑ֒A�E�I�Q�S"
Dcod(4) = "�U�a�b�e�i�s�u�w������������������������������A�B�I�Q�S�W�a�b�e�i�j�k�q�s�u�w�x�|���������������������������a�b�c�e�h�i�j�k�l�p�q�s�u�v�w�x�y�}�������������������������������"
Dcod(5) = "���������A�B�E�I�Q�S�U�V�W�a�e�i�v�w���������������������������������ŕɕ���A�E�I�Q�S�U�a�����������������������ז��������A�E�I�Q�W�a�b�e�h�i�k�q�s�u�w�������������������������a�b�e�i"
Dcod(6) = "�q�s�u�v�w�}�����������������������������������A�B�E�I�Q�S�U�V�W�a�v�������������ə�A�E������������������������a�b�e�h�i�q�s�u�������������������������a�b�e�i�q�s�u�v�w�x�|�}"
Dcod(7) = "�������������������������������������������A�B�E�I�Q�S�U�W�a�b�e�i�q�s�u�v�w���������������������������ŝם��A�E�I�Q�S�U�W�a�e�i�s�u�w��������������������������������A�B�E�I�Q�S"
Dcod(8) = "�U�W�a�b�e�i�q�s�u�w�x�{�|�����������������a�b�e�g�h�i�j�k�q�s�u�w�x�{�}��������������������������������������������A�B�E�I�Q�S�U�V�W�a�b�e�i�u�v�w�y�������������������������š֡עA"
Dcod(9) = "�E�I�S�U�W�a�e�i�s�u��������������������������������������������A�E�I�Q�U�a�e�i�q�u��X�����������������������a�b�c�d�e�h�i�j�k�l�q�s�u�w�{��������������������������������������"
Dcod(10) = "�����A�B�E�H�I�Q�S�U�V�W�a�b�e�i�s�u�v�w�{�������������������������ť֥���A�B�E�I�Q�S�a�e�����������������������������������������A�E�I�Q�U�W�a�b�e�i�q�s�u�����������������������a�b�e"
Dcod(11) = "�i�k�q�s�u�v�w�}�������������������������������������A�W�a�b�q�s�u�v�w���������������A�a�w���������������A�W�a�e�i�q�s�����������������a�b�d�e�h�i�j�k�q�s�u�v�w�{��������������������"
Dcod(12) = "�����������������ŬɬѬ׬�������������������A�B�E�I�Q�S�U�V�W�a�b�e�i�q�s�u�v�w���������������������������������­ŭɭ׭��������A�E�I�Q�S�U�a�b�e�i�q�s�u�w������������������"
Dcod(13) = "�������������®ŮɮѮ׮����������A�B�I�Q�U�W�a�b�e�i�j�q�s�u�w�����������������������a�b�d�e�i�q�s�v�w�}������������������������������A�E�I�W���������������������±űֱ���A�E"
Dcod(14) = "�I�Q�S�a�����������������������W�a�b�e�i�k�p�q�s�������������������������a�b�e�f�g�i�j�k�p�q�s�u�v�w�{�|�������������������������������������������ŴɴӴ�����������������������"
Dcod(15) = "�A�B�E�I�Q�S�U�W�a�b�c�e�i�k�l�q�s�t�u�v�w�{�|�}�������������������������������������������µŵɵѵӵյֵ׵��������A�B�E�I�Q�S�U�W�a�b�e�i�q�s�u�w��������������������������������������"
Dcod(16) = "�¶ŶɶѶӶ׶����������A�B�E�I�Q�S�U�W�Y�a�b�e�i�o�q�s�u�w�x�y�z�{�|�}���������������������������������������a�b�e�g�h�i�k�q�s�u�v�w�x�����������������������������������Ÿɸ�����"
Dcod(17) = "���������A�B�E�I�Q�S�U�W�a�e�i�q�s�v�w�����������������������������¹ɹӹչ׹�����A�E�I�Q�S�U�W�a�b�e�w�����������������������������������A�E�I�Q�a�b�e�i�q�s�u�w������������������"
Dcod(18) = "���������a�b�e�g�i�l�q�s�u�v�w�������������������������������������A�W�a�v���������������������½ɽֽ���A�E�I�Q�S�w����������������������A�a�q�u�w�������������������a�b�e�g�i�q�s�u"
Dcod(19) = "�v�w�x�������������������������������������������������A�B�E�I�Q�S�U�W�a�e�v�������������������������������A�E�I�Q�S�U�W�a�q¡¶�������������������A�E�I�Q�W�a�b�e�i�q�s�u�w"
Dcod(20) = "áâåèéêñóõ÷�a�b�e�i�q�s�u�wāĂąĉđēĕĖėġĢķ���������������������A�B�E�I�Q�S�U�W�a�e�i�q�s�u�v�wŁšŢťũűųŵŷ�����������������A�I�aƁƂƅƉƑƓƕƗơƥƩƷ������������"
Dcod(21) = "���������A�E�I�Q�a�b�e�i�q�s�wǡǢǥǩǱǳǵǷ�a�b�e�i�j�q�s�u�v�wȁȂȅȉȑȓȕȖȗȡȷ���������������������A�B�E�I�Q�S�U�W�a�e�vɁɅɡɢɥɩɱɳɵɷɼ�������A�E�U�W�aʁʂʅʉʑʓʕʗʡʶ"
Dcod(22) = "�����������������A�E�I�Q�W�a�b�e�h�i�k�q�s�uˁ˅ˉˑ˓ˡˢ˥˩˱˳˵˷�a�b�c�e�i�k�q�s�u�v�w�{̡̢̖̗́̂̅̉̑̓̕�������������������A�B�E�I�Q�S�U�W�a�e�i�q�s�v�ẃ͉͓͕ͥͩ͢͡ͱͳ͵ͷ"
Dcod(23) = "�����A�E�a�e�i�s�u΁΂΅ΈΉ΋ΑΓΕΗΡη�����������A�E�I�Q�U�W�a�e�i�q�s�uϡϢϥϩϱϳϵϷ�a�b�e�i�n�q�s�u�wЁЂЅЉБГЕЖЗСз�������������������A�B�E�I�Q�S�U�W�a�b�e�i�q�s�u�v�wсх"
Dcod(24) = "щѓѡѢѥѩѮѱѳѵѷѻ�����������������������A�B�E�I�S�U�W�a�e�i�s�uҁ҂҅҉Ҏґҕҗҡҥҩұҷ���������������������������A�B�E�I�Q�U�W�a�b�e�g�h�i�j�q�s�u�w�{ӁӅӉӑӓӗӡӢӥөӱӳӵӷ"

Check = 1
Counter = 1
Temp = ""
Do While Not Counter > Len(KK)
   If Not (176 > (Asc(MidB(KK, Check, 1)))) And Not (210 < (Asc(MidB(KK, Check, 1)))) Then
      LL = Asc(MidB(KK, Check, 1))
      PP = Asc(MidB(KK, Check + 1, 1))
      LL = LL - 176
      PP = PP - 160
      Temp = Temp + Mid(Dcod(LL), PP, 1)
      Check = Check + 2
   Else
      Temp = Temp + Mid(KK, Counter, 1)
      Check = Check + 1
   End If
      Counter = Counter + 1
Loop

ConV = Temp
End Function

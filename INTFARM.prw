#INCLUDE "rwmake.ch"

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �IntFarm      � Autor � Flavio Martins  � Data �  11/09/15   ���
��� Programa faz leitura do arquivo txt gerado pelas Farmacias conveniadas �� 
��� e gera os lan�amentos												   ��
�������������������������������������������������������������������������͹��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/

User Function INTFARM()


//���������������������������������������������������������������������Ŀ
//� Declaracao de Variaveis                                             �
//�����������������������������������������������������������������������

//Grava os dados do registro posicionado
Local cPerg := "ZZINTFARM" 
_cAlias := Alias()
_cIndex := IndexOrd()
_cRecno := Recno()
aLog	:= {}
sMat	:= ""
nRegsOk := 0
nTotOk  := 0
nTotErr := 0

Private oLeTxt

//���������������������������������������������������������������������Ŀ
//� Montagem da tela de processamento.                                  �
//�����������������������������������������������������������������������

@ 200,1 TO 340,395 DIALOG oLeTxt TITLE OemToAnsi("Integra��o - Vale Farm�cia")
@ 02,10 TO 065,190
@ 14,018 Say " Este programa vai ler o conte�do do arquivo texto enviado pela"
@ 22,018 Say " farm�cia e integrar a verba correspondente nos lan�amentos mensais."

@ 45,98 BMPBUTTON TYPE 17 ACTION Pergunte(cPerg,.T. )
@ 45,128 BMPBUTTON TYPE 01 ACTION OkLeTxt()
@ 45,158 BMPBUTTON TYPE 02 ACTION Close(oLeTxt)
Activate Dialog oLeTxt Centered

Return

Static Function OkLeTxt

//���������������������������������������������������������������������Ŀ
//� Abertura do arquivo texto                                           �
//�����������������������������������������������������������������������
vProcesso 		:= ALLTRIM(mv_par01)
vRoteiro		:= ALLTRIM(mv_par02)
vPeriodo		:= ALLTRIM(mv_par03)
vCodFarmacia 	:= ALLTRIM(mv_par04)  
vVerba			:= ""
vNomFarmacia	:= "" 

DO CASE 
	Case vCodFarmacia = "01"
		vNomFarmacia := "CLAMED" 
		vVerba := "522"
	Case vCodFarmacia = "02"
		vNomFarmacia := "MODERNA"
		vVerba := "511"
	Case vCodFarmacia = "03"
		vNomFarmacia := "SESI"
		vVerba := "693"
	Case vCodFarmacia = "04"
		vNomFarmacia := "DROGARAIA" 
		vVerba := "536" 
	Case vCodFarmacia = "05"
		vNomFarmacia := "PHARMATIVA" 
		vVerba := "677"
ENDCASE

IF (vVerba <> "")		 

   IF(XFILIAL("SRA")="14")
		Private cArqTxt := "D:\"+vNomFarmacia+"\ORMEC.TXT"  
   ELSE
		Private cArqTxt := "C:\"+vNomFarmacia+"\ORMEC.TXT"    
   ENDIF                  
	
   Private nHdl    := fOpen(cArqTxt,68)
   Private cEOL    := "CHR(13)+CHR(10)"
   
   If Empty(cEOL)
	    cEOL := CHR(13)+CHR(10)
   Else
        cEOL := Trim(cEOL)
	    cEOL := &cEOL
   Endif
	
   If nHdl == -1
      MsgAlert("O arquivo n�o se encontra no local "+cArqTxt+CHR(13)+"e por isso n�o pode ser aberto."+CHR(13)+"Crie a pasta "+vNomFarmacia+" no [C:\] e salve o arquivo enviado pela "+vNomFarmacia+" "+CHR(13)+"como ORMEC.TXT dentro desta pasta."+CHR(13)+"Verifique os atributos deste arquivo, remova o atributo 'Somente leitura'.","Aten��o!")
      Return
   Endif
	
	//���������������������������������������������������������������������Ŀ
	//� Inicializa a regua de processamento                                 �
	//�����������������������������������������������������������������������
	
	Processa({|| RunCont() },"Processando...")
	Return
ENDIF
	

Static Function RunCont

Local nTamFile, nTamLin, cBuffer, nBtLidos, nContLin, nContPag

//�����������������������������������������������������������������ͻ
//� Lay-Out do arquivo Texto gerado:                                �
//�����������������������������������������������������������������͹
//�Campo           � Inicio � Tamanho                               �
//�����������������������������������������������������������������Ķ
//� ??_FILIAL     � 01     � 02                                    �
//�����������������������������������������������������������������ͼ

nTamFile := fSeek(nHdl,0,2)
fSeek(nHdl,0,0)
nTamLin  := 20
cBuffer  := Space(nTamLin) // Variavel para criacao da linha do registro para leitura
nBtLidos := fRead(nHdl,@cBuffer,nTamLin) // Leitura da primeira linha do arquivo texto

ProcRegua(nTamFile) // Numero de registros a processar

While nBtLidos >= nTamLin

    //���������������������������������������������������������������������Ŀ
    //� Incrementa a regua                                                  �
    //�����������������������������������������������������������������������

    IncProc()

	//���������������������������������������������������������������������Ŀ
    //� Contador de linhas                                                  �
    //�����������������������������������������������������������������������
    
    //���������������������������������������������������������������������Ŀ
    //� Grava os campos obtendo os valores da linha lida do arquivo texto.  �
    //�����������������������������������������������������������������������

	GravaRC(cBuffer)

    //���������������������������������������������������������������������Ŀ
    //� Leitura da proxima linha do arquivo texto.                          �
    //�����������������������������������������������������������������������
	
    nBtLidos := fRead(nHdl,@cBuffer,nTamLin) // Leitura da proxima linha do arquivo texto

    dbSkip()
EndDo

//���������������������������������������������������������������������Ŀ
//� O arquivo texto deve ser fechado, bem como o dialogo criado na fun- �
//� cao anterior.                                                       �
//�����������������������������������������������������������������������

//Volta ao registro posicionado no chamado do ponto
DbSelectArea(_cAlias)
DbSetOrder(_cIndex)
DbGoTo(_cRecno)

fClose(nHdl)
Close(oLeTxt)
GeraLog(aLog)
Return()

//*******************************
Static Function GravaRC(sLinha)
//*******************************

Local sFilMat := Space(8)
                                 
If Len(Alltrim(Left(sLinha,8))) = 7 
	sFilMat := Alltrim(Left(sLinha,8)) + " "	
Else 
	sFilMat := Alltrim(Left(sLinha,8))
Endif

DbSelectArea("SRA")
DbSetOrder(1) //FILIAL+MATRICULA
DbGoTop()

If DbSeek(sFilMat) = .T.

	If DtoC(SRA->RA_DEMISSA) = "  /  /    " 

       If SRA->RA_CATFUNC <> "P"
			DbSelectArea("RGB")
			DbSetOrder(1) //FILIAL+MATRICULA+VERBA
			DbGoTop()

			If DbSeek(sFilMat+vVerba) = .T.				
				Alert(sFilMat+vVerba)
				If RGB->RGB_FILIAL+RGB->RGB_MAT = sMat								
					RecLock("RGB",.F.) 
					RGB->RGB_VALOR 	:= RGB->RGB_VALOR + val(substring(sLinha,9,8)+"."+substr(sLinha,17,2))
					MSUnLock("RGB")
					nRegsOk++
					nTotOk := nTotOk + val(substring(sLinha,9,8)+"."+substr(sLinha,17,2))				
				Else                                   
					AaDd(aLog,"Verba j� inclu�da para o funcion�rio(a) ["+Alltrim(sFilmat)+"]. Verifique os lan�amentos mensais deste(a) funcion�rio(a)."+CHR(13))
					nTotErr := nTotErr + val(substring(sLinha,9,8)+"."+substr(sLinha,17,2))					
				Endif
			Else
				RecLock("RGB",.T.)
				RGB->RGB_FILIAL	:= Left(sFilMat,2)
				RGB->RGB_PROCES := vProcesso
				RGB->RGB_ROTEIR := vRoteiro
				RGB->RGB_PERIOD := vPeriodo
				RGB->RGB_SEMANA := '01'
				RGB->RGB_MAT	:= Alltrim(Substr(sFilMat,3,6))
				RGB->RGB_PD 	:= vVerba
				RGB->RGB_TIPO1 	:= "V"
				RGB->RGB_HORAS 	:= 0
				RGB->RGB_VALOR 	:= val(substring(sLinha,9,8)+"."+substr(sLinha,17,2))
				RGB->RGB_CC 		:= SRA->RA_CC
				RGB->RGB_PARCELA := 0
				RGB->RGB_TIPO2 	:= "G"
				MSUnLock("RGB")
				nRegsOk++
				nTotOk := nTotOk + RGB->RGB_VALOR				
				sMat := sFilMat
			Endif		
		Else 
			AaDd(aLog,"Funcion�rio(a) ["+Alltrim(sFilmat)+"] encontra-se como PR�-LABORE, verifique o campo Cat. Func. na pasta Funcionais no cadastro deste funcion�rio(a)."+CHR(13))		
	   		nTotErr := nTotErr + val(substring(sLinha,9,8)+"."+substr(sLinha,17,2))								
		Endif			
	Else 
		AaDd(aLog,"Funcion�rio(a) ["+Alltrim(sFilmat)+"] n�o encontra-se admitido(a) e/ou foi transferido(a). Se este(a) funcion�rio(a) foi transferido(a), atualize a filial e a matr�cula na "+vNomFarmacia+"."+CHR(13))		
		nTotErr := nTotErr + val(substring(sLinha,9,8)+"."+substr(sLinha,17,2))					
	Endif
Else
	AaDd(aLog,"Funcion�rio(a) ["+Alltrim(sFilmat)+"] n�o encontrado(a). Atualize a filial e a matr�cula deste(a) funcion�rio(a) na "+vNomFarmacia+"."+CHR(13))
	nTotErr := nTotErr + val(substring(sLinha,9,8)+"."+substr(sLinha,17,2))					
Endif

Return

//*******************************
Static Function GeraLog(aLog)    
//*******************************

//���������������������������������������������������������������������Ŀ
//� Cria o arquivo texto                                                �
//�����������������������������������������������������������������������
Local sMsg, sMsgL1, sMsgL2, sMsgL3
Private cArqTxt := "C:\"+vNomFarmacia+"\LOG "+STRZERO(day(date()),2)+STRZERO(month(date()),2)+STRZERO(year(date()),4)+"-"+left(time(),2)+substr(time(),4,2)+right(time(),2)+".txt"
Private nHdl    := fCreate(cArqTxt)

Private cEOL    := "CHR(13)+CHR(10)"
If Empty(cEOL)
    cEOL := CHR(13)+CHR(10)
Else
    cEOL := Trim(cEOL)
    cEOL := &cEOL
Endif
                                                      	
If nHdl == -1
    MsgAlert("O arquivo de nome "+cArqTxt+" n�o pode ser executado! Verifique os par�metros.","Aten��o!")
    Return
Else                 
	sMsgL1 := CHR(13)
	AaDd(aLog,sMsgL1) 
	sMsgL1 := Transform(nRegsOk,"@E 9999")+" registro(s) processado(s) com sucesso, no valor de R$ "+Transform(nTotOk,"@E 99,999.99")+CHR(13)
	AaDd(aLog,sMsgL1)
	sMsgL2 := Transform(len(aLog)-2,"@E 9999")+" registro(s) processado(s) com erro(s), no valor de R$ "+Transform(nTotErr,"@E 99,999.99")+CHR(13)
	AaDd(aLog,sMsgL2)
	sMsgL3 := Transform((len(aLog)-3+nRegsOk),"@E 9999")+" registro(s) processado(s)  no total com o valor de R$ "+Transform((nTotOk+nTotErr),"@E 99,999.99")
	AaDd(aLog,sMsgL3)		
	sMsg := sMsgL1 + sMsgL2 + sMsgL3
	sMsg := sMsg + CHR(13)+CHR(13)+ "Verifique o arquivo "+Alltrim(cArqTxt)+CHR(13)
	sMsg := sMsg + "para maiores informa��es."
	MsgAlert(sMsg,"Resumo da importa��o")
Endif

//���������������������������������������������������������������������Ŀ
//� Inicializa a regua de processamento                                 �
//�����������������������������������������������������������������������

Processa({|| RunContLog() },"Processando...")
Return

Static Function RunContLog

Local nTamLin, cLin, cCpo


//���������������������������������������������������������������������Ŀ
//� Posicionamento do primeiro registro e loop principal. Pode-se criar �
//� a logica da seguinte maneira: Posiciona-se na filial corrente e pro �
//� cessa enquanto a filial do registro for a filial corrente. Por exem �
//� plo, substitua o dbGoTop() e o While !EOF() abaixo pela sintaxe:    �
//�                                                                     �
//� dbSeek(xFilial())                                                   �
//� While !EOF() .And. xFilial() == A1_FILIAL                           �
//�����������������������������������������������������������������������

ProcRegua(RecCount()) // Numero de registros a processar

For i:=1 to len(aLog)

    //���������������������������������������������������������������������Ŀ
    //� Incrementa a regua                                                  �
    //�����������������������������������������������������������������������

    IncProc()

    //�����������������������������������������������������������������ͻ
    //� Lay-Out do arquivo Texto gerado:                                �
    //�����������������������������������������������������������������͹
    //�Campo           � Inicio � Tamanho                               �
    //�����������������������������������������������������������������Ķ
    //� ??_FILIAL     � 01     � 02                                    �
    //�����������������������������������������������������������������ͼ

    nTamLin := 2
    cLin    := Space(nTamLin)+cEOL // Variavel para criacao da linha do registros para gravacao

    //���������������������������������������������������������������������Ŀ
    //� Substitui nas respectivas posicioes na variavel cLin pelo conteudo  �
    //� dos campos segundo o Lay-Out. Utiliza a funcao STUFF insere uma     �
    //� string dentro de outra string.                                      �
    //�����������������������������������������������������������������������

    cCpo := alog[i]
    cLin := Stuff(cLin,01,02,cCpo)

    //���������������������������������������������������������������������Ŀ
    //� Gravacao no arquivo texto. Testa por erros durante a gravacao da    �
    //� linha montada.                                                      �
    //�����������������������������������������������������������������������

    If fWrite(nHdl,cLin,Len(cLin)) != Len(cLin)

        If !MsgAlert("Ocorreu um erro na grava��o do arquivo. Continua?","Aten��o!")
            Exit
        Endif
        
    Endif

Next i

//���������������������������������������������������������������������Ŀ
//� O arquivo texto deve ser fechado, bem como o dialogo criado na fun- �
//� cao anterior.                                                       �
//�����������������������������������������������������������������������

fClose(nHdl)
//Close(oGeraTxt)

Return
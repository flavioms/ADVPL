#include 'protheus.ch'
#include 'parmtype.ch'

/*/{Protheus.doc} U_NDTDEBITO()
//TODO Descrição auto-gerada.
@author administrator
@since 15/10/2018
@version 1.0
@return ${return}, ${return_description}

@type function
/*/

User Function NDTDEBITO()
	Local nOpcA := 0
	Local cPerg := "ZZNFDEBTO"
	Local _cSQL
	Local _resultado := ""  
	PRIVATE cTitulo := "Nota de Débito"
	PRIVATE oPrn, oArial07, oArial09B, oArial14B
	PRIVATE aNota := {}
	Private aSay := {}
	Private aButton := {}   

	Pergunte(cPerg,.F.)
	
	AADD(aSay, 'Impressão da nota fiscal de débito')
	aAdd( aButton, { 5,.T.,{|| Pergunte(cPerg,.T. )}})
	aAdd( aButton, { 1,.T.,{|| nOpca := 1,FechaBatch()}})
	aAdd( aButton, { 2,.T.,{|| FechaBatch() }} )
	FormBatch( cTitulo, aSay, aButton )
	
	If nOpcA == 1
		Processa( {|| CarregaNota() }, "Processando..." )
		Processa( {|| ExecRel() }, "Processando..." )
	Endif

Return

Static function ExecRel()
	DEFINE FONT oArial10 NAME "Arial" SIZE 0,10 OF oPrn
	DEFINE FONT oArial10B NAME "Arial" SIZE 0,10 OF oPrn BOLD
	DEFINE FONT oArial12B NAME "Arial" SIZE 0,12 OF oPrn BOLD
	DEFINE FONT oArial20B NAME "Arial" SIZE 0,16 OF oPrn BOLD
	
	oPrn := TMSPrinter():New(cTitulo)
	oPrn:SetPortrait()
	oPrn:SetPaperSize(9)
	oPrn:StartPage()
	
	Cabec()
	DadosNota()
	DadosCliente()
	Descricao()
	
	oPrn:EndPage()
	oPrn:End()
	oPrn:Preview()
	oPrn:End()
	
return

Static function Cabec()
	
	oPrn:Box ( 080, 080, 480, 2350)
	oPrn:SayBitmap(090, 1200, "\system\LogoOrmec.bmp")
	oPrn:Say(350, 0900, "ORMEC ENGENHARIA LTDA", oArial20B)
return

Static function DadosNota()
	oPrn:Box (480, 080, 1000, 2350)
	oPrn:Say (510, 1000, "DADOS DA NOTA DE DÉBITO", oArial12B)
	oPrn:Line(600, 080, 600, 2350)
	oPrn:Box (600, 080, 1000, 1250)
	oPrn:Say(610, 090, "Filial: " + aNota[1], oArial10)
	oPrn:Say(670, 090, aNota[2] + " , " + aNota[3], oArial10)
	oPrn:Say(730, 090, aNota[4] + " - " + aNota[5] + " - CEP: " + aNota[6], oArial10)
	oPrn:Say(790, 090,"CNPJ: " + aNota[7], oArial10)
	oPrn:Say(850, 090,"Insc: " + aNota[8], oArial10)
	oPrn:Say(910, 090,"Tel: " + aNota[9], oArial10)
	
	oPrn:Say(610, 1310, "Fatura: "+ aNota[10], oArial10)
	oPrn:Say(670, 1310, "Data de Emissão: "+ aNota[11], oArial10)
	oPrn:Say(730, 1310, "Vencimento: "+ aNota[12], oArial10)
return

Static function DadosCliente()
	oPrn:Box (1000, 080, 1350, 2350)
	oPrn:Say (1030, 1040, "DADOS DO CLIENTE", oArial12B)
	oPrn:Line(1100, 080, 1100, 2350)
	oPrn:Say(1110, 090, "Cliente: " + aNota[13], oArial10)
	oPrn:Say(1170, 090, "Endereço: " + aNota[14], oArial10)
	oPrn:Say(1230, 090, "Bairro: "+ aNota[15], oArial10)
	oPrn:Say(1290, 090,"CNPJ: " + aNota[16], oArial10)
	
	oPrn:Say(1110, 1310,"Estado: " + aNota[17], oArial10)
	oPrn:Say(1170, 1310,"Cidade: " + aNota[18], oArial10)
	oPrn:Say(1230, 1310, "Insc. Estadual: "+ aNota[19], oArial10)
	oPrn:Say(1290, 1310, "CEP: "+ aNota[20], oArial10)
return

Static function Descricao()
	oPrn:Box (1400, 080, 1550, 2350)
	oPrn:Say (1450, 1150, "DESCRIÇÃO", oArial12B)
	
	oPrn:Box (1550, 080, 1600, 250)
	oPrn:Say (1555, 090, "ITEM", oArial10B)
	oPrn:Box (1550, 250, 1600, 2000)
	oPrn:Say (1555, 260, "DESCRIÇÃO", oArial10B)
	oPrn:Box (1550, 2000, 1600, 2350)
	oPrn:Say (1555, 2010, "PREÇO TOTAL", oArial10B)
	nLin := 1550
	For i := 1 To len(aNota[21])
		nLin += 50
		oPrn:Box (nLin, 080, nLin + 50, 250)
		oPrn:Box (nLin, 250, nLin + 50, 2000)
		oPrn:Box (nLin, 2000, nLin + 50, 2350)
		
		oPrn:Say (nLin + 5, 090, aNota[21][i], oArial10B)
		oPrn:Say (nLin + 5, 260, aNota[22][i], oArial10B)
		oPrn:Say (nLin + 5, 2010,"R$" + TRANSFORM(aNota[23][i], "@E 999,999,999.99"), oArial10B)
	next
	
	nLin += 110
	oPrn:Box (nLin, 080, nLin+400, 1900)
	oPrn:Say (nLin+10, 090, "Observações: "+ Alltrim(mv_par03), oArial10B)
	
	oPrn:Box (nLin, 1900, nLin+133.33, 2350)
	vTotalItens := 0
	For i := 1 To len(aNota[23])
		vTotalItens += aNota[23][i]
	Next
	oPrn:Say (nLin+10, 1910, "VALOR TOTAL: ", oArial10)
	oPrn:Say (nLin+50, 1910, "R$ " + Alltrim( TRANSFORM(vTotalItens, "@E 999,999,999.99") ), oArial10)
	 
	oPrn:Box (nLin+133.33, 1900, nLin+266.66, 2350)
	oPrn:Say (nLin+143.33, 1910, "DESCONTO:", oArial10)
	oPrn:Say (nLin+193.33, 1910, "R$" + ALLTRIM(aNota[24]), oArial10)
	
	oPrn:Box (nLin+266.66, 1900, nLin+400, 2350)
	oPrn:Say (nLin+276.66, 1910, "TOTAL DA FATURA:", oArial10)
	oPrn:Say (nLin+326.66, 1910, "R$" + ALLTRIM(aNota[25]), oArial10)
	
	nLin += 400
	oPrn:Box (nLin, 080, nLin+100, 500)
	oPrn:Say (nLin+20, 240, "FATURA", oArial10B)
	oPrn:Box (nLin, 500, nLin+100, 2350)
	oPrn:Say (nLin+20, 520, "RECEBI(EMOS) DE ORMEC ENGENHARIA LTDA, OS VALORES CONSTANTES DESTA FATURA", oArial10B)
	
	nLin += 100
	oPrn:Box (nLin, 080, nLin+200, 500)
	oPrn:Say (nLin+20, 220, aNota[10], oArial12B)
	oPrn:Box (nLin, 500, nLin+200, 2350)
	oPrn:Line(nLin+80, 520, nLin+80, 860)
	oPrn:Say (nLin+90, 630, "DATA", oArial12B)
	oPrn:Line(nLin+80, 1000, nLin+80, 1500)
	oPrn:Say (nLin+90, 1140, "ASSINATURA", oArial12B)
	
	nLin += 200
	oPrn:Box (nLin, 080, nLin+350, 2350)
	nLin += 10
	oPrn:Say(nLin, 090, "Filial: " + aNota[1], oArial10)
	nLin += 60
	oPrn:Say(nLin, 090, aNota[2] + " , " + aNota[3] + " " + aNota[4] + " - " + aNota[5] + " - CEP: " + aNota[6], oArial10)
	nLin += 60
	oPrn:Say(nLin, 090,"CNPJ: " + aNota[7], oArial10)
	nLin += 60
	oPrn:Say(nLin, 090,"Tel: " + aNota[9], oArial10)
	
	
return

Static function CarregaNota()
	Local	vD2Item		:= {}
	Local	vD2Desc 	:= {}
	Local	vD2Prod		:= {}
	Local	vD2Total 	:= {}
	Local aAreaAnt := GETAREA()
    Local aAreaM0 := SM0->(GetArea())
    Local cAliaTmp	:= GetNextAlias()
    Local cQuery 	:= ''
	
	vDoc 			:= ALLTRIM(mv_par01)
	vSerie 			:= ALLTRIM(mv_par02)
////////////////// Busca informações da Nota //////////////////////////
	dbSelectArea("SF2")
		dbSetOrder(10)
		dbSeek( XFilial("SF2")+vDoc+vSerie )
		
		If !Found()
			MSGBOX("Erro de consulta no SF2.","Alert","ALERT" )
			RETURN
		End
		
		IF LEN(ALLTRIM((SF2->F2_CLIENTE))) = 6 
			vCliente		:= ALLTRIM(SF2->F2_CLIENTE)
		ELSE
			vCliente        := ALLTRIM(SF2->F2_CLIENTE)+REPLICATE(' ', 6 - LEN(ALLTRIM((SF2->F2_CLIENTE))))
		ENDIF
		vLoja			:= ALLTRIM(SF2->F2_LOJA)
		vF2Desconto		:= SF2->F2_DESCONT
		vF2TotalFatura	:= SF2->F2_VALFAT  
		vF2Emissao		:= SF2->F2_EMISSAO
	dbCloseArea()
	
////////////////// Busca informações dos Itens da Nota //////////////////////////
	dbSelectArea("SD2")
		dbSetOrder(18)
		dbSeek( XFilial("SD2")+vDoc+vSerie+vCliente+vLoja )
		
		If !Found()
			MSGBOX("Erro de consulta no SD2.","Alert","ALERT" )
			RETURN
		End
		
		While (SD2->D2_DOC = vDoc .AND. SD2->D2_SERIE = vSerie)
			aAdd(vD2Item, SD2->D2_ITEM)
			aAdd(vD2Prod, ALLTRIM(SD2->D2_COD))
			aAdd(vD2Total, SD2->D2_TOTAL)
			DBSKIP()
		EndDo
	dbCloseArea()
	
////////////////// Busca a descrição do Produto //////////////////////////
	dbSelectArea("SB1")
		dbSetOrder(1)
		for nK := 1 to Len(vD2Prod)
			dbSeek( XFilial("SB1")+vD2Prod[nK] ) 
			IF FOUND()
				aAdd(vD2Desc, ALLTRIM(SB1->B1_DESC))
			ENDIF
		next
	dbCloseArea()
	
////////////////// Busca informações do Título //////////////////////////
	dbSelectArea("SE1")
		dbSetOrder(30)
		dbSeek("  "+vCliente+vLoja+vDoc+vF2Emissao)
		
		If !Found()
			MSGBOX("O PROGRAMA SERA FINALIZADO POIS NAO EXISTE TITULO DE CONTAS A RECEBER PARA ESTA NF")
			RETURN
		End
		
		vE1Emissao 		:= Right (DTOS(SE1->E1_EMISSAO),2) + "/" + SUBSTR(DTOS(SE1->E1_EMISSAO),5,2) + "/" + SUBSTR(DTOS(SE1->E1_EMISSAO),3,2)
		vE1Vencimento 	:= Right (DTOS(SE1->E1_VENCTO),2) + "/" + SUBSTR(DTOS(SE1->E1_VENCTO),5,2) + "/" + SUBSTR(DTOS(SE1->E1_VENCTO),3,2)
		vE1Valor		:= SE1->E1_VALOR
	dbCloseArea()

//////////////////////// Busca informações da Filial ////////////////////////

	BeginSQL Alias cAliaTmp
	 %noParser%
	 SELECT SM0010.M0_CODFIL, SM0010.M0_NOMECOM, SM0010.M0_ENDENT, SM0010.M0_BAIRENT, SM0010.M0_CGC, 
	 SM0010.M0_CEPENT, SM0010.M0_CIDENT, SM0010.M0_TEL, SM0010.M0_ESTENT, SM0010.M0_INSC
	 FROM SM0010
	 WHERE  SM0010.M0_CODFIL = %xFilial:SF2% AND SM0010.%notDel%
	EndSQL 
	
	IF(cAliaTmp)->(!Eof()) .AND. ALLTRIM((cAliaTmp)->M0_CODFIL) == XFilial("SF2")
		vM0Nomecom := ALLTRIM((cAliaTmp)->M0_NOMECOM) 
		vM0Endcob := ALLTRIM((cAliaTmp)->M0_ENDENT)
		vM0Baircob := ALLTRIM((cAliaTmp)->M0_BAIRENT)
		vM0CGC := ALLTRIM((cAliaTmp)->M0_CGC)
		vM0Cepcob := ALLTRIM((cAliaTmp)->M0_CEPENT)
		vM0Cidcob := ALLTRIM((cAliaTmp)->M0_CIDENT)
		vM0Telfilial := ALLTRIM((cAliaTmp)->M0_TEL)
		vM0Estcob := ALLTRIM((cAliaTmp)->M0_ESTENT)      
		vM0Insc := ALLTRIM((cAliaTmp)->M0_INSC)
	ENDIF
	(cAliaTmp)->(dbCloseArea())
	
//////////////////////// Busca informações do Cliente ////////////////////////
	dbSelectArea("SA1")
		dbSetOrder(1)
		dbSeek("  "+vCliente+vLoja)
		
		If !Found()
			MSGBOX("Erro de consulta no SA1.","Alert","ALERT" )
			RETURN
		Else
			vA1Nome 	:= ALLTRIM(SA1->A1_NOME)
			vA1Endcob 	:= ALLTRIM(SA1->A1_END)
			vA1Bairro 	:= ALLTRIM(SA1->A1_BAIRRO)
			vA1CGC 		:= ALLTRIM(SA1->A1_CGC)
			vA1Cep 		:= ALLTRIM(SA1->A1_CEP)
			vA1Inscr 	:= ALLTRIM(SA1->A1_INSCR)
			vA1Est 		:= ALLTRIM(SA1->A1_EST)
			vA1Tel 		:= ALLTRIM(SA1->A1_TEL)
			vA1Mun 		:= ALLTRIM(SA1->A1_MUN)
		End
	dbCloseArea()
	
/////////// CARREGANDO O ARRAY //////////////////////////


aAdd(aNota, vM0Nomecom)
aAdd(aNota, vM0Endcob)
aAdd(aNota, vM0Baircob)
aAdd(aNota, vM0Cidcob)
aAdd(aNota, vM0Estcob)
aAdd(aNota, vM0Cepcob)
aAdd(aNota, vM0CGC)
aAdd(aNota, vM0Insc)
aAdd(aNota, vM0Telfilial)

aAdd(aNota, vDoc)
aAdd(aNota, vE1Emissao)
aAdd(aNota, vE1Vencimento)

aAdd(aNota, vA1Nome)
aAdd(aNota, vA1Endcob)
aAdd(aNota, vA1Bairro)
aAdd(aNota, vA1CGC)
aAdd(aNota, vA1Est)
aAdd(aNota, vA1Mun)
aAdd(aNota, vA1Inscr)
aAdd(aNota, vA1Cep)

aAdd(aNota, vD2Item)
aAdd(aNota, vD2Desc)
aAdd(aNota, vD2Total)

aAdd(aNota, TRANSFORM(vF2Desconto, "@E 999,999,999.99"))
aAdd(aNota, TRANSFORM(vF2TotalFatura, "@E 999,999,999.99"))

return
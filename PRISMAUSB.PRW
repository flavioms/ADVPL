
#include "protheus.ch"

/*/
+-----------------------------------------------------------------------------
| Programa | PrismaUSB | Autor | FLAVIO MARTINS | 25/08/2015 | |
+-----------------------------------------------------------------------------
| Descrição | Este programa vai gerar um arquivo texto com os dados dos funcionarios
  para ser importado via USB no relógio de ponto PRISMA Super Fácil R02 |
+-----------------------------------------------------------------------------
| Uso | Ponto Eletrônico ADVPL |
+-----------------------------------------------------------------------------
/*/


User Function PrismaUSB()

	Local oPrismaUSB
	Private cPerg := "PRISMA"
	Private cAlias := "SRA"
	Public TMP1

	//CriaSx1(cPerg)
	Pergunte(cPerg,.F.)

	DEFINE MSDIALOG oPrismaUSB TITLE OemToAnsi("Geração de Arquivo Texto") ;
	FROM 000,000 TO 200,400 PIXEL
	//@ 005,005 TO 095,195 OF oPrismaUSB PIXEL 
	@ 005,005 TO 095,195 OF oPrismaUSB PIXEL 
	@ 010,010 Say " Este programa vai gerar um arquivo texto para ser importado via USB ";
	OF oPrismaUSB PIXEL 
	@ 018,010 Say " no relógio de ponto PRISMA Super Fácil R02 das filiais FVC e FCB. ";
	OF oPrismaUSB PIXEL
	@ 026,010 Say " O arquivo será gravado na pasta C:\Prisma\ verifique se esta pasta existe.";
	OF oPrismaUSB PIXEL


	DEFINE SBUTTON FROM 070, 045 TYPE 1 ;
	ACTION (OkPrismaUSB(),oPrismaUSB:End()) ENABLE OF oPrismaUSB

	DEFINE SBUTTON FROM 070, 085 TYPE 2 ;
	ACTION (oPrismaUSB:End()) ENABLE OF oPrismaUSB

	DEFINE SBUTTON FROM 070, 125 TYPE 5 ;
	ACTION (Pergunte(cPerg,.T.)) ENABLE OF oPrismaUSB

	ACTIVATE DIALOG oPrismaUSB CENTERED

Return Nil


Static Function OkPrismaUSB
	Local cLin

	dtEmissaoDe := DTOS(MV_PAR01)
	dtEmissaoAte := DTOS(MV_PAR02)
	matriculaDe := MV_PAR03
	matriculaAte := MV_PAR04
	cCustoDe := MV_PAR05
	cCustoAte := MV_PAR06

	cQuery := "SELECT RA_PIS, RA_NOME, RA_MAT "
	cQuery += "FROM SRA010 "
	cQuery += "WHERE RA_FILIAL='"+XFILIAL("SRA")+"' AND "
	cQuery += "RA_ADMISSA >= '"+dtEmissaoDe+"' AND RA_ADMISSA <= '"+dtEmissaoAte+"' AND "
	cQuery += "RA_MAT >= '"+matriculaDe+"' AND RA_MAT<= '"+matriculaAte+"' AND "
	cQuery += "RA_CC >= '"+cCustoDe+"' AND RA_CC<= '"+cCustoAte+"' AND "
	cQuery += "RA_SITFOLH <> 'D'"
	dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"TMP1",.T.,.T.)
	dbSelectArea("TMP1")
	dbGoTop()


	Private cArqTxt := "C:\PRISMA\rep_colaborador.txt"
	If MsFile (cArqTxt)
		MsErase(cArqTxt)
	EndIf

	Private nHdl := fCreate(cArqTxt)

	If nHdl == -1
		MsgAlert("O arquivo de nome "+cArqTxt+" não pode ser executado! Verifique os parâmetros.","Atenção!")
		Return
	Endif
	// Inicializa a régua de processamento
	Processa({|| RunCont() },"Processando...")

Return Nil

Static Function RunCont
	Local cLin
	ProcRegua(RecCount()) // Numero de registros a processar

	While ("TMP1")->(!EOF())
		//Incrementa a régua
		IncProc()
		
		cLin := "1+1+I["
		cLin += ALLTRIM(TMP1->RA_PIS)
		cLin += "["
		cLin += RTRIM(TMP1->RA_NOME)
		cLin += "[1[1["
		cLin += ALLTRIM(TMP1->RA_MAT)
		cLin += CRLF
		
		If fWrite(nHdl,cLin,Len(cLin)) != Len(cLin)
			If !MsgAlert("Ocorreu um erro na gravação do arquivo.","Continua?","Atenção!")
				Exit
			Endif
		Endif
		("TMP1")->(dbSkip())
	EndDo 
	fClose(nHdl)
	dbCloseArea()
Return Nil     

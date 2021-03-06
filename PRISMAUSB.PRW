
#include "protheus.ch"

/*/
+-----------------------------------------------------------------------------
| Programa | PrismaUSB | Autor | FLAVIO MARTINS | 25/08/2015 | |
+-----------------------------------------------------------------------------
| Descri��o | Este programa vai gerar um arquivo texto com os dados dos funcionarios
  para ser importado via USB no rel�gio de ponto PRISMA Super F�cil R02 |
+-----------------------------------------------------------------------------
| Uso | Ponto Eletr�nico ADVPL |
+-----------------------------------------------------------------------------
/*/


User Function PrismaUSB()

	Local oPrismaUSB
	Private cPerg := "PRISMA"
	Private cAlias := "SRA"
	Public TMP1

	//CriaSx1(cPerg)
	Pergunte(cPerg,.F.)

	DEFINE MSDIALOG oPrismaUSB TITLE OemToAnsi("Gera��o de Arquivo Texto") ;
	FROM 000,000 TO 200,400 PIXEL
	//@ 005,005 TO 095,195 OF oPrismaUSB PIXEL 
	@ 005,005 TO 095,195 OF oPrismaUSB PIXEL 
	@ 010,010 Say " Este programa vai gerar um arquivo texto para ser importado via USB ";
	OF oPrismaUSB PIXEL 
	@ 018,010 Say " no rel�gio de ponto PRISMA Super F�cil R02 das filiais FVC e FCB. ";
	OF oPrismaUSB PIXEL
	@ 026,010 Say " O arquivo ser� gravado na pasta C:\Prisma\ verifique se esta pasta existe.";
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
		MsgAlert("O arquivo de nome "+cArqTxt+" n�o pode ser executado! Verifique os par�metros.","Aten��o!")
		Return
	Endif
	// Inicializa a r�gua de processamento
	Processa({|| RunCont() },"Processando...")

Return Nil

Static Function RunCont
	Local cLin
	ProcRegua(RecCount()) // Numero de registros a processar

	While ("TMP1")->(!EOF())
		//Incrementa a r�gua
		IncProc()
		
		cLin := "1+1+I["
		cLin += ALLTRIM(TMP1->RA_PIS)
		cLin += "["
		cLin += RTRIM(TMP1->RA_NOME)
		cLin += "[1[1["
		cLin += ALLTRIM(TMP1->RA_MAT)
		cLin += CRLF
		
		If fWrite(nHdl,cLin,Len(cLin)) != Len(cLin)
			If !MsgAlert("Ocorreu um erro na grava��o do arquivo.","Continua?","Aten��o!")
				Exit
			Endif
		Endif
		("TMP1")->(dbSkip())
	EndDo 
	fClose(nHdl)
	dbCloseArea()
Return Nil     

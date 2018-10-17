#include 'protheus.ch'
#include 'parmtype.ch'

/*/{Protheus.doc} U_TREPOCORR()
//TODO Relat�rio de ocorr�ncias utilizando o objeto TReport
@author Flavio Martins
@since 22/05/2018
@version 1.0
@return ${return}, ${return_description}

@type function
/*/
user function TREPOCORR()
	Local oReport
	
	If TRepInUse()
		Pergunte("FILCC",.F.)
		oReport := ReportDef()
		oReport:PrintDialog()
	EndIf
return


Static Function ReportDef()
	Local oReport
	Local oSection
	Local oBreak
	
	oReport := TReport():New("RelOcorrencia","Relatorio de Ocorr�ncia","FILCC",{|oReport|;
	PrintReport(oReport)},"Relatorio de Ocorr�ncia")
	oSection := TRSection():New(oReport,"Funcion�rios",{"SRA","SRJ"})
	
	TRCell():New(oSection,"RA_FILIAL","SRA","Filial")
	TRCell():New(oSection,"RA_MAT","SRA","Matricula")
	TRCell():New(oSection,"RA_CC","SRA","Centro de Custo")
	TRCell():New(oSection,"RA_NOME","SRA","Nome")
	TRCell():New(oSection,"RA_ADMISSA","SRA","Dt. Admiss�o")
	TRCell():New(oSection,"RA_SITFOLH","SRA","Sit. Folha")
	TRCell():New(oSection,"RJ_DESC","SRJ","Desc. Fun��o")
	TRCell():New(oSection,"RA_PERICUL","SRA","Horas Periculosidade")
	TRCell():New(oSection,"RA_ADCINS","SRA","Possui  Insalubridade")
	TRCell():New(oSection,"RA_INSMAX","SRA","Hrs. Insalubridade")
	TRCell():New(oSection,"RA_OCORREN","SRA","Ocorrencia de Ag. nocivo")
	TRCell():New(oSection,"RA_PERCSAT","SRA","% Acidente do Trabalho")
	

	TRFunction():New(oSection:Cell("RA_MAT"),NIL,"COUNT",oBreak)
Return oReport

Static Function PrintReport(oReport)
	Local oSection := oReport:Section(1)
	Local cPart
	Local cFiltro := ""
	
	 MakeSqlExpr("FILCC")
	 oSection:BeginQuery()
	 
	 BeginSql alias "QRYSRA"
		SELECT RA_FILIAL, RA_MAT, RA_CC, RA_NOME, RA_ADMISSA, RA_SITFOLH, RJ_DESC, RA_PERICUL, RA_ADCINS, RA_INSMAX, RA_OCORREN, RA_PERCSAT
		FROM %table:SRA% 
		INNER JOIN %table:SRJ% ON RA_FILIAL = RJ_FILIAL AND RA_CODFUNC = RJ_FUNCAO AND %table:SRJ%.%notDel%
		WHERE %table:SRA%.%notDel%
			AND SRA010.RA_SITFOLH <> 'D' 
			AND SRA010.RA_FILIAL BETWEEN %Exp:mv_par01% AND %Exp:mv_par02%
			AND SRA010.RA_CC BETWEEN %Exp:mv_par03% AND %Exp:mv_par04%
		ORDER BY %Order:SRA,1%
	 EndSql
	 
	 oSection:EndQuery()
	 oSection:Print()
Return
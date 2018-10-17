#include 'protheus.ch'
#include 'parmtype.ch'

/*/{Protheus.doc} U_TREPFERIAS()
//TODO Relatório de Férias utilizando o objeto TReport.
@author Flavio Martins
@since 22/05/2018
@version 1.0
@return ${return}, ${return_description}

@type function
/*/
user function TREPFERIAS()
	Local oReport
	
	If TRepInUse()
		Pergunte("FILCCFER",.F.)
		oReport := ReportDef()
		oReport:PrintDialog()
	EndIf
return


Static Function ReportDef()
	Local oReport
	Local oSection
	Local oBreak
	
	oReport := TReport():New("RelFerias","Relatório de Férias","FILCCFER",{|oReport|;
	PrintReport(oReport)},"Relatório de Férias")
	oSection := TRSection():New(oReport,"Funcionários",{"SRA","SRJ","SRF"})
	TRCell():New(oSection,"RA_FILIAL","SRA","Filial")
	TRCell():New(oSection,"RA_MAT","SRA","Matricula")
	TRCell():New(oSection,"RA_CC","SRA","Centro de Custo")
	TRCell():New(oSection,"RA_DEPTO","SRA","Cod. Departamento")
	TRCell():New(oSection,"RA_NOME","SRA","Nome")
	TRCell():New(oSection,"RA_ADMISSA","SRA","Dt. Admissão")
	TRCell():New(oSection,"RA_SITFOLH","SRA","Sit. Folha")
	TRCell():New(oSection,"RJ_DESC","SRJ","Desc. Função")
	TRCell():New(oSection,"RF_DATABAS","SRF","Dt Base Ferias")
	TRCell():New(oSection,"DTFINAL","SRF","Dt Final")
	TRCell():New(oSection,"RF_DFERVAT","SRF","Dias Fer. Vencidas")
	TRCell():New(oSection,"RF_DFALVAT","SRF","Dias Fer. Prop")
	TRCell():New(oSection,"RF_DFALAAT","SRF","Dias Fal. Prop")
	TRFunction():New(oSection:Cell("RA_MAT"),NIL,"COUNT",oBreak)
Return oReport

Static Function PrintReport(oReport)
	Local oSection := oReport:Section(1)
	Local cPart
	Local cFiltro := ""
	
	 MakeSqlExpr("FILCCFER")
	 oSection:BeginQuery()
	 
	 BeginSql alias "QRYSRA"
		SELECT RA_FILIAL, RA_MAT, RA_CC, RA_NOME, RA_ADMISSA, RA_SITFOLH, RJ_DESC, RF_DATABAS, dateadd(day,300,RF_DATABAS) AS DTFINAL, RF_DFERVAT, RF_DFERAAT, RF_DFALVAT, RA_DEPTO
		FROM %table:SRA% 
		INNER JOIN %table:SRJ% ON RA_FILIAL = RJ_FILIAL AND RA_CODFUNC = RJ_FUNCAO AND %table:SRJ%.%notDel%
		INNER JOIN %table:SRF% ON RA_FILIAL = RF_FILIAL AND RA_MAT = RF_MAT AND %table:SRF%.%notDel%
		WHERE %table:SRA%.%notDel%
			AND RA_SITFOLH <> 'D'
			AND SRA010.RA_FILIAL BETWEEN %Exp:mv_par01% AND %Exp:mv_par02%
			AND SRA010.RA_CC BETWEEN %Exp:mv_par03% AND %Exp:mv_par04%
			AND SRA010.RA_DEPTO BETWEEN %Exp:mv_par07% AND %Exp:mv_par08%
			AND SRF010.RF_DATABAS BETWEEN %Exp:mv_par05% AND %Exp:mv_par06%
		ORDER BY %Order:SRA,1%
	 EndSql
	 
	 oSection:EndQuery()
	 oSection:Print()
Return
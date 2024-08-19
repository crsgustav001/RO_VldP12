#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "TBICONN.CH"

/*/{Protheus.doc} User Function RO_VldP12
	Gera planilha excel Planilha com rotinas descontinuadas
	@type Function
	@author Cristian Gustavo Dias Andrade 
	@use R_VENDAS
	@since 19/08/2024
	@version 1.1
/*/
User Function RO_VldP12()

	Local lFicaTela 	AS Logical
	Local oGet1     	AS Object
	Local oGet2     	AS Object

	Local oModal    	AS Object
	Local oPanel1   	AS Object
	Local oSay1     	AS Object
	Local oSay2     	AS Object
	Local oDlg1      	AS Object
	Local oButtonBar 	AS Object

	Private cGrpCliDe 	AS Character
	Private cGrpCliAte 	AS Character
	Private oProcess 	AS Object

	lFicaTela 		:= .T.
	oGet1     		:= NIL
	oGet2     		:= NIL

	oModal    		:= NIL
	oPanel1   		:= NIL
	oSay1     		:= NIL
	oSay2     		:= NIL
	oDlg1      		:= NIL
	oButtonBar 		:= NIL

	cGrpCliDe		:= Space( 6 )
	cGrpCliAte		:= "ZZZZZZ"
	oProcess 		:= NIL

	While lFicaTela

		oModal := FWDialogModal():New()
		oModal:SetTitle( "Parâmetros Planilha Grupo Cliente")
		oModal:SetSubTitle( "" )
		oModal:SetSize( 140, 140 ) // Seta a altura e largura da janela em pixel //
		oModal:SetEscClose( .T. )
		oModal:EnableFormBar( .T. )
		oModal:CreateDialog()
		oModal:CreateFormBar()

		oPanel1 := oModal:GetPanelMain()

		@ 010, 010 SAY oSay1 PROMPT "Grupo Cliente De:" SIZE 00, 20 OF oPanel1 PIXEL
		@ 030, 010 SAY oSay2 PROMPT "Grupo Cliente Até:" SIZE 00, 20 OF oPanel1 PIXEL


		oSay1:SetCSS( CSS_SAY )
		oSay2:SetCSS( CSS_SAY )

		@ 008, 075 MSGET oGet1 VAR cGrpCliDe PICTURE "@!" F3 "SA1" SIZE 50, 12 OF oPanel1 PIXEL
		@ 028, 075 MSGET oGet2 VAR cGrpCliAte PICTURE "@!" F3 "SA1" SIZE 50, 12 OF oPanel1 PIXEL

		oGet1:SetCSS( CSS_GET )
		oGet2:SetCSS( CSS_GET )

		oModal:AddButton( "&Sair"          , { || lFicaTela := .F., oModal:oOwner:End() }, "", NIL, .T., .F., .T., NIL )
		oModal:AddButton( "&Gerar Planilha", { || fExecDes()                   }, "", NIL, .T., .F., .T., NIL )

		oModal:SetInitBlock( { || oGet1:SetFocus() } )

		oModal:Activate()
	EndDo
Return( NIL )

/*/{Protheus.doc} fExecDes
	Função para executar as descargas.
	@type Function
	@author Cristian Gustavo Dias Andrade 
	@since 19/08/2024
	@version 1.1
/*/
Static Function fExecDes()
	oProcess := MsNewProcess():New( { || fAxExecDes() }, "Aguarde...", "Gerando arquivos..." )
	oProcess:Activate()
Return( NIL )

/*/{Protheus.doc} fAxExecDes
	Função auxiliar para executar as descargas.
	@type Function
	@author Cristian Gustavo Dias Andrade 
	@since 19/08/2024
	@version 1.1
/*/
Static Function fAxExecDes()

	oProcess:SetRegua1( 2 )
	oProcess:IncRegua1()

	fDados()

Return( NIL )

/*/{Protheus.doc} fDados
	Dados para gerar as planilhas.
	@type Function
	@author Cristian Gustavo Dias Andrade 
	@since 19/08/2024
	@version 1.1
	@param pI, Numeric, opção escolhida
/*/
//Static Function fDados( pI )
Static Function fDados()

	Local cAlias 	 AS Character

	cAlias 		:= GetNextAlias()

	// Checa se arquivo ja esta aberto
	IF(select( cAlias) > 0)
		(cAlias)->(DbCloseArea() )
	ENDIF

	BeginSql Alias cAlias

	%NoParser%

	SELECT 
		ZC9_FILIAL AS FL,
    	ZC9_CLIENT AS GR_CLI,
    	SEU_CODIGO.COD_C AS SEU_CODIGO,
    	ZC9_DTINI AS DT_INI,
    	ZC9_DTFIM AS DT_FIM,
    	TRIM(ZC9_PRODUT) AS PRODUTO,
    	TRIM(B1_DESC)||' '||TRIM(B1_X_DESC) AS DESCRICAO,
    	TRIM(B1_FABRIC) AS MONT,
    	ZC9_PREFIX AS PR_FIXADO_AGORA,
    	ROUND(ZC9010.ZC9_PREFIX / (CASE
                                      WHEN A5_X_STSPA > 0 THEN A5_X_STSPA
                                      ELSE A5_X_CUSTD
                                  END), 2) AS MARGEM_REAL_AGORA,
       ZC9_MARGEM AS ZC9_MARGEM,
       A5_FORNECE AS A5_FORNECE,
       A5_LOJA AS A5_LOJA,
       A5_X_CURVA AS CURVA,
       A5_PRODUTO AS A5_PRODUTO,
       (CASE
            WHEN A5_X_STSPA > 0 THEN A5_X_STSPA
            ELSE A5_X_CUSTD
        END) AS CUSTOTELA_ATUAL,
       A5_X_MARPA AS M_FIXA_ATUAL_TELA_CUSTO,
       A5_X_PRVEN AS P_VENDA_ATUAL_TELA_CUSTO,
       (CASE
            WHEN A5_X_STSPF > 0 THEN A5_X_STSPF
            ELSE A5_X_CSTDF
        END) AS CUSTOTELA_FUTURO,
       A5_X_MARPF AS M_FIXA_FUTURO,
       A5_X_CATFU AS P_VENDA_FUTURO,
       ZC9_USUARI AS USUARIO,
       ZC9010.ZC9_DTMANU AS DT_MANUT,
       COUNT(*) OVER() AS NUM_ROWS
	FROM ZC9010,
    	SB1010,
    	SA5010,
  	(SELECT ZS_CODCLI CLI,
          ZS_PROCLI COD_C,
          ZS_CODRAIL COD_R
   	FROM SZS010
   	WHERE ZS_FILIAL =' '
    	AND D_E_L_E_T_ =' '
   	GROUP BY ZS_CODCLI,
            ZS_PROCLI,
            ZS_CODRAIL) SEU_CODIGO
	WHERE ZC9_FILIAL = %Exp:cFilAnt% /*Filial logada*/
		AND ZC9_CLIENT >= %Exp:cCmpCliDe%
		AND ZC9_CLIENT <= %Exp:cCmpCliAte%
		AND ZC9_DTINI >= %Exp:cDataDe%
		AND ZC9_DTFIM <= %Exp:cDataAte%
		AND ZC9_DTCAN =' '
  		AND ZC9010.D_E_L_E_T_ =' '
  		AND ZC9_PRODUT = B1_COD
  		AND B1_FILIAL =' '
  		AND SB1010.D_E_L_E_T_ =' '
  		AND SUBSTR(ZC9_PRODUT, 1, 7) = SUBSTR(A5_PRODUTO(+), 1, 7)
  		AND ZC9_FILIAL = A5_FILIAL(+)
  		AND A5_X_LVDAS(+) ='1'
  		AND A5_X_LCPAS(+) ='1'
  		AND A5_X_TIPOF(+) ='P'
  		AND SA5010.D_E_L_E_T_(+) =' '
  		AND SUBSTR(ZC9_PRODUT, 1, 7) = SUBSTR(SEU_CODIGO.COD_R(+), 1, 7)
  		AND ZC9_CLIENT = SEU_CODIGO.CLI(+)
  		AND ZC9_DTCAN =' '
	ORDER BY ZC9_PRODUT, A5_PRODUTO;
	
	EndSql

	IF EMPTY('dDataDe') .OR. EMPTY('')
		FWAlertWarning( "Todos os parâmetros devem ser informados!!!", "ATENÇÃO !!!" )
		Return NIL

	ELSE
		fPlanilha( cAlias )
	EndIf

	(cAlias)->( dbCloseArea() )

Return( NIL )

/*/{Protheus.doc} fPlanilha
	Gera as planilhas.
	@type Function
	@author Cristian Gustavo Dias Andrade 
	@since 19/08/2024
	@version 1.1
	@param pI, Numeric, Opção escolhida
	@param pAlias, Character, Alias com os dados
/*/
//Static Function fPlanilha( pI, pAlias )
Static Function fPlanilha( pAlias )

	Local aLinha     	AS Array
	Local aCposStru     AS Array
	Local cCodFil   	AS Character
	Local cNomeArq   	AS Character
	Local cDirTmp   	AS Character
	Local cTable     	AS Character
	Local cWorkSheet 	AS Character
	Local nI         	AS Numeric
	Local nY         	AS Numeric
	Local oExlXlsx   	AS Object
	Local oOpnXlsx   	AS Object

	aLinha     		:= {}
	aCposStru     	:= {}
	cCodFil   		:= FWCodFil()
	cNomeArq   		:= ''
	cTable     		:= ''
	cWorkSheet 		:= ''
	nI         		:= 0
	nY         		:= 0
	oExlXlsx   		:= NIL
	oOpnXlsx   		:= NIL

	cDirTmp    := "C:\RELATORIOS\"
	cNomeArq   := "Grupo Cliente Fl " + cCodFil // retiradas as virgulas e o "dois pontos" //
	cNomeArq   := cDirTmp + StrTran( cNomeArq, " ", "_" ) + ".xlsx"
	cTable     := "Planilha Grupo Cliente"
	cWorkSheet := "Grupo Cliente"

	oExlXlsx := FwMsExcelXlsx():New()
	oExlXlsx:AddworkSheet( cWorkSheet )
	oExlXlsx:AddTable( cWorkSheet, cTable )

	//Carregamento e estruturação dos campos
	aCposStru := fCposStru()

	for nY := 1 to LEN(aCposStru)

		If aCposStru[nY][1] == "NUM_ROWS" // Não considera o campo NUM_ROWS
			Loop
		ElseIf aCposStru[nY][1] == "PRODUTO"
			aCposStru[nY][1] := "Produto"
		ElseIf aCposStru[nY][1] == "ZC9_MARGEM"
			aCposStru[nY][1] := "Margem desejada"
		ElseIf aCposStru[nY][1] == "A5_FORNECE"
			aCposStru[nY][1] := "Fornecedor"
		ElseIf aCposStru[nY][1] == "A5_LOJA"
			aCposStru[nY][1] := "Loja"
		ElseIf aCposStru[nY][1] == "A5_PRODUTO"
			aCposStru[nY][1] := "Produto Custo"
		EndIf

		//aCposStru[nY][1] == Nome Campo
		//aCposStru[nY][2] == Nº orientação campo
		//aCposStru[nY][3] == Codigo formatação campo
		//aCposStru[nY][4] == Opção totalização campo
		//aCposStru[nY][5] == Mascara campo

		oExlXlsx:AddColumn(cWorkSheet, cTable, aCposStru[nY][1], aCposStru[nY][2], aCposStru[nY][3], aCposStru[nY][4], aCposStru[nY][5])
	next

	oProcess:SetRegua2( (pAlias)->NUM_ROWS )

	While ! (pAlias)->( EOF() )

		oProcess:IncRegua2()

		aLinha := {}

		For nI := 1 To (pAlias)->( FCount() )
			If (pAlias)->( FieldName( nI ) ) <> "NUM_ROWS"

				IF (pAlias)->( FieldName( nI )) == 'DT_INI' .OR. (pAlias)->( FieldName( nI )) == 'DT_FIM' .OR. (pAlias)->( FieldName( nI )) == 'DT_MANUT'
					aAdd( aLinha,;
						IIF(EMPTY((pAlias)->( FieldGet( nI ) )), '', STOD((pAlias)->( FieldGet( nI ) ))) ;
						) //Conversão formato data para xx/xx/xxxx.
				ELSE
					aAdd( aLinha, (pAlias)->( FieldGet(nI) ) )
				EndIf

			EndIf
		Next nI

		oExlXlsx:AddRow( cWorkSheet, cTable, aLinha )

		(pAlias)->( dbSkip() )

	EndDo

	FErase( cNomeArq )

	//Criação planilha
	oExlXlsx:SetItalic( .T. )
	oExlXlsx:SetBold( .T. )
	oExlXlsx:SetUnderline( .T. )
	oExlXlsx:Activate()
	oExlXlsx:GetXMLFile( cNomeArq )

	//Exibição planilha
	oOpnXlsx := MsExcel():New()
	oOpnXlsx:WorkBooks:Open( cNomeArq )
	oOpnXlsx:SetVisible( .T. )
	oOpnXlsx:Destroy()
	oExlXlsx:DeActivate()

Return NIL

/*/{Protheus.doc} fCposStru
	Função que retorna a estrutura dos campos na planilha.
	@type function
	@version 1.1
	@author Cristian Gustavo Dias Andrade (Franco Consultoria)
	@since 19/08/2024
	@return Array, Retorna estrutura dos campos.
/*/
Static Function fCposStru()

	Local aRet 		AS Array
	Local nPosCmp	AS Numeric

	Local cMasNum	 AS Character
	Local cMasNumTot AS Character
	Local cMasChr	 AS Character
	Local cMasDat	 AS Character

	Local nPadChr	AS Numeric
	Local nPadNum	AS Numeric
	Local nPadDat	AS Numeric

	aRet 		:= {}
	nPosCmp 	:= 1

	cMasNum 	:= "@E 999,999,999"
	cMasNumTot	:= "@E 999,999,999.99"
	cMasMargem	:= "@E 999,999,999.9999999"
	cMasChr     := '@!'
	cMasDat     := "@D"

	nPadChr		:= 1
	nPadNum 	:= 2
	nPadDat 	:= 4

	//Titulo tabela
	//Alinhamento coluna: ( 1-Left,2-Center,3-Right )
	//Codigo de formatação ( 1-General,2-Number,3-Monetário,4-DateTime )
	//Indica se a coluna deve ser totalizada	.F.
	//Mascara de picture a ser aplicada. Somente para campos numéricos

	//Caracter
	aAdd(aRet, { "FL", nPosCmp, nPadChr, .F., ''} )
	aAdd(aRet, { "GR_CLI", nPosCmp, nPadChr, .F., ''} )
	aAdd(aRet, { "SEU_CODIGO", nPosCmp, nPadChr, .F., ''} )

	//Data
	aAdd(aRet, { "DT_INI", nPosCmp, nPadDat, .F., cMasDat} )
	aAdd(aRet, { "DT_FIM", nPosCmp, nPadDat, .F., cMasDat} )

	//Caracter
	aAdd(aRet, { "PRODUTO", nPosCmp, nPadChr, .F., ''} )
	aAdd(aRet, { "DESCRICAO", nPosCmp, nPadChr, .F., ''} )
	aAdd(aRet, { "MONT", nPosCmp, nPadChr, .F., ''} )

	//Numerico
	aAdd(aRet, { "PR_FIXADO_AGORA", nPosCmp, nPadNum, .F., cMasNumTot} )
	aAdd(aRet, { "MARGEM_REAL_AGORA", nPosCmp, nPadNum, .F., cMasNumTot} )
	aAdd(aRet, { "ZC9_MARGEM", nPosCmp, nPadNum, .F., cMasNumTot} )

	//Caracter
	aAdd(aRet, { "A5_FORNECE", nPosCmp, nPadChr, .F., ''} )
	aAdd(aRet, { "A5_LOJA", nPosCmp, nPadChr, .F., ''} )
	aAdd(aRet, { "CURVA", nPosCmp, nPadChr, .F., ''} )
	aAdd(aRet, { "A5_PRODUTO", nPosCmp, nPadChr, .F., ''} )

	//Numerico, alinhado a esquerda
	aAdd(aRet, { "CUSTOTELA_ATUAL", 3, nPadNum, .F., cMasNumTot} )
	aAdd(aRet, { "M_FIXA_ATUAL_TELA_CUSTO", 3, nPadNum, .F., cMasMargem} )
	aAdd(aRet, { "P_VENDA_ATUAL_TELA_CUSTO", 3, nPadNum, .F., cMasNumTot} )
	aAdd(aRet, { "CUSTOTELA_FUTURO", 3, nPadNum, .F., cMasNumTot} )
	aAdd(aRet, { "M_FIXA_FUTURO", 3, nPadNum, .F., cMasMargem} )
	aAdd(aRet, { "P_VENDA_FUTURO", 3, nPadNum, .F., cMasNumTot} )

	//Caracter
	aAdd(aRet, { "USUARIO", nPosCmp, nPadChr, .F., ''} )

	//Data
	aAdd(aRet, { "DT_MANUT", nPosCmp, nPadDat, .F., cMasDat} )

	//Numerico
	aAdd(aRet, { "NUM_ROWS", nPosCmp, nPadNum, .F.,cMasNum} )

Return( aRet )

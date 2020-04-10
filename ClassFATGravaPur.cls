VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Public Function RelFormPrecoPur_Prepara(ByVal iFilialEmpresa As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal sModelo As String, ByVal sNomeDiretorio As String, ByVal sNomeArquivo As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 15) As Long
Dim lNumIntRel As Long
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
'Dim objPlanilhaRange As Object 'Excel.Worksheet
Dim iLinha As Integer, iColuna As Integer
Dim sRegistro As String, sParam As String, dValorAux As Double
Dim iPos1 As Integer, iPos2 As Integer, iPos3 As Integer
Dim bArqAberto As Boolean, sConta As String
Dim dQtdeAprop As Double, dCustoMov As Double, dQtdeMov As Double, dCusto As Double
Dim bAlt As Boolean, lTransacao As Long, bFormataNum As Boolean
Dim objExcelApp As New ClassExcelApp, iNumAbas As Integer
Dim sSQL As String, lNumIntItemNF As Long, lNumIntNF As Long, lNumIntCTe As Long
Dim dtDataNF As Date, dCustoCTe As Double, dQtdeCTe As Double, dQtdeNF As Double, dCustoCTeAux As Double

Const NUM_FORMATO = "#,##0.00"
Const NUM_FORMATO_CUSTO = "#,##0.0000"

On Error GoTo Erro_RelFormPrecoPur_Prepara

    lErro = CF("RelVendEstProd_Prepara", lNumIntRel, iFilialEmpresa, dtDataDe, dtDataAte)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211627

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211581
    Next
    
    sSQL = "INSERT INTO RelFormPrecoPurAux (NumIntRel, Produto, Qtde, Valor, CustoMedio)( "
    sSQL = sSQL & "SELECT ?, Produto, SUM(Qtde), SUM(Valor), CASE WHEN SUM(Qtde) = 0 THEN 0 ELSE SUM(CustoTotal) / SUM(Qtde) END FROM RelComprasPur "
    sSQL = sSQL & "WHERE FilialEmpresa = ? AND Data BETWEEN ? AND ? "
    sSQL = sSQL & "GROUP BY Produto) "
    
    lErro = Comando_Executar(alComando(8), sSQL, lNumIntRel, iFilialEmpresa, dtDataDe, dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 211586
    
'    'Abre o excel
'    lErro = CF("Excel_Abrir")
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'    'Desabilita as mensagens do Excel
'    GL_objExcelSist.DisplayAlerts = False
'
'    Call GL_objExcelSist.Workbooks.Open(sModelo)
'    bArqAberto = True
'
'    Set objPastaTrabalho = GL_objExcelSist.ActiveWorkBook
        
    'Seleciona a planilha ativa na pasta de trabalho criada
    'Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
    
    lErro = objExcelApp.Abrir_Planilha(sModelo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    bArqAberto = True
    
    'Salvar como em outro lugar
    lErro = objExcelApp.Salvar(Replace(sModelo, ".x", "_" & Format(Date, "YYYYMMDD") & "_" & Format(Time, "HHMMSS") & ".x"))
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call objExcelApp.Mudar_Visibilidade(True)
    
    'For Each objPlanilhaExcel In objPastaTrabalho.Sheets
    
    iNumAbas = objExcelApp.Qtde_Abas
    
    For iIndice = 1 To iNumAbas
    
        lErro = objExcelApp.Planilha_Muda_Aba(iIndice)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        For iLinha = 1 To 100
            For iColuna = 1 To 100
            
                bAlt = False
                bFormataNum = True
            
                'If Not IsError(objPlanilhaExcel.cells(iLinha, iColuna)) Then
                If objExcelApp.Obtem_Valor_Celula(iLinha, iColuna) <> "" Then
                
                    'sRegistro = objPlanilhaExcel.cells(iLinha, iColuna)
                    sRegistro = objExcelApp.Obtem_Valor_Celula(iLinha, iColuna)
                    
                    If InStr(1, sRegistro, "@Faixa_Datas") <> 0 Then
                        bAlt = True
                        sRegistro = Replace(sRegistro, "@Faixa_Datas", "De: " & Format(dtDataDe, "dd/mm/yyyy") & " Até: " & Format(dtDataAte, "dd/mm/yyyy"))
                        bFormataNum = False
                    End If
                    If InStr(1, sRegistro, "@Venda_Pc") <> 0 Then
                        iPos1 = InStr(1, sRegistro, "@Venda_Pc")
                        iPos2 = InStr(iPos1, sRegistro, "(")
                        iPos3 = InStr(iPos1, sRegistro, ")")
                        sParam = Mid(sRegistro, iPos2 + 1, iPos3 - iPos2 - 1)
                        
                        lErro = Comando_Executar(alComando(0), "SELECT QtdeVend FROM RelVendEstProd WHERE NumIntRel = ? AND Produto = ? ", dValorAux, lNumIntRel, sParam)
                        If lErro <> AD_SQL_SUCESSO Then gError 211582
                        
                        lErro = Comando_BuscarPrimeiro(alComando(0))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                        
                        If lErro = AD_SQL_SUCESSO Then
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Venda_Pc(" & sParam & ")", Format(dValorAux, NUM_FORMATO))
                        Else
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Venda_Pc(" & sParam & ")", Format(0, NUM_FORMATO))
                        End If
                    End If
                    If InStr(1, sRegistro, "@Producao_Pc") <> 0 Then
                        iPos1 = InStr(1, sRegistro, "@Producao_Pc")
                        iPos2 = InStr(iPos1, sRegistro, "(")
                        iPos3 = InStr(iPos1, sRegistro, ")")
                        sParam = Mid(sRegistro, iPos2 + 1, iPos3 - iPos2 - 1)
                        
                        lErro = Comando_Executar(alComando(1), "SELECT QtdeProd FROM RelVendEstProd WHERE NumIntRel = ? AND Produto = ? ", dValorAux, lNumIntRel, sParam)
                        If lErro <> AD_SQL_SUCESSO Then gError 211584
                        
                        lErro = Comando_BuscarPrimeiro(alComando(1))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211585
                        
                        If lErro = AD_SQL_SUCESSO Then
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Producao_Pc(" & sParam & ")", Format(dValorAux, NUM_FORMATO))
                        Else
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Producao_Pc(" & sParam & ")", Format(0, NUM_FORMATO))
                        End If
                    End If
                    If InStr(1, sRegistro, "@Producao_Rs") <> 0 Then
                        iPos1 = InStr(1, sRegistro, "@Producao_Rs")
                        iPos2 = InStr(iPos1, sRegistro, "(")
                        iPos3 = InStr(iPos1, sRegistro, ")")
                        sParam = Mid(sRegistro, iPos2 + 1, iPos3 - iPos2 - 1)
                        
                        lErro = Comando_Executar(alComando(2), "SELECT QtdeProd * (VlrVend/QtdeVend) FROM RelVendEstProd WHERE NumIntRel = ? AND Produto = ? AND QtdeVend > 0", dValorAux, lNumIntRel, sParam)
                        If lErro <> AD_SQL_SUCESSO Then gError 211586
                        
                        lErro = Comando_BuscarPrimeiro(alComando(2))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211587
                        
                        If lErro = AD_SQL_SUCESSO Then
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Producao_Rs(" & sParam & ")", Format(dValorAux * 0.75, NUM_FORMATO))
                        Else
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Producao_Rs(" & sParam & ")", Format(0, NUM_FORMATO))
                        End If
                    End If
                    If InStr(1, sRegistro, "@Venda_Rs") <> 0 Then
                        iPos1 = InStr(1, sRegistro, "@Venda_Rs")
                        iPos2 = InStr(iPos1, sRegistro, "(")
                        iPos3 = InStr(iPos1, sRegistro, ")")
                        sParam = Mid(sRegistro, iPos2 + 1, iPos3 - iPos2 - 1)
                        
                        lErro = Comando_Executar(alComando(3), "SELECT VlrVend FROM RelVendEstProd WHERE NumIntRel = ? AND Produto = ? ", dValorAux, lNumIntRel, sParam)
                        If lErro <> AD_SQL_SUCESSO Then gError 211588
                        
                        lErro = Comando_BuscarPrimeiro(alComando(3))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211589
                                           
                        If lErro = AD_SQL_SUCESSO Then
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Venda_Rs(" & sParam & ")", Format(dValorAux, NUM_FORMATO))
                        Else
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Venda_Rs(" & sParam & ")", Format(0, NUM_FORMATO))
                        End If
                    End If
                    If InStr(1, sRegistro, "@Valor_Cta") <> 0 Then
                        iPos1 = InStr(1, sRegistro, "@Valor_Cta")
                        iPos2 = InStr(iPos1, sRegistro, "(")
                        iPos3 = InStr(iPos1, sRegistro, ")")
                        sParam = Mid(sRegistro, iPos2 + 1, iPos3 - iPos2 - 1)
                        
                        sConta = sParam & String(10 - Len(sParam), "0")
                        
                        lErro = Comando_Executar(alComando(4), "SELECT SUM(Cre-Deb) FROM MvDiaCta WHERE FilialEmpresa = ? AND Conta = ? AND Data BETWEEN ? AND ? ", dValorAux, iFilialEmpresa, sConta, dtDataDe, dtDataAte)
                        If lErro <> AD_SQL_SUCESSO Then gError 211590

                        lErro = Comando_BuscarPrimeiro(alComando(4))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211591
                        
'                        If left(sConta, 1) = "2" Or left(sConta, 1) = "1" Then
'                            lErro = Comando_Executar(alComando(4), "SELECT SUM(Cre) FROM MvDiaCta WHERE FilialEmpresa = ? AND Conta = ? AND Data BETWEEN ? AND ? ", dValorAux, iFilialEmpresa, sConta, dtDataDe, dtDataAte)
'                            If lErro <> AD_SQL_SUCESSO Then gError 211590
'
'                            lErro = Comando_BuscarPrimeiro(alComando(4))
'                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211591
'                        Else
'                            lErro = Comando_Executar(alComando(7), "SELECT SUM(Deb) FROM MvDiaCta WHERE FilialEmpresa = ? AND Conta = ? AND Data BETWEEN ? AND ? ", dValorAux, iFilialEmpresa, sConta, dtDataDe, dtDataAte)
'                            If lErro <> AD_SQL_SUCESSO Then gError 211590
'
'                            lErro = Comando_BuscarPrimeiro(alComando(7))
'                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211591
'                        End If
                                                                       
                        If lErro = AD_SQL_SUCESSO Then
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Valor_Cta(" & sParam & ")", Format(Abs(dValorAux), NUM_FORMATO))
                        End If
                        
                    End If
                    If InStr(1, sRegistro, "@Cust_Prod_Emb") <> 0 Then
                        iPos1 = InStr(1, sRegistro, "@Cust_Prod_Emb")
                        iPos2 = InStr(iPos1, sRegistro, "(")
                        iPos3 = InStr(iPos1, sRegistro, ")")
                        sParam = Mid(sRegistro, iPos2 + 1, iPos3 - iPos2 - 1)
                        dCusto = 0
                    
                        lErro = Comando_Executar(alComando(5), "SELECT QtdeAprop, CustoMov, QtdeMov FROM RelFormPrecoAprop WHERE FilialEmpresa = ? AND Produto = ? AND Data BETWEEN ? AND ? AND Natureza = ?", dQtdeAprop, dCustoMov, dQtdeMov, iFilialEmpresa, sParam, dtDataDe, dtDataAte, NATUREZA_PROD_EMBALAGENS)
                        If lErro <> AD_SQL_SUCESSO Then gError 211592
                        
                        lErro = Comando_BuscarPrimeiro(alComando(5))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211593
                        
                        Do While lErro <> AD_SQL_SEM_DADOS
                        
                            dCusto = dCusto + dQtdeAprop * (dCustoMov / dQtdeMov)
                        
                            lErro = Comando_BuscarProximo(alComando(5))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211594
                        
                        Loop
                    
                        bAlt = True
                        sRegistro = Replace(sRegistro, "@Cust_Prod_Emb(" & sParam & ")", Format(dCusto, NUM_FORMATO))
                    
                    End If
                    If InStr(1, sRegistro, "@Cust_Prod_MP") <> 0 Then
                        iPos1 = InStr(1, sRegistro, "@Cust_Prod_MP")
                        iPos2 = InStr(iPos1, sRegistro, "(")
                        iPos3 = InStr(iPos1, sRegistro, ")")
                        sParam = Mid(sRegistro, iPos2 + 1, iPos3 - iPos2 - 1)
                        dCusto = 0
                    
                        lErro = Comando_Executar(alComando(5), "SELECT QtdeAprop, CustoMov, QtdeMov FROM RelFormPrecoAprop WHERE FilialEmpresa = ? AND Produto = ? AND Data BETWEEN ? AND ? AND Natureza IN (?,?) ", dQtdeAprop, dCustoMov, dQtdeMov, iFilialEmpresa, sParam, dtDataDe, dtDataAte, NATUREZA_PROD_MATERIA_PRIMA, NATUREZA_PROD_PRODUTO_INTERMEDIARIO)
                        If lErro <> AD_SQL_SUCESSO Then gError 211595
                        
                        lErro = Comando_BuscarPrimeiro(alComando(5))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211596
                        
                        Do While lErro <> AD_SQL_SEM_DADOS
                        
                            dCusto = dCusto + dQtdeAprop * (dCustoMov / dQtdeMov)
                        
                            lErro = Comando_BuscarProximo(alComando(5))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211597
                        
                        Loop
                    
                        bAlt = True
                        sRegistro = Replace(sRegistro, "@Cust_Prod_MP(" & sParam & ")", Format(dCusto, NUM_FORMATO))
                    
                    End If
                    If InStr(1, sRegistro, "@Compra_Pc") <> 0 Then
                        iPos1 = InStr(1, sRegistro, "@Compra_Pc")
                        iPos2 = InStr(iPos1, sRegistro, "(")
                        iPos3 = InStr(iPos1, sRegistro, ")")
                        sParam = Mid(sRegistro, iPos2 + 1, iPos3 - iPos2 - 1)
                        
                        lErro = Comando_Executar(alComando(0), "SELECT Qtde FROM RelFormPrecoPurAux WHERE Produto = ? ", dValorAux, sParam)
                        If lErro <> AD_SQL_SUCESSO Then gError 211582
                        
                        lErro = Comando_BuscarPrimeiro(alComando(0))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                        
                        If lErro = AD_SQL_SUCESSO Then
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Compra_Pc(" & sParam & ")", Format(dValorAux, NUM_FORMATO))
                        Else
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Compra_Pc(" & sParam & ")", Format(0, NUM_FORMATO))
                        End If
                    End If
                    If InStr(1, sRegistro, "@Compra_Rs") <> 0 Then
                        iPos1 = InStr(1, sRegistro, "@Compra_Rs")
                        iPos2 = InStr(iPos1, sRegistro, "(")
                        iPos3 = InStr(iPos1, sRegistro, ")")
                        sParam = Mid(sRegistro, iPos2 + 1, iPos3 - iPos2 - 1)
                        
                        lErro = Comando_Executar(alComando(0), "SELECT Valor FROM RelFormPrecoPurAux WHERE Produto = ? ", dValorAux, sParam)
                        If lErro <> AD_SQL_SUCESSO Then gError 211582
                        
                        lErro = Comando_BuscarPrimeiro(alComando(0))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                        
                        If lErro = AD_SQL_SUCESSO Then
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Compra_Rs(" & sParam & ")", Format(dValorAux, NUM_FORMATO))
                        Else
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Compra_Rs(" & sParam & ")", Format(0, NUM_FORMATO))
                        End If
                    End If
                    If InStr(1, sRegistro, "@Custo_Medio") <> 0 Then
                        iPos1 = InStr(1, sRegistro, "@Custo_Medio")
                        iPos2 = InStr(iPos1, sRegistro, "(")
                        iPos3 = InStr(iPos1, sRegistro, ")")
                        sParam = Mid(sRegistro, iPos2 + 1, iPos3 - iPos2 - 1)
                        
                        dValorAux = 0
                        'lErro = Comando_Executar(alComando(0), "SELECT CustoMedio FROM RelFormPrecoPurAux WHERE Produto = ? ", dValorAux, sParam)
                        'If lErro <> AD_SQL_SUCESSO Then gError 211582
                        
                        'lErro = Comando_BuscarPrimeiro(alComando(0))
                        'If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                        
                        'lErro = CF("Calcula_Custo", dCusto, iFilialEmpresa, APROPR_CUSTO_MEDIO, Year(dtDataAte), Month(dtDataAte), Year(dtDataAte), Month(dtDataAte), sParam)
                        'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                       
                       'mudado para o custo médio da última compra
                       'LENTO D+
                        'lErro = Comando_Executar(alComando(10), "SELECT CustoTotal/Qtde FROM RelComprasPur WHERE Produto = ? AND FilialEmpresa = ? AND Data <= ? AND Qtde > 0.009 ORDER BY Data DESC ", dValorAux, sParam, iFilialEmpresa, dtDataAte)
                        'If lErro <> AD_SQL_SUCESSO Then gError 211582
                        
                        'lErro = Comando_BuscarPrimeiro(alComando(10))
                        'If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                        
                        lErro = Comando_Executar(alComando(10), "SELECT N.NumIntDoc, N.DataEntrada, I.NumIntDoc, I.Quantidade FROM TiposDocInfo AS T, NFiscal AS N, ItensNFiscal AS I WHERE N.NumIntDoc = I.NumIntNF AND T.Codigo = N.TipoNFiscal AND N.Status <> 7 AND T.Compras = 1 AND I.Produto = ? AND N.FilialEmpresa = ? AND N.DataEmissao <= ? AND T.Codigo NOT IN (104,105,165,166,168,170) AND T.Complementar = 0 ORDER BY N.DataEmissao DESC ", lNumIntNF, dtDataNF, lNumIntItemNF, dQtdeNF, sParam, iFilialEmpresa, dtDataAte)
                        If lErro <> AD_SQL_SUCESSO Then gError 211582
                        
                        lErro = Comando_BuscarPrimeiro(alComando(10))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                        
                        If lErro = AD_SQL_SUCESSO Then
                        
                            lErro = Comando_Executar(alComando(11), "SELECT SUM(Custo)/SUM(Quantidade) FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = 1 AND NumIntDocOrigem = ? ", dValorAux, lNumIntItemNF)
                            If lErro <> AD_SQL_SUCESSO Then gError 211582
                            
                            lErro = Comando_BuscarPrimeiro(alComando(11))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                            
                                                       
                            'Obtem os conhecimentos de transportes associados a NFe se existirem
                            lErro = Comando_Executar(alComando(12), "SELECT I.NumIntDoc, I.Quantidade FROM NFiscal AS N, ItensNFiscal AS I WHERE N.NumIntDoc = I.NumIntNF AND N.NumIntNotaOriginal = ? AND N.Status <> 7 AND N.TipoNFiscal  = 104 AND I.Produto = ?", lNumIntItemNF, dQtdeCTe, lNumIntNF, sParam)
                            If lErro <> AD_SQL_SUCESSO Then gError 211582
                            
                            lErro = Comando_BuscarPrimeiro(alComando(12))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                            
                            If lErro = AD_SQL_SUCESSO Then
                            
                                dCustoCTe = 0
                            
                                Do While lErro <> AD_SQL_SEM_DADOS
                                
                                
                                    lErro = Comando_Executar(alComando(13), "SELECT SUM(Custo) FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = 1 AND NumIntDocOrigem = ? AND Produto = ? ", dCustoCTeAux, lNumIntItemNF, sParam)
                                    If lErro <> AD_SQL_SUCESSO Then gError 211582
                                    
                                    lErro = Comando_BuscarPrimeiro(alComando(13))
                                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                                    
                                    dCustoCTe = dCustoCTe + dCustoCTeAux
                                
                                    lErro = Comando_BuscarProximo(alComando(12))
                                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                                
                                Loop
                                
                                dValorAux = dValorAux + dCustoCTe / dQtdeCTe
                            
                            Else
                                
                                'Procura o CTe pela data
                                lErro = Comando_Executar(alComando(14), "SELECT I.NumIntDoc, I.Quantidade FROM TiposDocInfo AS T, NFiscal AS N, ItensNFiscal AS I WHERE N.NumIntDoc = I.NumIntNF AND T.Codigo = N.TipoNFiscal AND N.Status <> 7 AND I.Produto = ? AND N.FilialEmpresa = ? AND N.DataEntrada = ? AND T.Codigo = 104 ", lNumIntItemNF, dQtdeCTe, sParam, iFilialEmpresa, dtDataNF)
                                If lErro <> AD_SQL_SUCESSO Then gError 211582
                                
                                lErro = Comando_BuscarPrimeiro(alComando(14))
                                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                                
                                If lErro = AD_SQL_SUCESSO Then
                                
                                    lErro = Comando_Executar(alComando(15), "SELECT SUM(Custo) FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = 1 AND NumIntDocOrigem = ? ", dCustoCTeAux, lNumIntItemNF)
                                    If lErro <> AD_SQL_SUCESSO Then gError 211582
                                    
                                    lErro = Comando_BuscarPrimeiro(alComando(15))
                                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211583
                               
                                    dValorAux = dValorAux + dCustoCTeAux / dQtdeCTe
                                
                                End If
                                
                                
                            End If
                        
                        End If
                        
                        'If lErro = AD_SQL_SUCESSO Then
                            bAlt = True
                            sRegistro = Replace(sRegistro, "@Custo_Medio(" & sParam & ")", Format(dValorAux, NUM_FORMATO_CUSTO))
                        'Else
                         '   bAlt = True
                        '    sRegistro = Replace(sRegistro, "@Custo_Medio(" & sParam & ")", Format(0, NUM_FORMATO_CUSTO))
                        'End If
                    End If
                    
                    If bAlt Then
'                        Set objPlanilhaRange = objPlanilhaExcel.cells(iLinha, iColuna)
'                        If bFormataNum Then
'                            objPlanilhaRange.FormulaR1C1 = "=" & Replace(sRegistro, ",", ".")
'                        Else
'                            objPlanilhaRange.FormulaR1C1 = sRegistro
'                        End If
                        lErro = objExcelApp.Preencher_Celula(iLinha, iColuna, sRegistro, IIf(IsNumeric(sRegistro), NUM_FORMATO, ""))
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                    End If
                End If
            Next
        Next
    Next
       
'    objPastaTrabalho.SaveAS sNomeDiretorio & sNomeArquivo & ".xls"
'    GL_objExcelSist.Visible = True

    lErro = objExcelApp.Salvar(sNomeDiretorio & sNomeArquivo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = objExcelApp.Mudar_Visibilidade(True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'    'abrir transacao
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 211627
    
    lErro = Comando_Executar(alComando(6), "DELETE FROM RelVendEstProd WHERE NumIntRel = ? ", lNumIntRel)
    If lErro <> AD_SQL_SUCESSO Then gError 211598

    lErro = Comando_Executar(alComando(9), "DELETE FROM RelFormPrecoPurAux WHERE NumIntRel = ? ", lNumIntRel)
    If lErro <> AD_SQL_SUCESSO Then gError 211598

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
     'fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211628

    RelFormPrecoPur_Prepara = SUCESSO

    Exit Function

Erro_RelFormPrecoPur_Prepara:

    RelFormPrecoPur_Prepara = gErr

    Select Case gErr

        Case 211581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211582 To 211589
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELVENDESTPROD", gErr)

        Case 211590, 211591
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACTA", gErr)

        Case 211592 To 211597
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELFORMPRECOAPROP", gErr)
            
        Case 211598
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_RELVENDESTPROD", gErr)

        Case 211627
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 211628
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211599)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    'If bArqAberto Then Call CF("Excel_Fechar")
    If bArqAberto Then Call objExcelApp.Fechar

    Exit Function

End Function

Public Function RelComparativoAno_Prepara(ByVal iFilialEmpresa As Integer, ByVal iAno1 As Integer, ByVal iAno2 As Integer, ByVal lCliente As Long, ByVal iTipoVend As Integer, ByVal iVendedor As Integer, ByVal sProdDe As String, ByVal sProdAte As String, ByVal sCat As String, ByVal sCatDe As String, ByVal sCatAte As String, ByVal iLayout As Integer, lNumIntRel As Long) As Long
'        lErro = CF("RelComparativoAno_Prepara", iFilialEmpresa, iAno1, iAno2, lCliente, iTipoVend, iVendedor, sProdDe, sProdAte, sCat, sCatDe, sCatAte, iLayout, lNumIntRel)

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 4) As Long, sSQL As String
Dim colVendedores As New Collection
Dim lNumIntRelAux As Long, objVendedor As ClassVendedor
Dim sProduto As String, sDescProd As String, iAno As Integer, dMes(1 To 12) As Double
Dim sProdAnt As String, sGrupo As String, iQuebra As Integer, iSeq As Integer
Dim sProdMask As String

On Error GoTo Erro_RelComparativoAno_Prepara

    'abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211802
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211803
    Next
        
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUMINTREL_COMPANO", lNumIntRel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    If iVendedor <> 0 Then
        'Obter NumIntRel
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUMINTREL_FATPRODVEND", lNumIntRelAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        If iTipoVend = VENDEDOR_DIRETO Then
            Set objVendedor = New ClassVendedor
            objVendedor.iCodigo = iVendedor
            colVendedores.Add objVendedor
        Else
            lErro = CF("Vendedor_Le_Hierarquia", iVendedor, colVendedores, 0)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
    End If

    For Each objVendedor In colVendedores
        lErro = Comando_Executar(alComando(1), "INSERT INTO RelFatProdVend (NumIntRel, Vendedor) VALUES (?,?)", lNumIntRelAux, objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 211804
    Next
    
    If iLayout = 1 Or iLayout = 2 Then
        
        'Monta o Select
        lErro = RelComparativoAnoProdSQL_Prepara(iLayout, iFilialEmpresa, iAno1, iAno2, lCliente, lNumIntRelAux, iVendedor, sProdDe, sProdAte, sCat, sCatDe, sCatAte, sSQL)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        sProduto = String(STRING_PRODUTO, 0)
        sDescProd = String(STRING_PRODUTO_DESCRICAO, 0)
    
        'Passa os Binds
        lErro = RelComparativoAnoProdInt_Prepara(alComando(2), sProduto, sDescProd, iAno, dMes(1), dMes(2), dMes(3), dMes(4), dMes(5), dMes(6), dMes(7), dMes(8), dMes(9), dMes(10), dMes(11), dMes(12), iFilialEmpresa, iAno1, iAno2, lCliente, lNumIntRelAux, iVendedor, sProdDe, sProdAte, sCat, sCatDe, sCatAte, sSQL)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211806
        
        iSeq = 0
        Do While lErro = AD_SQL_SUCESSO
        
            iSeq = iSeq + 1
            If sProduto <> sProdAnt Then
                iQuebra = iQuebra + 1
                sProdAnt = sProduto
                
                lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdMask)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                sGrupo = sProdMask & SEPARADOR & sDescProd
                
            End If
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO RelComparativoAno (NumIntRel,QuebraGrupo,Ano,Seq,NomeGrupo,Mes01,Mes02,Mes03,Mes04,Mes05,Mes06,Mes07,Mes08,Mes09,Mes10,Mes11,Mes12) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                                lNumIntRel, iQuebra, iAno, iSeq, sGrupo, dMes(1), dMes(2), dMes(3), dMes(4), dMes(5), dMes(6), dMes(7), dMes(8), dMes(9), dMes(10), dMes(11), dMes(12))
            If lErro <> AD_SQL_SUCESSO Then gError 211805
    
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211806
         
        Loop
        
    End If

    lErro = Comando_Executar(alComando(4), "DELETE FROM RelFatProdVend WHERE NumIntRel = ? ", lNumIntRelAux)
    If lErro <> AD_SQL_SUCESSO Then gError 211804

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
     'fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211807

    RelComparativoAno_Prepara = SUCESSO

    Exit Function

Erro_RelComparativoAno_Prepara:

    RelComparativoAno_Prepara = gErr

    Select Case gErr

        Case 211802
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 211803
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211804
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELFATPRODVEND", gErr)

        Case 211805
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_COMPARATIVOANO", gErr)

        Case 211806
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFATFILCLI", gErr)
    
        Case 211807
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211808)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function RelComparativoAnoProdSQL_Prepara(ByVal iLayout As Integer, ByVal iFilialEmpresa As Integer, ByVal iAno1 As Integer, ByVal iAno2 As Integer, ByVal lCliente As Long, ByVal lNumIntRel As Long, ByVal iVendedor As Integer, ByVal sProdDe As String, ByVal sProdAte As String, ByVal sCat As String, ByVal sCatDe As String, ByVal sCatAte As String, sSQL As String) As Long
'lErro = RelComparativoAnoProdSQL_Prepara(iFilialEmpresa, iAno1, iAno2, lCliente, lNumIntRel, iVendedor, sProdDe, sProdAte, sCat, sCatDe, sCatAte, sSQL)
Dim lErro As Long
Dim sSubWhere As String
Dim sTabela As String
Dim iUsuLigRegVenda As Integer

On Error GoTo Erro_RelComparativoAnoProdSQL_Prepara

    lErro = CF("RegiaoVenda_Le_UsuarioCobrador", gsUsuario, iUsuLigRegVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If lCliente <> 0 Or iVendedor <> 0 Or iUsuLigRegVenda <> 0 Then
        sTabela = "SldMesFatFilCli"
    Else
        sTabela = "SldMesFat"
    End If
    
    If iLayout = 1 Then
        If lCliente <> 0 Or iVendedor <> 0 Or iUsuLigRegVenda <> 0 Then
            sSQL = "SELECT S.Ano, S.Produto, P.Descricao, SUM(ValorFaturado1 - ValorDevolvido1) AS Mes01,SUM(ValorFaturado2 - ValorDevolvido2) AS Mes02, SUM(ValorFaturado3 - ValorDevolvido3) AS Mes03, SUM(ValorFaturado4 - ValorDevolvido4) AS Mes04, SUM(ValorFaturado5 - ValorDevolvido5) AS Mes05, SUM(ValorFaturado6 - ValorDevolvido6) AS Mes06, SUM(ValorFaturado7 - ValorDevolvido7) AS Mes07, SUM(ValorFaturado8 - ValorDevolvido8) AS Mes08, SUM(ValorFaturado9 - ValorDevolvido9) AS Mes09, SUM(ValorFaturado10 - ValorDevolvido10) AS Mes010, SUM(ValorFaturado11 - ValorDevolvido11) AS Mes011, SUM(ValorFaturado12 - ValorDevolvido12) AS Mes012  FROM SldMesFatFilCli AS S, Produtos AS P, RegioesVendas As R, FiliaisClientes AS F WHERE P.Codigo = S.Produto AND F.Regiao = R.Codigo AND S.Cliente = F.CodCliente AND S.Filial = F.CodFilial AND S.Ano IN (?,?) "
        Else
            sSQL = "SELECT S.Ano, S.Produto, P.Descricao, SUM(ValorFaturado1 - ValorDevolvido1) AS Mes01,SUM(ValorFaturado2 - ValorDevolvido2) AS Mes02, SUM(ValorFaturado3 - ValorDevolvido3) AS Mes03, SUM(ValorFaturado4 - ValorDevolvido4) AS Mes04, SUM(ValorFaturado5 - ValorDevolvido5) AS Mes05, SUM(ValorFaturado6 - ValorDevolvido6) AS Mes06, SUM(ValorFaturado7 - ValorDevolvido7) AS Mes07, SUM(ValorFaturado8 - ValorDevolvido8) AS Mes08, SUM(ValorFaturado9 - ValorDevolvido9) AS Mes09, SUM(ValorFaturado10 - ValorDevolvido10) AS Mes010, SUM(ValorFaturado11 - ValorDevolvido11) AS Mes011, SUM(ValorFaturado12 - ValorDevolvido12) AS Mes012  FROM SldMesFat AS S, Produtos AS P WHERE P.Codigo = S.Produto AND S.Ano IN (?,?) "
        End If
    Else
        If lCliente <> 0 Or iVendedor <> 0 Or iUsuLigRegVenda <> 0 Then
            sSQL = "SELECT S.Ano, S.Produto, P.Descricao, SUM(QtdFaturada1 - QtdDevolvida1) AS Mes01,SUM(QtdFaturada2 - QtdDevolvida2) AS Mes02, SUM(QtdFaturada3 - QtdDevolvida3) AS Mes03, SUM(QtdFaturada4 - QtdDevolvida4) AS Mes04, SUM(QtdFaturada5 - QtdDevolvida5) AS Mes05, SUM(QtdFaturada6 - QtdDevolvida6) AS Mes06, SUM(QtdFaturada7 - QtdDevolvida7) AS Mes07, SUM(QtdFaturada8 - QtdDevolvida8) AS Mes08, SUM(QtdFaturada9 - QtdDevolvida9) AS Mes09, SUM(QtdFaturada10 - QtdDevolvida10) AS Mes010, SUM(QtdFaturada11 - QtdDevolvida11) AS Mes011, SUM(QtdFaturada12 - QtdDevolvida12) AS Mes012  FROM SldMesFatFilCli AS S, Produtos AS P, RegioesVendas As R, FiliaisClientes AS F WHERE P.Codigo = S.Produto AND F.Regiao = R.Codigo AND S.Cliente = F.CodCliente AND S.Filial = F.CodFilial AND S.Ano IN (?,?) "
        Else
            sSQL = "SELECT S.Ano, S.Produto, P.Descricao, SUM(QuantFaturada1 - QuantDevolvida1) AS Mes01,SUM(QuantFaturada2 - QuantDevolvida2) AS Mes02, SUM(QuantFaturada3 - QuantDevolvida3) AS Mes03, SUM(QuantFaturada4 - QuantDevolvida4) AS Mes04, SUM(QuantFaturada5 - QuantDevolvida5) AS Mes05, SUM(QuantFaturada6 - QuantDevolvida6) AS Mes06, SUM(QuantFaturada7 - QuantDevolvida7) AS Mes07, SUM(QuantFaturada8 - QuantDevolvida8) AS Mes08, SUM(QuantFaturada9 - QuantDevolvida9) AS Mes09, SUM(QuantFaturada10 - QuantDevolvida10) AS Mes010, SUM(QuantFaturada11 - QuantDevolvida11) AS Mes011, SUM(QuantFaturada12 - QuantDevolvida12) AS Mes012  FROM SldMesFat AS S, Produtos AS P WHERE P.Codigo = S.Produto AND S.Ano IN (?,?) "
        End If
    End If
    
    If iFilialEmpresa <> 0 Then
        sSQL = sSQL & "AND S.FilialEmpresa = ? "
    Else
        sSQL = sSQL & "AND S.FilialEmpresa < 50 "
    End If

    If lCliente <> 0 Then sSQL = sSQL & "AND S.Cliente = ? "

    If iVendedor <> 0 Then sSQL = sSQL & "AND S.Cliente IN (SELECT F.CodCliente FROM RelFatProdVend AS X, FiliaisClientes AS F WHERE X.Vendedor = F.Vendedor AND X.NumIntRel = ?) "

    If sProdDe <> "" Then sSQL = sSQL & "AND S.Produto >= ? "

    If sProdAte <> "" Then sSQL = sSQL & "AND S.Produto <= ? "
    
    If iUsuLigRegVenda <> 0 Then
        sSQL = sSQL & "AND R.UsuarioCobrador = '" & gsUsuario & "' "
    End If

    sSubWhere = ""
    If Len(Trim(sCat)) > 0 Then
        If Len(Trim(sCatDe)) > 0 Then
            sSubWhere = sSubWhere & "AND X.Item >= ? "
        End If
        If Len(Trim(sCatAte)) > 0 Then
            sSubWhere = sSubWhere & "AND X.Item <= ? "
        End If
        sSQL = sSQL & " AND EXISTS ( SELECT X.Produto FROM ProdutoCategoria AS X WHERE S.Produto= X.Produto AND X.Categoria = ? " & sSubWhere & ") "
    End If
    
    sSQL = sSQL & " GROUP BY S.Ano, S.Produto, P.Descricao ORDER BY Produto, Ano"
    
    RelComparativoAnoProdSQL_Prepara = SUCESSO
    
    Exit Function

Erro_RelComparativoAnoProdSQL_Prepara:

    RelComparativoAnoProdSQL_Prepara = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211809)

    End Select

    Exit Function

End Function

Private Function RelComparativoAnoProdInt_Prepara(ByVal lComando As Long, vsProduto As Variant, vsDescProd As Variant, viAno As Variant, vdMes1 As Variant, vdMes2 As Variant, vdMes3 As Variant, vdMes4 As Variant, vdMes5 As Variant, vdMes6 As Variant, vdMes7 As Variant, vdMes8 As Variant, vdMes9 As Variant, vdMes10 As Variant, vdMes11 As Variant, vdMes12 As Variant, ByVal viFilialEmpresa As Variant, ByVal viAno1 As Variant, ByVal viAno2 As Variant, ByVal vlCliente As Variant, ByVal vlNumIntRelAux As Variant, ByVal viVendedor As Variant, ByVal vsProdDe As Variant, ByVal vsProdAte As Variant, ByVal vsCat As Variant, ByVal vsCatDe As Variant, ByVal vsCatAte As Variant, ByVal sSQL As String) As Long
'        lErro = RelComparativoAnoProdInt_Prepara(alComando(2), sProduto, sDescProd, iAno, dMes1, dMes2, dMes3, dMes4, dMes5, dMes6, dMes7, dMes8, dMes9, dMes10, dMes11, dMes12, iFilialEmpresa, iAno1, iAno2, lCliente, lNumIntRelAux, iVendedor, sProdDe, sProdAte, sCat, sCatDe, sCatAte, sSQL)
Dim lErro As Long

On Error GoTo Erro_RelComparativoAnoProdInt_Prepara

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211810

    lErro = Comando_BindVarInt(lComando, viAno)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211811

    lErro = Comando_BindVarInt(lComando, vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211812
    
    lErro = Comando_BindVarInt(lComando, vsDescProd)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211813
    
    lErro = Comando_BindVarInt(lComando, vdMes1)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211814
    
    lErro = Comando_BindVarInt(lComando, vdMes2)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211815
    
    lErro = Comando_BindVarInt(lComando, vdMes3)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211816
    
    lErro = Comando_BindVarInt(lComando, vdMes4)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211817
    
    lErro = Comando_BindVarInt(lComando, vdMes5)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211818
    
    lErro = Comando_BindVarInt(lComando, vdMes6)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211819
    
    lErro = Comando_BindVarInt(lComando, vdMes7)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211820
    
    lErro = Comando_BindVarInt(lComando, vdMes8)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211821
    
    lErro = Comando_BindVarInt(lComando, vdMes9)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211822
    
    lErro = Comando_BindVarInt(lComando, vdMes10)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211823
    
    lErro = Comando_BindVarInt(lComando, vdMes11)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211824
    
    lErro = Comando_BindVarInt(lComando, vdMes12)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211825
    
    lErro = Comando_BindVarInt(lComando, viAno1)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211826
    
    lErro = Comando_BindVarInt(lComando, viAno2)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211827
    
    If viFilialEmpresa <> 0 Then
        lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211828
    End If
    If vlCliente <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlCliente)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211829
    End If
    If viVendedor <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlNumIntRelAux)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211830
    End If
    If vsProdDe <> "" Then
        lErro = Comando_BindVarInt(lComando, vsProdDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211831
    End If
    If vsProdAte <> "" Then
        lErro = Comando_BindVarInt(lComando, vsProdAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211832
    End If
    If vsCat <> "" Then
        lErro = Comando_BindVarInt(lComando, vsCat)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211833
    End If
    If vsCatDe <> "" Then
        lErro = Comando_BindVarInt(lComando, vsCatDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211834
    End If
    If vsCatAte <> "" Then
        lErro = Comando_BindVarInt(lComando, vsCatAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 211835
    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 211836

    RelComparativoAnoProdInt_Prepara = SUCESSO

    Exit Function

Erro_RelComparativoAnoProdInt_Prepara:

    RelComparativoAnoProdInt_Prepara = gErr

    Select Case gErr

        Case 211810 To 211836
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211837)

    End Select

    Exit Function

End Function

Public Function RelComissPur_Prepara(ByVal colReg As Collection, ByVal iFilialEmpresa As Integer, ByVal dtEmiDe As Date, ByVal dtEmiAte As Date, ByVal dtBxDe As Date, ByVal dtBxAte As Date, ByVal lCliDe As Long, ByVal lCliAte As Long, ByVal iTipoVend As Integer, ByVal iVendedorFiltro As Integer, ByVal iBx As Integer, ByVal iDet As Integer, lNumIntRel As Long) As Long

Dim lErro As Long
Dim iIndice As Integer, iSeq As Integer
Dim lTransacao As Long, sSQL As String
Dim alComando(0 To 4) As Long
Dim objVendedor As ClassVendedor
Dim objVendPur As ClassVendedorPur
Dim objVendPurExc As ClassVendedorPurExc
Dim colVendFiltro As New Collection
Dim colVendedores As New Collection

Dim lNumIntNF As Long, lNumIntItemNF As Long, sSerie As String, lNumNF As Long, iItemNF As Integer
Dim dtDataEmissao As Date, dValorNF As Double, dValorProdutos As Double, dValorItem As Double
Dim iFaturamento As Integer, sProduto As String, sDescricaoItem As String, lCliente As Long
Dim iFilial As Integer, iVendedor As Integer, iRegiao As Integer, sCidade As String
Dim sNomeCliente As String, iClasseDocCPR As Integer, lNumIntDocCPR As Long, dPercComissP As Double, dMetaPadrao As Double
Dim dQtdeProd As Double, iClasseUM As Integer, sUMVenda As String, sUMNF As String, lNumIntBx As Long

Dim lNumIntBaixa As Long, lNumIntParcela As Long, iNumParcela As Integer, dValorTitulo As Double
Dim dValorIRRF As Double, dValorRecebido As Double
Dim dValorBaixado As Double

Dim dtDataDeAux As Date, dtDataAteAux As Date, dValorParcRecRateadoProd As Double
Dim iVendedorAnt As Integer, sProdutoAnt As String, sDescProdAnt As String
Dim dMetaPadraoAnt As Double, dPercComissPAnt As Double, sUMAnt As String, dFator As Double
Dim dMetaProdVend As Double, dPercComissProdVend As Double
Dim dFatorNF As Double, dFatorRec As Double, bTerminou As Boolean
Dim dAcumuladoProd As Double, dAcumuladoVend As Double, sNomeVend As String, dAcumuladoProdQtde As Double
Dim dValor As Double, lNumIntDoc As Long, dtData As Date, bAchou As Boolean
Dim dVlrComissao As Double, dAcumuladoItemNF As Double, dVlrMedio As Double, dVlrMeta As Double

On Error GoTo Erro_RelComissPur_Prepara

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 215012

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 215013
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUMINTREL_RELCOMISSPUR", lNumIntRel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê todos os vendedores com o seus respectivos nomes
    lErro = CF("Vendedor_Le_Todos", colVendedores)
    If lErro <> SUCESSO And lErro <> 104287 Then gError ERRO_SEM_MENSAGEM
    
    For Each objVendedor In colVendedores
        
        objVendedor.dSaldoComissao = 0
    
        lErro = CF("Vendedor_Le_Customizado", objVendedor)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Next

    'Obtem a lista de vendedores que serão filtrados
    If iVendedorFiltro <> 0 Then
        If iTipoVend = VENDEDOR_DIRETO Then
            Set objVendedor = New ClassVendedor
            objVendedor.iCodigo = iVendedorFiltro
            colVendFiltro.Add objVendedor
        Else
            lErro = CF("Vendedor_Le_Hierarquia", iVendedorFiltro, colVendFiltro, 0)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
    End If
    
    iSeq = 0
    bTerminou = False
    
    If dtBxDe = DATA_NULA Then
        dtDataDeAux = #1/1/1900#
    Else
        dtDataDeAux = dtBxDe
    End If
    
    If dtBxAte = DATA_NULA Then
        dtDataAteAux = DateAdd("yyyy", 10, gdtDataHoje)
    Else
        dtDataAteAux = dtBxAte
    End If
    
    sSerie = String(STRING_MAXIMO, 0)
    sProduto = String(STRING_MAXIMO, 0)
    sDescricaoItem = String(STRING_MAXIMO, 0)
    sCidade = String(STRING_MAXIMO, 0)
    sNomeCliente = String(STRING_MAXIMO, 0)
    sUMVenda = String(STRING_MAXIMO, 0)
    sUMNF = String(STRING_MAXIMO, 0)
    
    lErro = RelComissPur_PreparaSQL(iFilialEmpresa, dtEmiDe, dtEmiAte, dtBxDe, dtBxAte, lCliDe, lCliAte, iBx, MARCADO, colVendFiltro, colReg, sSQL)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = RelComissPur_PreparaInt(alComando(0), sSQL, iFilialEmpresa, dtEmiDe, dtEmiAte, dtBxDe, dtBxAte, lCliDe, lCliAte, iBx, lNumIntNF, lNumIntItemNF, sSerie, lNumNF, iItemNF, dtDataEmissao, dValorNF, dValorProdutos, dValorItem, iFaturamento, sProduto, sDescricaoItem, lCliente, iFilial, iVendedor, iRegiao, sCidade, sNomeCliente, iClasseDocCPR, lNumIntDocCPR, dPercComissP, dMetaPadrao, dQtdeProd, iClasseUM, sUMVenda, sUMNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 215014
    
    If lErro = AD_SQL_SEM_DADOS Then bTerminou = True

    dAcumuladoProd = 0
    dAcumuladoProdQtde = 0
    dAcumuladoVend = 0
    dAcumuladoItemNF = 0
    
    sProdutoAnt = sProduto
    sDescProdAnt = sDescricaoItem
    iVendedorAnt = iVendedor
    sUMAnt = sUMVenda
    dMetaPadraoAnt = dMetaPadrao
    dPercComissPAnt = dPercComissP
    
    Do While Not bTerminou
       
        If iBx = MARCADO Then
            lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, NumParcela, ValorTitulo, ValorIRRF, ValorBaixado, Data, NumIntBaixa FROM BaixasClasseDocCR WHERE ClasseDocCPR = ? AND NumIntDoc = ? AND Data BETWEEN ? AND ? ORDER BY Data", _
            lNumIntDoc, iNumParcela, dValorTitulo, dValorIRRF, dValor, dtData, lNumIntBx, iClasseDocCPR, lNumIntDocCPR, dtDataDeAux, dtDataAteAux)
        Else
            lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, NumParcela, ValorTitulo, ValorIRRF, SaldoParcela, DataVencimento, 0 FROM ReceberClasseDocCR WHERE ClasseDocCPR = ? AND NumIntDoc = ? AND DataVencimento BETWEEN ? AND ? ORDER BY DataVencimento", _
            lNumIntDoc, iNumParcela, dValorTitulo, dValorIRRF, dValor, dtData, lNumIntBx, iClasseDocCPR, lNumIntDocCPR, dtDataDeAux, dtDataAteAux)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 215015
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 215016
    
        Do While lErro <> AD_SQL_SEM_DADOS
   
            iSeq = iSeq + 1

            'CÁLCULO BASEADO NO VALOR DOS ITENS
            dFatorRec = 0
            If (dValorTitulo - dValorIRRF) > DELTA_VALORMONETARIO Then
                dFatorRec = dValor / (dValorTitulo - dValorIRRF)
            End If
            
            If iFaturamento = 1 Then
                dValorParcRecRateadoProd = Arredonda_Moeda(dValorItem * dFatorRec)
            Else
                dValorParcRecRateadoProd = -Arredonda_Moeda(dValorItem * dFatorRec)
            End If
            
            lErro = CF("UM_Conversao", iClasseUM, sUMNF, sUMVenda, dFator)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
            lErro = Comando_Executar(alComando(2), "INSERT INTO RelComissPurDet (NumIntRel,Produto,Vendedor,Seq,SerieNF,NumNF,DataEmissaoNF,ValorTotalNF,ItemNF,ValorItemNF,NumParcela,DataBaixa,ValorTotalParc,ValorParcBx,ValorParcBxRateadoProd,Cliente,NomeCliente, Cidade, NumIntItemNF,NumIntParcRec, NumIntBaixa, QdteItemNF) " & _
                                                   "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntRel, sProduto, iVendedor, iSeq, sSerie, lNumNF, dtDataEmissao, dValorNF, iItemNF, dValorItem, iNumParcela, dtData, (dValorTitulo - dValorIRRF), dValor, dValorParcRecRateadoProd, lCliente, sNomeCliente, sCidade, lNumIntItemNF, lNumIntParcela, lNumIntBx, dQtdeProd * dFator)
            If lErro <> AD_SQL_SUCESSO Then gError 215017
            
            dAcumuladoProd = dAcumuladoProd + dValorParcRecRateadoProd
            dAcumuladoProdQtde = dAcumuladoProdQtde + dQtdeProd * dFator
            dAcumuladoVend = dAcumuladoVend + dValorParcRecRateadoProd
            dAcumuladoItemNF = dAcumuladoItemNF + dValorItem
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 215018
            
        Loop
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 215019
        
        If lErro = AD_SQL_SEM_DADOS Then bTerminou = True
        
        'Se acabou ou trocou o produto ou trocou o vendedor
        If bTerminou Or sProduto <> sProdutoAnt Or iVendedor <> iVendedorAnt Then
                
            'Padrão do Produto
            dMetaProdVend = dMetaPadraoAnt
            dPercComissProdVend = dPercComissPAnt
                
            'Procura pela exceção no vendedor
            'bAchou = False
            For Each objVendedor In colVendedores
                If objVendedor.iCodigo = iVendedorAnt Then
                    Set objVendPur = objVendedor.objInfoUsu
                    bAchou = True
                    Exit For
                End If
            Next
            'If Not bAchou Then gError 999999
            
            For Each objVendPurExc In objVendPur.colExcecoes
                If objVendPurExc.sProduto = sProdutoAnt Then
                    '-1=Vale o padrão
                    If objVendPurExc.dMetaComissao <> -1 Then dMetaProdVend = objVendPurExc.dMetaComissao
                    If objVendPurExc.dPercComissao <> -1 Then dPercComissProdVend = objVendPurExc.dPercComissao
                    Exit For
                End If
            Next
            If iBx = DESMARCADO Then dMetaProdVend = 0  'Não usar a meta de a comissão não for pela baixa
            
            dVlrMedio = 0
            If dAcumuladoProdQtde > 0 Then
                dVlrMedio = dAcumuladoItemNF / dAcumuladoProdQtde
            End If
            
'            If dAcumuladoProdQtde >= dMetaProdVend And dAcumuladoProdQtde > 0 Then
'                dVlrComissao = (dAcumuladoProd / dAcumuladoProdQtde) * dPercComissProdVend * (dAcumuladoProdQtde - dMetaProdVend) 'ganha pelo que excedeu
'            Else
'                dVlrComissao = 0
'            End If

            dVlrComissao = (dAcumuladoProd - (dMetaProdVend * dVlrMedio)) * dPercComissProdVend
            If dVlrComissao < 0 Then dVlrComissao = 0
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO RelComissPurProd (NumIntRel,Vendedor,Produto,Descricao, ValorVenda, QtdeVenda, UM, PercComissao, MetaComissao, ValorComissao, VlrMedioVend ) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntRel, iVendedorAnt, sProdutoAnt, sDescProdAnt, dAcumuladoProd, dAcumuladoProdQtde, sUMAnt, dPercComissProdVend, dMetaProdVend, dVlrComissao, dVlrMedio)
            If lErro <> AD_SQL_SUCESSO Then gError 215020
                        
            objVendedor.dSaldoComissao = objVendedor.dSaldoComissao + dVlrComissao
            
            dAcumuladoProd = 0
            dAcumuladoProdQtde = 0
            dAcumuladoItemNF = 0

            sProdutoAnt = sProduto
            sDescProdAnt = sDescricaoItem
            
            sUMAnt = sUMVenda
            dMetaPadraoAnt = dMetaPadrao
            dPercComissPAnt = dPercComissP
        
        End If
        
        'Se acabou ou trocou o vendedor (quando exibe quebrado por vendedor)
        If bTerminou Or iVendedor <> iVendedorAnt Then
        
            sNomeVend = ""
            If iVendedorAnt <> 0 Then
                For Each objVendedor In colVendedores
                    If objVendedor.iCodigo = iVendedorAnt Then
                        sNomeVend = objVendedor.sNomeReduzido
                        Exit For
                    End If
                Next
            End If
        
            lErro = Comando_Executar(alComando(4), "INSERT INTO RelComissPurVend (NumIntRel,Vendedor,NomeVendedor,ValorVenda, ValorComissao) VALUES (?,?,?,?,?)", _
                lNumIntRel, iVendedorAnt, sNomeVend, dAcumuladoVend, objVendedor.dSaldoComissao)
            If lErro <> AD_SQL_SUCESSO Then gError 215021
            
            dAcumuladoVend = 0
            
            iVendedorAnt = iVendedor
        
        End If
    
    Loop
        
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 215022

    RelComissPur_Prepara = SUCESSO
    
    Exit Function
    
Erro_RelComissPur_Prepara:

    RelComissPur_Prepara = gErr

    Select Case gErr

        Case 215012
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 215013
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 215014 To 215021
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case 215022
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 215023)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Private Function RelComissPur_PreparaSQL(ByVal iFilialEmpresa As Integer, ByVal dtEmiDe As Date, ByVal dtEmiAte As Date, ByVal dtBxDe As Date, ByVal dtBxAte As Date, ByVal lCliDe As Integer, ByVal lCliAte As Integer, ByVal iBx As Integer, ByVal iDevolucoes As Integer, ByVal colVendedores As Collection, ByVal colRegioes As Collection, sSQL As String) As Long

Dim lErro As Long
Dim sWhere As String
Dim sOrderBy As String
Dim sSubWhere As String
Dim objVendedor As ClassVendedor
Dim vValor As Variant

On Error GoTo Erro_RelComissPur_PreparaSQL
  
    sSQL = "SELECT N.NumIntDoc AS NumIntNF, I.NumIntDoc AS NumIntItemNF, N.Serie, N.NumNotaFiscal, I.Item, N.DataEmissao, N.ValorTotal, N.ValorProdutos, "
    sSQL = sSQL & "ROUND(I.Quantidade * I.PrecoUnitario - I.ValorDesconto,2), T.Faturamento , I.Produto, P.Descricao, "
    sSQL = sSQL & "F.CodCliente, F.CodFilial, F.Vendedor, F.Regiao, E.Cidade, C.NomeReduzido, "
    sSQL = sSQL & "N.ClasseDocCPR , N.NumIntDocCPR, P.PercComissao, P.MetaComissao,I.Quantidade, P.ClasseUM, P.SiglaUMVenda, I.UnidadeMed "
    sSQL = sSQL & "FROM TiposDocInfo AS T, ItensNFiscal As I, NFiscal As N, Produtos As P, Clientes AS C, FiliaisClientes AS F, Enderecos AS E "

    sWhere = "WHERE N.NumIntDoc = I.NumIntNF AND T.Codigo = N.TipoNFiscal AND I.Produto = P.Codigo AND N.Status <> 7 AND "
    sWhere = sWhere & "C.Codigo = N.Cliente And F.CodCliente = N.Cliente And F.CodFilial = N.FilialCli And F.Endereco = E.Codigo "
    
    If iDevolucoes = MARCADO Then
        sWhere = sWhere & "AND T.Faturamento IN (1,2) "
    Else
        sWhere = sWhere & "AND T.Faturamento = 1 "
    End If
        
    If iFilialEmpresa <> 0 Then
        sWhere = sWhere & "AND N.FilialEmpresa = ? "
    End If
    
    If dtEmiDe <> DATA_NULA Then
        sWhere = sWhere & "AND N.DataEmissao >= ? "
    End If
    
    If dtEmiAte <> DATA_NULA Then
        sWhere = sWhere & "AND N.DataEmissao <= ? "
    End If
    
    If lCliDe <> 0 Then
        sWhere = sWhere & "AND N.Cliente >= ? "
    End If
    
    If lCliAte <> 0 Then
        sWhere = sWhere & "AND N.Cliente <= ? "
    End If
       
    If colVendedores.Count <> 0 Then
        sSubWhere = ""
        For Each objVendedor In colVendedores
            If Len(Trim(sSubWhere)) > 0 Then sSubWhere = sSubWhere & ","
            sSubWhere = sSubWhere & CStr(objVendedor.iCodigo)
        Next
        sWhere = sWhere & "AND F.Vendedor IN (" & sSubWhere & ") "
    End If
    
   If colRegioes.Count <> 0 Then
        sSubWhere = ""
        For Each vValor In colRegioes
            If Len(Trim(sSubWhere)) > 0 Then sSubWhere = sSubWhere & ","
            sSubWhere = sSubWhere & CStr(vValor)
        Next
        sWhere = sWhere & "AND F.Regiao IN (" & sSubWhere & ") "
    End If
    
    sSubWhere = ""
    If iBx = DESMARCADO Then
    
        If dtBxDe <> DATA_NULA Then
            sSubWhere = sSubWhere & "AND X.DataVencimento >= ? "
        End If
        
        If dtBxAte <> DATA_NULA Then
            sSubWhere = sSubWhere & "AND X.DataVencimento <= ? "
        End If
    
        sWhere = sWhere & "AND EXISTS (SELECT * FROM ReceberClasseDocCR AS X WHERE X.ClasseDocCPR = N.ClasseDocCPR AND X.NumIntDoc=N.NumIntDocCPR AND Round(X.SaldoParcela, 2) > 0 " & sSubWhere & ") "
    Else
    
        If dtBxDe <> DATA_NULA Then
            sSubWhere = sSubWhere & "AND X.Data >= ? "
        End If
        
        If dtBxAte <> DATA_NULA Then
            sSubWhere = sSubWhere & "AND X.Data <= ? "
        End If
        
        sWhere = sWhere & "AND EXISTS (SELECT * FROM BaixasClasseDocCR AS X WHERE X.ClasseDocCPR = N.ClasseDocCPR AND X.NumIntDoc=N.NumIntDocCPR " & sSubWhere & ") "
    End If
    
    sOrderBy = "ORDER BY F.Vendedor, I.Produto "
    
    sSQL = sSQL & sWhere & sOrderBy
    
    RelComissPur_PreparaSQL = SUCESSO
     
    Exit Function
    
Erro_RelComissPur_PreparaSQL:

    RelComissPur_PreparaSQL = gErr
     
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 215024)
     
    End Select
     
    Exit Function

End Function

Private Function RelComissPur_PreparaInt(ByVal lComando As Long, ByVal sSQL As String, ByVal viFilialEmpresa As Variant, ByVal vdtDataEmiDe As Variant, ByVal vdtDataEmiAte As Variant, ByVal vdtDataBxDe As Variant, ByVal vdtDataBxAte As Variant, ByVal vlClienteDe As Variant, ByVal vlClienteAte As Variant, ByVal viBx As Variant, vlNumIntNF As Variant, vlNumIntItemNF As Variant, vsSerie As Variant, vlNumNotaFiscal As Variant, viItemNF As Variant, vdtDataEmissao As Variant, vdValorTotal As Variant, vdValorProdutos As Variant, vdValorItem As Variant, viFaturamento As Variant, vsProduto As Variant, vsDescricaoItem As Variant, vlCliente As Variant, viFilial As Variant, viVendedor As Variant, viRegiao As Variant, vsCidadeRet As Variant, vsNomeCliente As Variant, viClasseDocCPR As Variant, vlNumIntDocCPR As Variant, vdPercComissP As Variant, vdMetaPadrao As Variant, vdQtdeProd As Variant, viClasseUM As Variant, vsUMVenda As Variant, vsUMNF As Variant) As Long

Dim lErro As Long
Dim viStatus As Variant

On Error GoTo Erro_RelComissPur_PreparaInt

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM

    lErro = Comando_BindVarInt(lComando, vlNumIntNF)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM

    lErro = Comando_BindVarInt(lComando, vlNumIntItemNF)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsSerie)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vlNumNotaFiscal)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viItemNF)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdtDataEmissao)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdValorTotal)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdValorProdutos)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdValorItem)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viFaturamento)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsDescricaoItem)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vlCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viFilial)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viVendedor)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viRegiao)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsCidadeRet)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsNomeCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viClasseDocCPR)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vlNumIntDocCPR)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdPercComissP)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdMetaPadrao)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdQtdeProd)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viClasseUM)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsUMVenda)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
     
    lErro = Comando_BindVarInt(lComando, vsUMNF)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
     
    If viFilialEmpresa <> 0 Then
        lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    
    If vdtDataEmiDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataEmiDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
        
    If vdtDataEmiAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataEmiAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    
    If vlClienteDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlClienteDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    
    If vlClienteAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlClienteAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    
    If vdtDataBxDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataBxDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
        
    If vdtDataBxAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataBxAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
       
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    RelComissPur_PreparaInt = SUCESSO
     
    Exit Function
    
Erro_RelComissPur_PreparaInt:

    RelComissPur_PreparaInt = gErr
     
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 215025)
     
    End Select
     
    Exit Function

End Function


Public Function RelVendCliPur_Prepara(ByVal lNumIntRelCat As Long, ByVal iFilialEmpresaFiltro As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal lPedidoDe As Long, ByVal lPedidoAte As Long, ByVal sProdDe As String, ByVal sProdAte As String, lNumIntRel As Long) As Long
        'lErro = CF("RelVendCliPur_Prepara", lNumIntRelCat, giFilialEmpresa, StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text), StrParaLong(PedidoInicial.Text), StrParaLong(PedidoFinal.Text), sProd_I, sProd_F, lNumIntRel)

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 4) As Long, sSQL As String
Dim iSeq As Integer, lCodigo As Long, iFilialEmpresa As Integer, dtData As Date, lCliente As Long, iFilial As Integer, sNomeCliente As String
Dim sRazaoCliente As String, sCidade As String, iCondPagto As Integer, dValorPedido As Double, dValorProdFaixa As Double
Dim dValorPedFaixa As Double, iVendedor As Integer, sNomeVendedor As String, sNomeCondPagto As String, iFormaPagto As Integer, sNomeFormaPagto As String
Dim dValorProd As Double, lNumIntDocPed As Long

On Error GoTo Erro_RelVendCliPur_Prepara

    'abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211802
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211803
    Next
        
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUMINTREL_VENDCLIPUR", lNumIntRel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Monta o Select
    lErro = RelVendCliPurSQL_Prepara(iFilialEmpresaFiltro, dtDataDe, dtDataAte, lPedidoDe, lPedidoAte, sProdDe, sProdAte, sSQL)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sNomeCliente = String(STRING_MAXIMO, 0)
    sRazaoCliente = String(STRING_MAXIMO, 0)
    sCidade = String(STRING_MAXIMO, 0)
    sNomeVendedor = String(STRING_MAXIMO, 0)
    sNomeCondPagto = String(STRING_MAXIMO, 0)
    sNomeFormaPagto = String(STRING_MAXIMO, 0)

    'Passa os Binds
    lErro = RelVendCliPurInt_Prepara(alComando(2), lNumIntDocPed, lCodigo, iFilialEmpresa, dtData, lCliente, iFilial, sNomeCliente, sRazaoCliente, sCidade, iCondPagto, dValorPedido, dValorProd, iVendedor, sNomeVendedor, sNomeCondPagto, iFormaPagto, sNomeFormaPagto, lNumIntRelCat, iFilialEmpresaFiltro, dtDataDe, dtDataAte, lPedidoDe, lPedidoAte, sProdDe, sProdAte, sSQL)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211806
    
    iSeq = 0
    Do While lErro = AD_SQL_SUCESSO
    
        iSeq = iSeq + 1
                
        dValorPedFaixa = 0
        dValorProdFaixa = 0
        
        If sProdDe = "" And sProdAte = "" Then
            dValorPedFaixa = dValorPedido
            dValorProdFaixa = dValorProd
        Else
        
            If sProdDe = "" Then
                lErro = Comando_Executar(alComando(4), "SELECT SUM(PrecoTotal - ValorDescontoItem) AS ValorProd, SUM(PrecoTotal - ValorDescontoItem + ICMSSubstValor + IPIValor + ValorFreteItem + ValorOutrasDespesasItem + ValorSeguroItem - ICMSValorIsento) As ValorPed FROM TributacaoDocItem WHERE TipoDoc = 1 AND NumIntDoc = ? AND Produto <=? ", dValorProdFaixa, dValorPedFaixa, lNumIntDocPed, sProdAte)
            ElseIf sProdAte = "" Then
                lErro = Comando_Executar(alComando(4), "SELECT SUM(PrecoTotal - ValorDescontoItem) AS ValorProd, SUM(PrecoTotal - ValorDescontoItem + ICMSSubstValor + IPIValor + ValorFreteItem + ValorOutrasDespesasItem + ValorSeguroItem - ICMSValorIsento) As ValorPed FROM TributacaoDocItem WHERE TipoDoc = 1 AND NumIntDoc = ? AND Produto >=? ", dValorProdFaixa, dValorPedFaixa, lNumIntDocPed, sProdDe)
            Else
                lErro = Comando_Executar(alComando(4), "SELECT SUM(PrecoTotal - ValorDescontoItem) AS ValorProd, SUM(PrecoTotal - ValorDescontoItem + ICMSSubstValor + IPIValor + ValorFreteItem + ValorOutrasDespesasItem + ValorSeguroItem - ICMSValorIsento) As ValorPed FROM TributacaoDocItem WHERE TipoDoc = 1 AND NumIntDoc = ? AND Produto >=? AND Produto <=? ", dValorProdFaixa, dValorPedFaixa, lNumIntDocPed, sProdDe, sProdAte)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 211588
        
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211806
        
        End If
       
        lErro = Comando_Executar(alComando(3), "INSERT INTO RelVendCliPur (NumIntRel,Seq,Codigo,FilialEmpresa,Data,Cliente,Filial,NomeCliente,RazaoCliente,Cidade,CondPagto,ValorPedido,ValorProdFaixa,ValorPedFaixa,Vendedor,NomeVendedor,NomeCondPagto,FormaPagto,NomeFormaPagto) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                            lNumIntRel, iSeq, lCodigo, iFilialEmpresa, dtData, lCliente, iFilial, sNomeCliente, sRazaoCliente, sCidade, iCondPagto, dValorPedido, dValorProdFaixa, dValorPedFaixa, iVendedor, sNomeVendedor, sNomeCondPagto, iFormaPagto, sNomeFormaPagto)
        If lErro <> AD_SQL_SUCESSO Then gError 211805

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211806
     
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
     'fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211807

    RelVendCliPur_Prepara = SUCESSO

    Exit Function

Erro_RelVendCliPur_Prepara:

    RelVendCliPur_Prepara = gErr

    Select Case gErr

        Case 211802
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 211803
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211804
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELFATPRODVEND", gErr)

        Case 211805
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_COMPARATIVOANO", gErr)

        Case 211806
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFATFILCLI", gErr)
    
        Case 211807
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211808)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function RelVendCliPurSQL_Prepara(ByVal iFilialEmpresa As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal lPedidoDe As Long, ByVal lPedidoAte As Long, ByVal sProdDe As String, ByVal sProdAte As String, sSQL As String) As Long

Dim lErro As Long
Dim sSubWhere As String
Dim sTabela As String

On Error GoTo Erro_RelVendCliPurSQL_Prepara

    sSQL = "SELECT        P.NumIntDoc, P.FilialEmpresa, P.Codigo, P.Cliente, P.DataEmissao, P.CondicaoPagto, P.ValorTotal, P.ValorProdutos, P.Filial, C.NomeReduzido, C.RazaoSocial, E.Cidade, "
    sSQL = sSQL & "       V.Codigo, V.NomeReduzido, CP.Codigo, CP.DescReduzida, FP.Codigo, FP.Descricao "
    sSQL = sSQL & "FROM   (SELECT Cliente, Filial FROM RelFiltroFilCliCat WHERE NumIntRel = ?) AS X INNER JOIN"
    sSQL = sSQL & "       dbo.PedVenTodos AS P INNER JOIN"
    sSQL = sSQL & "       dbo.Clientes AS C ON P.Cliente = C.Codigo INNER JOIN"
    sSQL = sSQL & "       dbo.FiliaisClientes AS F ON P.Cliente = F.CodCliente AND P.Filial = F.CodFilial INNER JOIN"
    sSQL = sSQL & "       dbo.Enderecos AS E ON F.Endereco = E.Codigo ON X.Cliente = P.Cliente AND X.Filial = P.Filial LEFT OUTER JOIN"
    sSQL = sSQL & "       dbo.FormasPagamento AS FP RIGHT OUTER JOIN"
    sSQL = sSQL & "       dbo.CondicoesPagto AS CP ON FP.Codigo = CP.FormaPagamento ON P.CondicaoPagto = CP.Codigo LEFT OUTER JOIN"
    sSQL = sSQL & "       dbo.Vendedores AS V ON F.Vendedor = V.Codigo WHERE "
                         
    If iFilialEmpresa <> 0 Then
        sSQL = sSQL & "P.FilialEmpresa = ? "
    Else
        sSQL = sSQL & "P.FilialEmpresa < 50 "
    End If

    If dtDataDe <> DATA_NULA Then sSQL = sSQL & "AND P.DataEmissao >= ? "

    If dtDataAte <> DATA_NULA Then sSQL = sSQL & "AND P.DataEmissao <= ? "
    
    If lPedidoDe <> 0 Then sSQL = sSQL & "AND P.Codigo >= ? "

    If lPedidoAte <> 0 Then sSQL = sSQL & "AND P.Codigo <= ? "
    
    sSubWhere = ""
    If sProdDe <> "" Then sSubWhere = sSubWhere & "AND I.Produto >= ? "
    If sProdAte <> "" Then sSubWhere = sSubWhere & "AND I.Produto <= ? "
    If Len(Trim(sSubWhere)) <> 0 Then
        sSQL = sSQL & "AND EXISTS (SELECT * FROM ItensPedidoDeVendaTodos AS I WHERE P.Codigo = I.CodPedido AND P.FilialEmpresa = I.FilialEmpresa " & sSubWhere & ")"
    End If
    
    RelVendCliPurSQL_Prepara = SUCESSO
    
    Exit Function

Erro_RelVendCliPurSQL_Prepara:

    RelVendCliPurSQL_Prepara = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211809)

    End Select

    Exit Function

End Function

Private Function RelVendCliPurInt_Prepara(ByVal lComando As Long, vlNumIntDocPed As Variant, vlCodigo As Variant, viFilialEmpresa As Variant, vdtData As Variant, vlCliente As Variant, viFilial As Variant, vsNomeCliente As Variant, vsRazaoCliente As Variant, vsCidade As Variant, viCondPagto As Variant, vdValorPedido As Variant, vdValorProd As Variant, viVendedor As Variant, vsNomeVendedor As Variant, vsNomeCondPagto As Variant, viFormaPagto As Variant, vsNomeFormaPagto As Variant, ByVal lNumIntRelCat As Variant, ByVal viFilialEmpresaFiltro As Variant, ByVal vdtDataDe As Variant, ByVal vdtDataAte As Variant, ByVal vlPedidoDe As Variant, ByVal vlPedidoAte As Variant, ByVal vsProdDe As Variant, ByVal vsProdAte As Variant, ByVal sSQL As String) As Long

Dim lErro As Long

On Error GoTo Erro_RelVendCliPurInt_Prepara

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vlNumIntDocPed)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM

    lErro = Comando_BindVarInt(lComando, vlCodigo)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vlCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdtData)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viCondPagto)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdValorPedido)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vdValorProd)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viFilial)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsNomeCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsRazaoCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsCidade)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viVendedor)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsNomeVendedor)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viCondPagto)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsNomeCondPagto)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, viFormaPagto)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BindVarInt(lComando, vsNomeFormaPagto)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
       
    lErro = Comando_BindVarInt(lComando, lNumIntRelCat)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    
    If viFilialEmpresaFiltro <> 0 Then
        lErro = Comando_BindVarInt(lComando, viFilialEmpresaFiltro)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If

    If vdtDataDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    If vdtDataAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If

    If vlPedidoDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlPedidoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    If vlPedidoAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlPedidoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
       
    If vsProdDe <> "" Then
        lErro = Comando_BindVarInt(lComando, vsProdDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If
    If vsProdAte <> "" Then
        lErro = Comando_BindVarInt(lComando, vsProdAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM
    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError ERRO_SEM_MENSAGEM

    RelVendCliPurInt_Prepara = SUCESSO

    Exit Function

Erro_RelVendCliPurInt_Prepara:

    RelVendCliPurInt_Prepara = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211837)

    End Select

    Exit Function

End Function


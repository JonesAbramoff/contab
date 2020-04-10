VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelectPhar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Incluída leitura dos campos ProdutoKit, QuantKit
Function NFiscalItens_Le_Customizado(ByVal objNFiscal As ClassNFiscal) As Long
'Lê os valores customizados dos ítens da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim lComando As Long
Dim objItemNFPhar As ClassItemNF
Dim objItemNF As ClassItemNF
Dim sProdutoKit As String
Dim dQuantKit As Double
Dim sUMKit As String

On Error GoTo Erro_NFiscalItens_Le_Customizado

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 197684

    For Each objItemNF In objNFiscal.ColItensNF
        
        sProdutoKit = String(STRING_PRODUTO, 0)
        sUMKit = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        
        'Pesquisa os ítens vinculados à Nota Fiscal
        lErro = Comando_Executar(lComando, "SELECT ProdutoKit, QuantKit, UMKit FROM ItensNFiscal WHERE NumIntDoc = ? ", sProdutoKit, dQuantKit, sUMKit, objItemNF.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 197685

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197686

        If lErro = AD_SQL_SUCESSO Then
        
            If Len(Trim(sProdutoKit)) > 0 Then
        
                Set objItemNFPhar = New ClassItemNF
                
                objItemNFPhar.sProduto = sProdutoKit
                objItemNFPhar.dQuantidade = dQuantKit
                objItemNFPhar.sUnidadeMed = sUMKit
                
                Set objItemNF.objInfoUsu = objItemNFPhar

            End If

        End If

    Next

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    NFiscalItens_Le_Customizado = SUCESSO

    Exit Function

Erro_NFiscalItens_Le_Customizado:

    NFiscalItens_Le_Customizado = gErr

    Select Case gErr

        Case 197684
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197685, 197686
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197687)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function EstoqueTerc_Le_Reserva(ByVal objNF As ClassNFiscal, ByVal objEstTerc As ClassEstoqueTerc, ByVal objProduto As ClassProduto, dQtdeReserv As Double) As Long

Dim lErro As Long
Dim lComando As Long
Dim sSQL As String, dQtd As Double, sUM As String
Dim dFator As Double

On Error GoTo Erro_EstoqueTerc_Le_Reserva

    dQtdeReserv = 0

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 211063
    
    'Le todas as NFs que vão ser entregues por esse terceiro para esse produto e que não tenham
    'dado entrada na devolução simbólica ou que não serão enviadas ao cliente
    sSQL = "SELECT SUM(I.Quantidade), I.UnidadeMed FROM NFiscal AS N, ItensNFiscal AS I "
    sSQL = sSQL & "WHERE N.NumIntDoc = I.NumIntNF AND N.Status <> 7 AND N.TipoNFiscal = 233 AND I.Produto = ? AND N.FornEntTerc = ? AND N.FilialFornEntTerc = ? AND N.FilialEmpresa = ? AND N.DataEmissao >= ? AND "
    sSQL = sSQL & "NOT EXISTS (SELECT * FROM PharLMLogSEP AS S WHERE S.NumPedido = N.NumNotaFiscal AND S.DataPedido = N.DataEmissao AND (S.NumIntNF <> 0 OR S.Situacao <> 'SE')) AND N.NumNotaFiscal <> ? GROUP BY I.UnidadeMed"

    sUM = String(STRING_UM_SIGLA, 0)

    lErro = Comando_Executar(lComando, sSQL, dQtd, sUM, objEstTerc.sProduto, objEstTerc.lCliForn, objEstTerc.iFilial, objEstTerc.iFilialEmpresa, #10/21/2011#, objNF.lNumNotaFiscal)
    If lErro <> AD_SQL_SUCESSO Then gError 211064

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211065
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUM, dFator)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dQtdeReserv = dQtdeReserv + dQtd * dFator
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211066
    
    Loop
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)

    EstoqueTerc_Le_Reserva = SUCESSO

    Exit Function

Erro_EstoqueTerc_Le_Reserva:

    EstoqueTerc_Le_Reserva = gErr

    Select Case gErr

        Case 211063
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211064 To 211066
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211067)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Public Function RastreamentoIndireto_Le(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long, lComando As Long
Dim objItemNF As ClassItemNF
Dim objRastroItemNF As ClassRastroItemNF
Dim sProdMask As String, dQtde As Double, sProdFormat As String
Dim sLote As String, dtValLote As Date

Const OL_CODFORN_LMLOG = 403

On Error GoTo Erro_RastreamentoIndireto_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 211187
   
    'Se for uma NF que foi entregue pela LM
    If objNF.lFornEntTerc = OL_CODFORN_LMLOG Then
   
        sProdMask = String(STRING_MAXIMO, 0)
        sLote = String(STRING_MAXIMO, 0)
    
        lErro = Comando_Executar(lComando, "SELECT Produto, Quantidade, Lote, ValidadeLote FROM PharLMLogSepItens WHERE NumPedido = ? AND DataImport BETWEEN ? AND ? AND Lote <> '' ORDER BY Produto ", sProdMask, dQtde, sLote, dtValLote, objNF.lNumNotaFiscal, objNF.dtDataEmissao - PERIODO_EMISSAO, objNF.dtDataEmissao + PERIODO_EMISSAO)
        If lErro <> AD_SQL_SUCESSO Then gError 211188
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211189
    
        Do While lErro <> AD_SQL_SEM_DADOS
            
            sProdFormat = Replace(sProdMask, ".", "")
            sProdFormat = sProdFormat & String(10 - Len(sProdFormat), " ")
            
            For Each objItemNF In objNF.ColItensNF
            
                If objItemNF.sProduto = sProdFormat Then

                    Set objRastroItemNF = New ClassRastroItemNF
                    
                    objRastroItemNF.sLote = Trim(sLote)
                    objRastroItemNF.sProduto = objItemNF.sProduto
                    objRastroItemNF.dLoteQdtAlocada = dQtde
                    
                    If Len(objRastroItemNF.sLote) > 0 Then objItemNF.colRastreamento.Add objRastroItemNF
                
                End If
            Next
        
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211190
        
        Loop
    
    End If
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)

    RastreamentoIndireto_Le = SUCESSO
    
    Exit Function
    
Erro_RastreamentoIndireto_Le:

    RastreamentoIndireto_Le = gErr
    
    Select Case gErr

        Case 211187
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211188 To 211190
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PHARLMLOGSEPITENS", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211192)

    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

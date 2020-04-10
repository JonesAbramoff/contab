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

Const TIPOPRODUTO_SERVICO = 15
Const REMEMISSAONF_NUM_LINHAS_PRODUTO = 13
Const REMEMISSAONF_NUM_LINHAS_SERVICO = 4

'##################################################################
'Inserido por Wagner 24/10/2005
Public Function RelEmissaoNF_Prepara(ByVal lNumNotaIni As Long, ByVal lNumNotaFim As Long, ByVal sSerie As String, ByVal iFilialEmpresa As Integer, lNumIntRel As Long) As Long

Dim lErro As Long
Dim lErro1 As Long
Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lNumIntDocNF As Long
Dim lNumIntDocNFAnt As Long
Dim lNumIntDocItemNF As Long
Dim iPagProdutos As Integer
Dim iPagServicos As Integer
Dim iLinhaProdutos As Integer
Dim iLinhaServicos As Integer
Dim iTipo As Integer
Dim iLinha As Integer
Dim iPag As Integer
Dim iTotPag As Integer
Dim lAux As Long
Dim iTipoProduto As Integer

On Error GoTo Erro_RelEmissaoNF_Prepara

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140581
    Next
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 140582

    'Obtém o NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_EMISSAONF", lNumIntRel)
    If lErro <> SUCESSO Then gError 140583
    
    'Le os Itens da NF que farão parte do relatório
    lErro = Comando_Executar(alComando(1), "SELECT NF.NumIntDoc, NF.NumIntDocItemNF, P.Tipo FROM NFiscais_FornCli AS NF, Produtos AS P WHERE P.Codigo = NF.Produto AND NF.Emitente = 0 AND NF.Serie = ? AND (NF.NumNotaFiscal >= ? AND NF.NumNotaFiscal <= ?) AND NF.FilialEmpresa = ? ORDER BY NF.NumIntDoc, NF.NumIntDocItemNF", _
    lNumIntDocNF, lNumIntDocItemNF, iTipo, sSerie, lNumNotaIni, lNumNotaFim, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 140584
    
    'Busca o próximo registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140585
        
    iPagServicos = 1
    iPagProdutos = 1
            
    'Enquanto houver registros
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Se for um serviço
        If iTipo = TIPOPRODUTO_SERVICO Then
            iLinhaServicos = iLinhaServicos + 1
            iLinha = iLinhaServicos
            If iLinhaServicos Mod (REMEMISSAONF_NUM_LINHAS_SERVICO + 1) = 0 Then iPagServicos = iPagServicos + 1
            iPag = iPagServicos
            iTipoProduto = 2
        Else
            iLinhaProdutos = iLinhaProdutos + 1
            iLinha = iLinhaProdutos
            If iLinhaProdutos Mod (REMEMISSAONF_NUM_LINHAS_PRODUTO + 1) = 0 Then iPagProdutos = iPagProdutos + 1
            iPag = iPagProdutos
            iTipoProduto = 1
        End If

        'Insere em RelEmissaoNF
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelEmissaoNF (NumIntRel, NumIntDocNF, NumIntDocItemNF, ProdServ, Linha, Pagina, TotalPaginas) VALUES (?,?,?,?,?,?,?)", _
                lNumIntRel, lNumIntDocNF, lNumIntDocItemNF, iTipoProduto, iLinha, iPag, iTotPag)
        If lErro <> AD_SQL_SUCESSO Then gError 140586

        lNumIntDocNFAnt = lNumIntDocNF

        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140587

        'Se mudou a NF ou se acabou os registros, atualiza o total de páginas da NF anterior
        If lNumIntDocNF <> lNumIntDocNFAnt Or lErro = AD_SQL_SEM_DADOS Then
        
            If iPagProdutos < iPagServicos Then
                iTotPag = iPagServicos
            Else
                iTotPag = iPagProdutos
            End If
            
            'Le os registros da NF anterior
            lErro1 = Comando_ExecutarPos(alComando(3), "SELECT NumIntDocNF FROM RelEmissaoNF WHERE NumIntDocNF = ? AND NumIntRel = ?", 0, lAux, lNumIntDocNFAnt, lNumIntRel)
            If lErro1 <> AD_SQL_SUCESSO Then gError 140588
            
            'Busca o próximo registro
            lErro1 = Comando_BuscarPrimeiro(alComando(3))
            If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 140589
            
            Do While lErro1 <> AD_SQL_SEM_DADOS
        
                'Atualiza Total de Páginas
                lErro1 = Comando_ExecutarPos(alComando(2), "UPDATE RelEmissaoNF SET TotalPaginas = ? ", alComando(3), iTotPag)
                If lErro1 <> AD_SQL_SUCESSO Then gError 140590
        
                'Busca o próximo registro
                lErro1 = Comando_BuscarProximo(alComando(3))
                If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 140591
        
            Loop
        
            'Reinicia Contadores
            iLinhaServicos = 0
            iLinhaProdutos = 0
            iPagServicos = 1
            iPagProdutos = 1
            
        End If

    Loop
    
    'confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 140592
 
    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelEmissaoNF_Prepara = SUCESSO

    Exit Function

Erro_RelEmissaoNF_Prepara:

    RelEmissaoNF_Prepara = gErr

    Select Case gErr

        Case 140581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140582
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 140583

        Case 140584 To 140585, 140587 To 140589, 140591
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)

        Case 140586
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELEMISSAONF", gErr)
        
        Case 140590
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RELEMISSAONF", gErr)
        
        Case 140592
             Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150066)

    End Select

    Call Transacao_Rollback
    
    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function
'##################################################################


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function ProdutoKit_Le_InfoUsu(objProdutoKit As ClassProdutoKit) As Long

Dim lErro As Long
Dim lComando As Long
Dim objProdutoKitCro As New ClassProdutoKitCro
Dim iTipoCarga As Integer, sGrupoPesagem As String
    
On Error GoTo Erro_ProdutoKit_Le_InfoUsu

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 184298

    sGrupoPesagem = String(STRING_GRUPO_PESAGEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT TipoCarga, GrupoPesagem FROM ProdutoKit WHERE ProdutoRaiz = ? AND Versao = ? AND Nivel = ? AND Seq = ?", _
        iTipoCarga, sGrupoPesagem, objProdutoKit.sProdutoRaiz, objProdutoKit.sVersao, objProdutoKit.iNivel, objProdutoKit.iSeq)
    If lErro <> AD_SQL_SUCESSO Then gError 184299

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184300
    If lErro <> AD_SQL_SUCESSO Then gError 184301

    objProdutoKitCro.iTipoCarga = iTipoCarga
    objProdutoKitCro.sGrupoPesagem = sGrupoPesagem
    Set objProdutoKit.objInfoUsu = objProdutoKitCro
    
    'Fechamento comando
    Call Comando_Fechar(lComando)
    
    ProdutoKit_Le_InfoUsu = SUCESSO
    
    Exit Function
    
Erro_ProdutoKit_Le_InfoUsu:

    ProdutoKit_Le_InfoUsu = gErr

    Select Case gErr

        Case 184298
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 184299, 184300, 184301
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184297)

    End Select
    
    'Fechamento comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Function Kit_Le_ComponenteBase(ByVal objProdutoKit As ClassProdutoKit) As Long
''Preenche objKit.colComponentes a partir do produto e da versão passados em objKit
'
'Dim lErro As Long
'Dim lComando As Long
'Dim tProdutoKit As typeProdutoKit
'Dim objProdutoKitCro As ClassProdutoKitCro
'Dim sGrupoPesagem As String, iTipoCarga As Integer
'
'On Error GoTo Erro_Kit_Le_ComponenteBase
'
'    'Abertura comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then Error 21828
'
'    tProdutoKit.sProdutoRaiz = String(STRING_PRODUTO, 0)
'    tProdutoKit.sProduto = String(STRING_PRODUTO, 0)
'    tProdutoKit.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)
'    tProdutoKit.sVersao = String(STRING_KIT_VERSAO, 0)
'    sGrupoPesagem = String(STRING_GRUPO_PESAGEM, 0)
'
'    lErro = Comando_Executar(lComando, "SELECT ProdutoRaiz, Versao, Nivel, Seq, Produto, SeqPai, Quantidade, UnidadeMed, Composicao, PosicaoArvore, PercentualPerda, CustoStandard, TipoCarga,GrupoPesagem FROM ProdutoKit WHERE ProdutoRaiz = ? AND Versao = ? AND TipoCarga =? ORDER BY Nivel, Seq ", tProdutoKit.sProdutoRaiz, tProdutoKit.sVersao, tProdutoKit.iNivel, tProdutoKit.iSeq, tProdutoKit.sProduto, tProdutoKit.iSeqPai, tProdutoKit.dQuantidade, tProdutoKit.sUnidadeMed, tProdutoKit.iComposicao, tProdutoKit.iPosicaoArvore, tProdutoKit.dPercentualPerda, tProdutoKit.dCustoStandard, iTipoCarga, sGrupoPesagem, objProdutoKit.sProdutoRaiz, objProdutoKit.sVersao, PRODUTOKIT_BASECARGA)
'    If lErro <> AD_SQL_SUCESSO Then Error 21829
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21830
'
'    If lErro = AD_SQL_SEM_DADOS Then Error 21831
'
'    objProdutoKit.dQuantidade = tProdutoKit.dQuantidade
'    objProdutoKit.iComposicao = tProdutoKit.iComposicao
'    objProdutoKit.iNivel = tProdutoKit.iNivel
'    objProdutoKit.iSeq = tProdutoKit.iSeq
'    objProdutoKit.iSeqPai = tProdutoKit.iSeqPai
'    objProdutoKit.sProduto = tProdutoKit.sProduto
'    objProdutoKit.sUnidadeMed = tProdutoKit.sUnidadeMed
'    objProdutoKit.sProdutoRaiz = tProdutoKit.sProdutoRaiz
'    objProdutoKit.sVersao = tProdutoKit.sVersao
'    objProdutoKit.iPosicaoArvore = tProdutoKit.iPosicaoArvore
'    objProdutoKit.dPercentualPerda = tProdutoKit.dPercentualPerda
'    objProdutoKit.dCustoStandard = tProdutoKit.dCustoStandard
'
'    objProdutoKitCro.iTipoCarga = iTipoCarga
'    objProdutoKitCro.sGrupoPesagem = sGrupoPesagem
'    Set objProdutoKit.objInfoUsu = objProdutoKitCro
'
'    'Fechamento comando
'    Call Comando_Fechar(lComando)
'
'    Kit_Le_ComponenteBase = SUCESSO
'
'    Exit Function
'
'Erro_Kit_Le_ComponenteBase:
'
'    Kit_Le_ComponenteBase = Err
'
'    Select Case Err
'
'        Case 21828
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 21829, 21830, 21832
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", Err)
'
'        Case 21831
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)
'
'    End Select
'
'   'Fechamento comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

'###############################################
'Inserido por Wagner 16/09/05
Function MovEstoque_Le_Customizado(ByVal objMovEstoque As ClassMovEstoque, ByVal lComando As Long) As Long

Dim lErro As Long
Dim sResponsavel As String
Dim objMovEstoqueCro As New ClassMovEstoqueCro

On Error GoTo Erro_MovEstoque_Le_Customizado

    sResponsavel = String(STRING_MOVESTOQUE_RESPONSAVEL, 0)

    lErro = Comando_Executar(lComando, "SELECT Responsavel FROM MovimentoEstoque WHERE FilialEmpresa = ? AND Codigo = ? AND NumIntDocEst = 0", sResponsavel, giFilialEmpresa, objMovEstoque.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 140238

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140239

    If lErro = AD_SQL_SEM_DADOS Then gError 140240

    objMovEstoqueCro.sResponsavel = sResponsavel

    Set objMovEstoque.objInfoUsu = objMovEstoqueCro
    
    MovEstoque_Le_Customizado = SUCESSO

    Exit Function

Erro_MovEstoque_Le_Customizado:

    MovEstoque_Le_Customizado = gErr

    Select Case gErr
        
        Case 140238, 140239
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 140240 'Tratado na Rotina Chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Function

End Function
'###############################################

'Inserido por Wagner 10/12/2004
Function Kit_Valida_Quantidade(ByVal objKit As ClassKit) As Long

Dim lErro As Long, sUMRaiz As String
Dim objProdutoKit As ClassProdutoKit
Dim dFator As Double
Dim dQtdProdutoRaiz As Double
Dim dQtdTotalSemRaiz As Double
Dim iClasseProdutoRaiz As Integer
Dim bClasseDiferente As Boolean
Dim objProduto As ClassProduto
Dim vbMsgResult As VbMsgBoxResult

On Error GoTo Erro_Kit_Valida_Quantidade

    dQtdProdutoRaiz = 0
    dQtdTotalSemRaiz = 0
    bClasseDiferente = False

    For Each objProdutoKit In objKit.colComponentes
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objProdutoKit.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 131681
        If lErro = 28030 Then gError 131682
    
        If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
            
            sUMRaiz = objProdutoKit.sUnidadeMed
            dQtdProdutoRaiz = objProdutoKit.dQuantidade
            iClasseProdutoRaiz = objProduto.iClasseUM
        
        Else
            
            If objProdutoKit.iNivel = (KIT_NIVEL_RAIZ + 1) Then
            
                If iClasseProdutoRaiz <> objProduto.iClasseUM Then
                    bClasseDiferente = True
                    Exit For
                End If
                
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, sUMRaiz, dFator)
                If lErro <> SUCESSO Then gError 131684
    
                dQtdTotalSemRaiz = dQtdTotalSemRaiz + (objProdutoKit.dQuantidade * dFator)
                
            End If
        
        End If
    
    Next
    
    If Not bClasseDiferente Then
    
        If Abs(dQtdProdutoRaiz - dQtdTotalSemRaiz) > QTDE_ESTOQUE_DELTA Then
        
            gError 131683
        
        End If
    
    End If

    Exit Function

Erro_Kit_Valida_Quantidade:

    Kit_Valida_Quantidade = gErr

    Select Case gErr
    
        Case 131681
        
        Case 131682
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
    
        Case 131683, 131684
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDE_RAIZ_DIFERE_QTDE_COMPONENTES", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function

End Function

'###############################################
'Inserido por Wagner 23/02/2007
Function RastreamentoLote_Le_Customizado(ByVal objRastroLote As ClassRastreamentoLote) As Long
'Grava o RastreamentoLote

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 0) As Long
Dim iFilialForn As Integer
Dim lFornecedor As Long, sLocalizacao As String
Dim objRastroLoteCro As New ClassRastreamentoLoteCro

On Error GoTo Erro_RastreamentoLote_Le_Customizado

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185928
    Next
   
    sLocalizacao = String(STRING_RASTRO_LOCALIZACAO, 0)
    
    'Verifica se o RastreamentoLote já está cadastrado
    lErro = Comando_Executar(alComando(0), "SELECT Fornecedor, FilialForn, Localizacao FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", lFornecedor, iFilialForn, sLocalizacao, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)
    If lErro <> AD_SQL_SUCESSO Then gError 185929

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185930
        
    'Se encontrou o RastreamentoLote, ALTERAÇÃO
    If lErro = AD_SQL_SUCESSO Then

        objRastroLoteCro.lFornecedor = lFornecedor
        objRastroLoteCro.iFilialForn = iFilialForn
        objRastroLoteCro.sLocalizacao = sLocalizacao

    End If
    
    Set objRastroLote.objInfoUsu = objRastroLoteCro

    'Fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RastreamentoLote_Le_Customizado = SUCESSO

    Exit Function

Erro_RastreamentoLote_Le_Customizado:

    RastreamentoLote_Le_Customizado = gErr

    Select Case gErr

        Case 185928
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185929, 185930
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RastroLoteSaldo_Obter_QuantOP_Customizado(ByVal lNumIntDocLote As Long, dQuantOP As Double) As Long
'Grava o RastreamentoLote

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim dQuantReq As Double
Dim dQuantReqTot As Double
Dim dQuantOPAux As Double
Dim dQuantOPAuxTot As Double
Dim sProduto As String
Dim sLote As String
Dim lNumIntItemOP As Long

On Error GoTo Erro_RastroLoteSaldo_Obter_QuantOP_Customizado

    dQuantOP = 0

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185934
    Next
   
    sProduto = String(STRING_PRODUTO, 0)
    sLote = String(STRING_LOTE_RASTREAMENTO, 0)
   
    'Le a quantidade nos itens não baixados de op
    lErro = Comando_Executar(alComando(0), "SELECT SUM(C.QtdCarga * CI.KGCarga) AS QuantOPAux, R.Produto, R.Lote, C.NumIntItemOP FROM ItensOPCargaInsumo CI, ItensOPCarga C, ItensOrdemProducao I, RastreamentoLote R WHERE CI.NumIntItemOP = C.NumIntItemOP AND CI.NumIntItemOP = I.NumIntDoc AND R.Produto = CI.Produto AND R.Lote = CI.Lote AND I.Situacao <> ? AND R.NumIntDoc = ? GROUP BY R.Produto, R.Lote, C.NumIntItemOP", dQuantOPAux, sProduto, sLote, lNumIntItemOP, ITEMOP_SITUACAO_BAIXADA, lNumIntDocLote)
    If lErro <> AD_SQL_SUCESSO Then gError 185935

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185936
        
    'Abate o que já foi requisitado
    Do While lErro = AD_SQL_SUCESSO
    
        dQuantOPAuxTot = dQuantOPAuxTot + dQuantOPAux
    
        'Le as requisições de produção
        lErro = Comando_Executar(alComando(1), "SELECT SUM(ME.Quantidade) FROM RastreamentoLote AS L, RastreamentoMovto AS R, MovimentoEstoque AS ME " & _
                                            "WHERE L.NumIntDoc = R.NumIntDocLote AND R.NumIntDocOrigem = ME.NumIntDoc AND ME.TipoMov IN (?,?,?) AND ME.NumIntDocEst=0 AND (R.TipoDocOrigem = 0 OR R.TipoDocOrigem IS NULL) AND ME.TipoNumIntDocOrigem=5 AND ME.Produto = ? AND L.Lote = ? AND ME.NumIntDocOrigem = ? ", _
                                            dQuantReq, MOV_EST_REQ_PRODUCAO_BENEF3, MOV_EST_REQ_PRODUCAO, MOV_EST_DEV_MATERIAL_PRODUCAO, sProduto, sLote, lNumIntItemOP)
        If lErro <> AD_SQL_SUCESSO Then gError 185937
            
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185938

        'Abate o que já foi requisitado
        If lErro = AD_SQL_SUCESSO Then dQuantReqTot = dQuantReqTot + dQuantReq

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185936

    Loop
    
    'A quantidade em OP é a quantidade nos insumos da OP - o que já foi requisitado para o produto\lote em questão
    dQuantOP = dQuantOPAuxTot - dQuantReqTot
    If dQuantOP < 0 Then dQuantOP = 0

    'Fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RastroLoteSaldo_Obter_QuantOP_Customizado = SUCESSO

    Exit Function

Erro_RastroLoteSaldo_Obter_QuantOP_Customizado:

    RastroLoteSaldo_Obter_QuantOP_Customizado = gErr

    Select Case gErr

        Case 185934
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185935, 185936
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPCARGA", gErr)
            
        Case 185937, 185938
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
'######################################################

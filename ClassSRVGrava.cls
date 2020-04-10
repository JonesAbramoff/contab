VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSRVGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function SRVConfig_Grava_Configs(ByVal colConfigs As Collection) As Long
'grava na tabela SRVConfig uma colecao de registros de configuracao

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_SRVConfig_Grava_Configs
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 183396

    lErro = CF("SRVConfig_Grava_Configs_Trans", colConfigs)
    If lErro <> SUCESSO Then gError 183397

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 183398
          
    SRVConfig_Grava_Configs = SUCESSO
    
    Exit Function
    
Erro_SRVConfig_Grava_Configs:
    
    SRVConfig_Grava_Configs = gErr
    
    Select Case gErr
    
        Case 183396
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 183397

        Case 183398
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183399)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function SRVConfig_Grava_Configs_Trans(ByVal colConfigs As Collection) As Long
'grava na tabela SRVConfig uma colecao de registros de configuracao
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long
Dim objSRVConfig As ClassSRVConfig

On Error GoTo Erro_SRVConfig_Grava_Configs_Trans
    
   'para cada elemento da colecao
    For Each objSRVConfig In colConfigs
    
        lErro = SRVConfig_Grava_Trans(objSRVConfig)
        If lErro <> SUCESSO Then gError 183400
        
    Next
          
    SRVConfig_Grava_Configs_Trans = SUCESSO
    
    Exit Function
    
Erro_SRVConfig_Grava_Configs_Trans:
    
    SRVConfig_Grava_Configs_Trans = gErr
    
    Select Case gErr
    
        Case 183400
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183401)
            
    End Select
    
    Exit Function
    
End Function

Private Function SRVConfig_Grava_Trans(ByVal objSRVConfig As ClassSRVConfig) As Long
'grava registro em SRVConfig. Apenas altera o conteudo, se nao existir dá erro.
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, tSRVConfig As typeSRVConfig
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_SRVConfig_Grava_Trans

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183402
    Next
    
    'inicializar buffer
    tSRVConfig.sConteudo = String(STRING_SRVCONFIG_CONTEUDO, 0)
    'ler registo
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Conteudo FROM SRVConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, tSRVConfig.sConteudo, objSRVConfig.sCodigo, objSRVConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 183403
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183404
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then gError 183405
        
    'se o conteudo foi modificado
    If objSRVConfig.sConteudo <> tSRVConfig.sConteudo Then
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE SRVConfig SET Conteudo = ?", alComando(1), objSRVConfig.sConteudo)
        If lErro <> AD_SQL_SUCESSO Then gError 183406
    End If
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    SRVConfig_Grava_Trans = SUCESSO
    
    Exit Function
    
Erro_SRVConfig_Grava_Trans:

    SRVConfig_Grava_Trans = gErr
     
    Select Case gErr
     
        Case 183402
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
     
        Case 183403, 183404
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SRVCONFIG", gErr, objSRVConfig.sCodigo, objSRVConfig.iFilialEmpresa)
     
        Case 183405
            Call Rotina_Erro(vbOKOnly, "ERRO_SRVCONFIG_INEXISTENTE", gErr, objSRVConfig.sCodigo, objSRVConfig.iFilialEmpresa)
        
        Case 183406
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SRVCONFIG", gErr, objSRVConfig.sCodigo, objSRVConfig.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183407)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function SolicitacaoSRV_Grava(ByVal objSolicSRV As ClassSolicSRV) As Long
'grava na tabela SolicitacaoSRV uma solicitacao

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 1) As Long
Dim tSolicSRV As typeSolicSRV
Dim iIndice As Integer, bNovoReg As Boolean

On Error GoTo Erro_SolicitacaoSRV_Grava
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 183704

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183705
    Next
    
    bNovoReg = False

    'Faz o lock nos objetos vinculados ao Pedido
    lErro = CF("SolicSRVObjVinc_Lock", objSolicSRV)
    If lErro <> SUCESSO Then gError 183706

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Data, Hora, Cliente, FilialCliente, NumIntDoc, Vendedor, Atendente FROM SolicitacaoSRV WHERE FilialEmpresa = ? AND Codigo = ?", 0, _
    tSolicSRV.dtData, tSolicSRV.dHora, tSolicSRV.lCliente, tSolicSRV.iFilial, tSolicSRV.lNumIntDoc, tSolicSRV.iVendedor, tSolicSRV.iAtendente, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 183707
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183708
        
    If lErro = AD_SQL_SEM_DADOS Then
    
        bNovoReg = True

        'Insere o Pedido novo no Banco de Dados
        lErro = SolicitacaoSRV_Insere_BD(objSolicSRV)
        If lErro <> SUCESSO Then gError 183709

    Else
    
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 183710
    
        objSolicSRV.lNumIntDoc = tSolicSRV.lNumIntDoc
    
        'Altera o Pedido de Venda e armazena os ítens do Pedido no BD em objPedVenda_BD
        lErro = SolicitacaoSRV_Altera_BD(alComando(1), objSolicSRV)
        If lErro <> SUCESSO Then gError 183711
    
    End If
    
    If objSolicSRV.iGravarCRM = MARCADO Then
        If Not (objSolicSRV.objRelacCli Is Nothing) Then
            objSolicSRV.objRelacCli.lNumIntDocOrigem = objSolicSRV.lNumIntDoc
            objSolicSRV.objRelacCli.iTipoDoc = RELACCLI_TIPODOC_SOLSRV
            
            If bNovoReg Then
                objSolicSRV.objRelacCli.lCodigo = 0
            End If
                    
            lErro = CF("RelacionamentoClientes_Grava_Trans", objSolicSRV.objRelacCli)
            If lErro <> SUCESSO Then gError 183709
    
        End If
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 183712

    lErro = CF("WorkFlow_Trata_Transacao", "SRV", objSolicSRV.objTela, gsUsuario, objSolicSRV)
    If lErro <> SUCESSO Then gError 183713
          
    SolicitacaoSRV_Grava = SUCESSO
    
    Exit Function
    
Erro_SolicitacaoSRV_Grava:
    
    SolicitacaoSRV_Grava = gErr
    
    Select Case gErr
    
        Case 183704
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 183705
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183706, 183709, 183711, 183713
        
        Case 183707, 183708
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSRV", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
        
        Case 183710
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_SOLICITACAOSRV", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
        
        Case 183712
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183714)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function SolicSRVObjVinc_Lock(ByVal objSolicSRV As ClassSolicSRV) As Long
'Faz "lock" nos objetos associados a Solicitacao de Servico

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim objItensSolicSrv As ClassItensSolicSRV
Dim objProduto As New ClassProduto
Dim objRastroLote As New ClassRastreamentoLote
Dim objFilialCliente As New ClassFilialCliente
Dim objContrato As New ClassContrato
Dim objGarantia As New ClassGarantia
Dim iIndice As Integer

On Error GoTo Erro_SolicSRVObjVinc_Lock

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183650
    Next

    'Loca a FilialCliente
    lErro = CF("FilialCliente_Lock", objSolicSRV.lCliente, objSolicSRV.iFilial, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 43050 Then gError 183651
    
    If lErro <> SUCESSO Then gError 183652
    
    If objFilialCliente.iAtivo <> CLIENTE_ATIVO Then gError 183672

    If objSolicSRV.iAtendente <> 0 Then

        'Efetua lock do atendente
        lErro = CF("Atendentes_LockShared", objSolicSRV.iAtendente, objSolicSRV.iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 102552 Then gError 183653

        If lErro <> SUCESSO Then gError 183654

    End If

    If objSolicSRV.iVendedor <> 0 Then

        lErro = CF("Vendedor_Lock", alComando(1), objSolicSRV.iVendedor)
        If lErro <> SUCESSO Then gError 183655

    End If

    'Para cada item de Pedido Venda
    For Each objItensSolicSrv In objSolicSRV.colItens

        objProduto.sCodigo = objItensSolicSrv.sProduto
        
        'Locka  o produto
        lErro = CF("Produto_Lock_Shared", alComando(2), objProduto)
        If lErro <> SUCESSO Then gError 183656

        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 183673

        If gobjSRV.iVerificaLote = VERIFICA_LOTE And objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then

            objRastroLote.sProduto = objItensSolicSrv.sProduto
            objRastroLote.sCodigo = objItensSolicSrv.sLote
            objRastroLote.iFilialOP = objItensSolicSrv.iFilialOP
    
            'Lê rastreamento do lote a partir do produto, filialOP e código do lote passados
            lErro = CF("RastreamentoLote_Le1", alComando(3), objRastroLote)
            If lErro <> SUCESSO And lErro <> 71889 Then gError 183657
            
            If lErro <> SUCESSO Then gError 183658

        End If

        objProduto.sCodigo = objItensSolicSrv.sServico
        
        'Locka  o servico
        lErro = CF("Produto_Lock_Shared", alComando(4), objProduto)
        If lErro <> SUCESSO Then gError 183659
        
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 183674
        
        If objItensSolicSrv.lGarantia <> 0 And gobjSRV.iValidaGarantia = VALIDA_GARANTIA Then
        
            objGarantia.iFilialEmpresa = objSolicSRV.iFilialEmpresa
            objGarantia.lCodigo = objItensSolicSrv.lGarantia
        
            'Locka  o servico
            lErro = CF("Garantia_LockShared", objGarantia)
            If lErro <> SUCESSO And lErro <> 183646 Then gError 183660
        
            If lErro <> SUCESSO Then gError 183661
            
            If objGarantia.iAtivo <> GARANTIA_ATIVA Then gError 183677
            
        End If
        
        If Len(Trim(objItensSolicSrv.sContrato)) <> 0 And gobjSRV.iValidaManutencao = VALIDA_MANUTENCAO Then
        
            'Locka  o servico
            lErro = CF("Contrato_Lock_Shared", objSolicSRV.iFilialEmpresa, objItensSolicSrv.sContrato, objContrato)
            If lErro <> SUCESSO And lErro <> 183652 Then gError 183662
        
            If lErro <> SUCESSO Then gError 183663
            
            If objContrato.iAtivo <> CONTRATO_ATIVO Then gError 183676
            
        End If
        
        
    Next

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

     SolicSRVObjVinc_Lock = SUCESSO

    Exit Function

Erro_SolicSRVObjVinc_Lock:

    SolicSRVObjVinc_Lock = gErr

    Select Case gErr

        Case 183650
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183651, 183653, 183655, 183656, 183657, 183659, 183660, 183662

        Case 183652
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objSolicSRV.iFilial, objSolicSRV.lCliente)

        Case 183654
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_NAO_ENCONTRADO", gErr, objSolicSRV.iAtendente, objSolicSRV.iFilialEmpresa)

        Case 183658
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_CADASTRADO_RASTREAMENTO", gErr, objItensSolicSrv.sLote, objItensSolicSrv.sProduto, objItensSolicSrv.iFilialOP)

        Case 183661
            Call Rotina_Erro(vbOKOnly, "ERRO_GARANTIA_INEXISTENTE", gErr, objSolicSRV.iFilialEmpresa, objItensSolicSrv.lGarantia)

        Case 183663
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_INEXISTENTE", gErr, objSolicSRV.iFilialEmpresa, objItensSolicSrv.sContrato)

        Case 183672
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_INATIVA", gErr, objSolicSRV.iFilial, objSolicSRV.lCliente)

        Case 183673, 183674
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, objProduto.sCodigo)

        Case 183676
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_INATIVO", gErr, objSolicSRV.iFilialEmpresa, objItensSolicSrv.sContrato)

        Case 183677
            Call Rotina_Erro(vbOKOnly, "ERRO_GARANTIA_INATIVA", gErr, objSolicSRV.iFilialEmpresa, objItensSolicSrv.lGarantia)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183664)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function SolicitacaoSRV_Insere_BD(ByVal objSolicSRV As ClassSolicSRV) As Long
'Insere nova Solicitacao de Servico

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lNumIntDocAux As Long
Dim iIndice As Integer
Dim objItensSolicSrv As ClassItensSolicSRV
Dim alComando(1 To 2) As Long


On Error GoTo Erro_SolicitacaoSRV_Insere_BD

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183665
    Next

    'Gera o número interno automático da Solicitacao
    lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_SOLICITACAOSRV", lNumIntDoc)
    If lErro <> SUCESSO Then gError 183666

    lErro = Comando_Executar(alComando(1), "INSERT INTO SolicitacaoSRV (FilialEmpresa, Codigo, Data, Hora, Cliente, FilialCliente, NumIntDoc, Vendedor, Atendente, ClienteBenef, FilialClienteBenef, Prazo, PrazoTipo, DataEntrega, Obs, Tipo, Fase) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo, objSolicSRV.dtData, CDbl(objSolicSRV.dtHora), objSolicSRV.lCliente, objSolicSRV.iFilial, lNumIntDoc, objSolicSRV.iVendedor, objSolicSRV.iAtendente, objSolicSRV.lClienteBenef, objSolicSRV.iFilialClienteBenef, objSolicSRV.iPrazo, objSolicSRV.iPrazoTipo, objSolicSRV.dtDataEntrega, objSolicSRV.sObs, objSolicSRV.lTipo, objSolicSRV.lFase)
    If lErro <> AD_SQL_SUCESSO Then gError 183667
    
    objSolicSRV.lNumIntDoc = lNumIntDoc

    'Gera o número interno automático da Solicitacao
    lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_ITENSSOLICSRV", lNumIntDocAux, objSolicSRV.colItens.Count)
    If lErro <> SUCESSO Then gError 183668

    For Each objItensSolicSrv In objSolicSRV.colItens
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO ItensSolicSRV (NumIntDoc, NumIntSolicSRV, Produto, DataVenda, Servico, UM, Quantidade, Lote, FilialOP, Solicitacao, Garantia, Contrato, StatusItem, Reparo, DataBaixa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDocAux, lNumIntDoc, objItensSolicSrv.sProduto, objItensSolicSrv.dtDataVenda, objItensSolicSrv.sServico, objItensSolicSrv.sUM, objItensSolicSrv.dQuantidade, objItensSolicSrv.sLote, objItensSolicSrv.iFilialOP, objItensSolicSrv.sSolicitacao, objItensSolicSrv.lGarantia, objItensSolicSrv.sContrato, objItensSolicSrv.iStatusItem, objItensSolicSrv.sReparo, objItensSolicSrv.dtDataBaixa)
        If lErro <> AD_SQL_SUCESSO Then gError 183669
        
        lNumIntDocAux = lNumIntDocAux + 1
            
    Next

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    SolicitacaoSRV_Insere_BD = SUCESSO

    Exit Function

Erro_SolicitacaoSRV_Insere_BD:

    SolicitacaoSRV_Insere_BD = gErr

    Select Case gErr

        Case 183665
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183666, 183668

        Case 183667
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SOLICITACAOSRV", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)

        Case 183669
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSSOLICSRV", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo, objItensSolicSrv.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183670)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function SolicitacaoSRV_Altera_BD(ByVal lComando As Long, ByVal objSolicSRV As ClassSolicSRV) As Long
'Altera uma Solicitacao de Servico

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lNumIntDocAux As Long
Dim iIndice As Integer
Dim objItensSolicSrv As ClassItensSolicSRV
Dim alComando(1 To 5) As Long
Dim objSolicSRV_BD As New ClassSolicSRV


On Error GoTo Erro_SolicitacaoSRV_Altera_BD

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183693
    Next

    lErro = Comando_ExecutarPos(alComando(1), "UPDATE SolicitacaoSRV SET Data = ?, Hora = ?, Cliente = ?, FilialCliente = ?, Vendedor = ?, Atendente = ?, ClienteBenef = ?, FilialClienteBenef = ?, Prazo = ?, PrazoTipo = ?, DataEntrega = ?, Obs = ?, Tipo = ?, Fase = ?", lComando, objSolicSRV.dtData, CDbl(objSolicSRV.dtHora), objSolicSRV.lCliente, objSolicSRV.iFilial, objSolicSRV.iVendedor, objSolicSRV.iAtendente, objSolicSRV.lClienteBenef, objSolicSRV.iFilialClienteBenef, objSolicSRV.iPrazo, objSolicSRV.iPrazoTipo, objSolicSRV.dtDataEntrega, objSolicSRV.sObs, objSolicSRV.lTipo, objSolicSRV.lFase)
    If lErro <> AD_SQL_SUCESSO Then gError 183694

    objSolicSRV_BD.lNumIntDoc = objSolicSRV.lNumIntDoc

    lErro = CF("ItensSolicSRV_Le", objSolicSRV_BD)
    If lErro <> SUCESSO Then gError 183695

    'Remove de colItens do BD os itens presente na colItens da tela
    For Each objItensSolicSrv In objSolicSRV.colItens
        For iIndice = objSolicSRV_BD.colItens.Count To 1 Step -1
            If objItensSolicSrv.lNumIntDoc = objSolicSRV_BD.colItens(iIndice).lNumIntDoc Then
                objSolicSRV_BD.colItens.Remove iIndice
                Exit For
            End If
        Next
    Next

    For Each objItensSolicSrv In objSolicSRV.colItens
    
        'se o item ja esta cadastrado
        If objItensSolicSrv.lNumIntDoc <> 0 Then
    
            lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM ItensSolicSRV WHERE NumIntDoc = ?", 0, lNumIntDoc, objItensSolicSrv.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 183696
    
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183697
    
            If lErro = AD_SQL_SEM_DADOS Then gError 183698
            
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensSolicSRV SET Solicitacao = ?, StatusItem = ?, Reparo = ?, DataBaixa = ? ", alComando(2), objItensSolicSrv.sSolicitacao, objItensSolicSrv.iStatusItem, objItensSolicSrv.sReparo, objItensSolicSrv.dtDataBaixa)
            If lErro <> AD_SQL_SUCESSO Then gError 183699
            
        Else
    
            'Gera o número interno automático da Solicitacao
            lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_ITENSSOLICSRV", lNumIntDocAux)
            If lErro <> SUCESSO Then gError 183700
        
            lErro = Comando_Executar(alComando(4), "INSERT INTO ItensSolicSRV (NumIntDoc, NumIntSolicSRV, Produto, DataVenda, Servico, UM, Quantidade, Lote, FilialOP, Solicitacao, Garantia, Contrato, StatusItem, Reparo, DataBaixa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDocAux, objSolicSRV.lNumIntDoc, objItensSolicSrv.sProduto, objItensSolicSrv.dtDataVenda, objItensSolicSrv.sServico, objItensSolicSrv.sUM, objItensSolicSrv.dQuantidade, objItensSolicSrv.sLote, objItensSolicSrv.iFilialOP, objItensSolicSrv.sSolicitacao, objItensSolicSrv.lGarantia, objItensSolicSrv.sContrato, objItensSolicSrv.iStatusItem, objItensSolicSrv.sReparo, objItensSolicSrv.dtDataBaixa)
            If lErro <> AD_SQL_SUCESSO Then gError 183701
            
        End If
            

    Next

    'Exclui os ítens de Solicitacao que foram excluídos na tela
    For Each objItensSolicSrv In objSolicSRV_BD.colItens

        lErro = Comando_Executar(alComando(5), "DELETE FROM ItensSolicSRV WHERE NumIntDoc = ?", objItensSolicSrv.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 183702
    
    Next
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    SolicitacaoSRV_Altera_BD = SUCESSO

    Exit Function

Erro_SolicitacaoSRV_Altera_BD:

    SolicitacaoSRV_Altera_BD = gErr

    Select Case gErr

        Case 183693
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183694
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SOLICITACAOSRV", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)

        Case 183695, 183700

        Case 183696, 183697
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSSOLICITACAOSRV1", gErr, objItensSolicSrv.lNumIntDoc)

        Case 183698
            Call Rotina_Erro(vbOKOnly, "ERRO_ITENSSOLICSRV_INEXISTENTE", gErr, objItensSolicSrv.lNumIntDoc, objItensSolicSrv.sProduto, objItensSolicSrv.sServico)

        Case 183699
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITENSSOLICSRV", gErr, objItensSolicSrv.lNumIntDoc)
        
        Case 183701
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSSOLICSRV", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo, objItensSolicSrv.sProduto)
        
        Case 183702
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSSOLICSRV", gErr, objItensSolicSrv.lNumIntDoc, objItensSolicSrv.sProduto, objItensSolicSrv.sServico)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183703)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function SolicitacaoSRV_Exclui(ByVal objSolicSRV As ClassSolicSRV) As Long
'Exclui da tabela SolicitacaoSRV uma solicitacao

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 4) As Long
Dim tSolicSRV As typeSolicSRV
Dim iIndice As Integer
Dim lNumIntDoc As Long, lCodRelac As Long

On Error GoTo Erro_SolicitacaoSRV_Exclui
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 183720

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183721
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM SolicitacaoSRV WHERE FilialEmpresa = ? AND Codigo = ?", 0, lNumIntDoc, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 183722
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183723
        
    If lErro = AD_SQL_SEM_DADOS Then gError 183724

    lErro = Comando_Executar(alComando(4), "SELECT Codigo FROM RelacionamentoClientes WHERE TipoDoc = ? AND NumIntDocOrigem = ?", lCodRelac, RELACCLI_TIPODOC_SOLSRV, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 211223
        
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211224
        
    If lErro <> AD_SQL_SEM_DADOS Then gError 211225

    lErro = Comando_Executar(alComando(2), "DELETE FROM ItensSolicSRV WHERE NumIntSolicSRV=?", lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 183725

    lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM SolicitacaoSRV", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 183726

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 183727

    SolicitacaoSRV_Exclui = SUCESSO
    
    Exit Function
    
Erro_SolicitacaoSRV_Exclui:
    
    SolicitacaoSRV_Exclui = gErr
    
    Select Case gErr
    
        Case 183720
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 183721
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183722, 183723
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSRV", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
        
        Case 183724
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAOSRV_NAO_CADASTRADA", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
        
        Case 183725
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSSOLICSRV1", gErr, lNumIntDoc)
        
        Case 183726
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SOLICITACAOSRV", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)

        Case 183727
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 211223, 211224
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
                
        Case 211225
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SOLICITACAOSRV_RELACCLI", gErr, lCodRelac)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183728)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function Garantia_Grava(ByVal objGarantia As ClassGarantia) As Long
'grava na tabela Garantia

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 1) As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer

On Error GoTo Erro_Garantia_Grava
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 183935

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183926
    Next

    'Faz o lock nos objetos vinculados a Garantia
    lErro = CF("GarantiaObjVinc_Lock", objGarantia)
    If lErro <> SUCESSO Then gError 183927

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM Garantia WHERE FilialEmpresa = ? AND Codigo = ?", 0, lNumIntDoc, objGarantia.iFilialEmpresa, objGarantia.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 183928
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183929
        
    If lErro = AD_SQL_SEM_DADOS Then

        'Insere a garantia nova
        lErro = Garantia_Insere_BD(objGarantia)
        If lErro <> SUCESSO Then gError 183930

    Else
    
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 183931
    
        objGarantia.lNumIntDoc = lNumIntDoc
    
        'Altera a garantia
        lErro = Garantia_Altera_BD(objGarantia, alComando(1))
        If lErro <> SUCESSO Then gError 183932
    
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 183933

    lErro = CF("WorkFlow_Trata_Transacao", "SRV", objGarantia.objTela, gsUsuario, objGarantia)
    If lErro <> SUCESSO Then gError 183934
          
    Garantia_Grava = SUCESSO
    
    Exit Function
    
Erro_Garantia_Grava:
    
    Garantia_Grava = gErr
    
    Select Case gErr
    
        Case 183926
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183927, 183930, 183932, 183934
        
        Case 183928, 183929
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIA", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo)
        
        Case 183931
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_GARANTIA", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo)
        
        Case 183933
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 183935
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183936)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function GarantiaObjVinc_Lock(ByVal objGarantia As ClassGarantia) As Long
'Faz "lock" nos objetos associados a Garantia

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim objProduto As New ClassProduto
Dim objRastroLote As New ClassRastreamentoLote
Dim iIndice As Integer
Dim objGarantiaNumSerie As ClassGarantiaNumSerie
Dim objGarantiaProduto As ClassGarantiaProduto

On Error GoTo Erro_GarantiaObjVinc_Lock

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183916
    Next

    objProduto.sCodigo = objGarantia.sProduto
    
    'Locka  o produto
    lErro = CF("Produto_Lock_Shared", alComando(1), objProduto)
    If lErro <> SUCESSO Then gError 183917
    
    If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 183918

    If gobjSRV.iVerificaLote = VERIFICA_LOTE Then
    
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_OP Then

            objRastroLote.sProduto = objGarantia.sProduto
            objRastroLote.sCodigo = objGarantia.sLote
            objRastroLote.iFilialOP = objGarantia.iFilialOP
    
            'Lê rastreamento do lote a partir do produto, filialOP e código do lote passados
            lErro = CF("RastreamentoLote_Le1", alComando(2), objRastroLote)
            If lErro <> SUCESSO And lErro <> 71889 Then gError 183919
            
            If lErro <> SUCESSO Then gError 183920

        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then

            objRastroLote.sProduto = objGarantia.sProduto

            For Each objGarantiaNumSerie In objGarantia.colNumSerie
            
                objRastroLote.sCodigo = objGarantiaNumSerie.sNumSerie

                'Lê rastreamento do lote a partir do produto, filialOP e código do lote passados
                lErro = CF("RastreamentoLote_Le1", alComando(3), objRastroLote)
                If lErro <> SUCESSO And lErro <> 71889 Then gError 183921
                
                If lErro <> SUCESSO Then gError 183922

            Next

        End If

    End If
    
    'Para cada item de Pedido Venda
    For Each objGarantiaProduto In objGarantia.colProduto

        objProduto.sCodigo = objGarantiaProduto.sProduto
        
        'Locka  o produto
        lErro = CF("Produto_Lock_Shared", alComando(4), objProduto)
        If lErro <> SUCESSO Then gError 183923

        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 183924

    Next
    

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    GarantiaObjVinc_Lock = SUCESSO

    Exit Function

Erro_GarantiaObjVinc_Lock:

    GarantiaObjVinc_Lock = gErr

    Select Case gErr

        Case 183916
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183917, 183919, 183921, 183923

        Case 183918, 183924
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, objProduto.sCodigo)

        Case 183920
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_CADASTRADO_RASTREAMENTO", gErr, objGarantia.sLote, objGarantia.sProduto, objGarantia.iFilialOP)

        Case 183922
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMSERIE_NAO_CADASTRADO", gErr, objGarantia.sProduto, objGarantia.sLote)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183925)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Garantia_Insere_BD(ByVal objGarantia As ClassGarantia) As Long
'Insere nova Garantia

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim alComando(1 To 3) As Long
Dim objGarantiaProduto As ClassGarantiaProduto
Dim objGarantiaNumSerie As ClassGarantiaNumSerie

On Error GoTo Erro_Garantia_Insere_BD

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183937
    Next

    'Gera o número interno automático da Solicitacao
    lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_GARANTIA", lNumIntDoc)
    If lErro <> SUCESSO Then gError 183938

    lErro = Comando_Executar(alComando(1), "INSERT INTO Garantia (NumIntDoc, Codigo, FilialEmpresa, Produto, DataVenda, Quantidade, Lote, FilialOP, Serie, NumNotaFiscal, TipoGarantia, GarantiaTotal, GarantiaTotalPrazo, Ativo, Fornecedor,FilialFornecedor,CliFabr, FilialCliFabr) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDoc, objGarantia.lCodigo, objGarantia.iFilialEmpresa, objGarantia.sProduto, objGarantia.dtDataVenda, objGarantia.dQuantidade, objGarantia.sLote, objGarantia.iFilialOP, objGarantia.sSerie, objGarantia.lNumNotaFiscal, objGarantia.lTipoGarantia, objGarantia.iGarantiaTotal, objGarantia.iGarantiaTotalPrazo, objGarantia.iAtivo, objGarantia.lFornecedor, objGarantia.iFilialFornecedor, objGarantia.lCliFabr, objGarantia.iFilialCliFabr)
    If lErro <> AD_SQL_SUCESSO Then gError 183939

    For Each objGarantiaProduto In objGarantia.colProduto

        lErro = Comando_Executar(alComando(2), "INSERT INTO GarantiaProduto (NumIntGarantia, Produto, Prazo) VALUES (?,?,?)", lNumIntDoc, objGarantiaProduto.sProduto, objGarantiaProduto.iPrazo)
        If lErro <> AD_SQL_SUCESSO Then gError 183940

    Next
    
    For Each objGarantiaNumSerie In objGarantia.colNumSerie

        lErro = Comando_Executar(alComando(3), "INSERT INTO GarantiaNumSerie (NumIntGarantia, NumSerie) VALUES (?,?)", lNumIntDoc, objGarantiaNumSerie.sNumSerie)
        If lErro <> AD_SQL_SUCESSO Then gError 183941

    Next
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Garantia_Insere_BD = SUCESSO

    Exit Function

Erro_Garantia_Insere_BD:

    Garantia_Insere_BD = gErr

    Select Case gErr

        Case 183937
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183938

        Case 183939
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GARANTIA", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo)

        Case 183940
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GARANTIAPRODUTO", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo, objGarantiaProduto.sProduto)
        
        Case 183941
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GARANTIANUMSERIE", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo, objGarantiaNumSerie.sNumSerie)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183942)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Garantia_Altera_BD(objGarantia As ClassGarantia, ByVal lComando As Long) As Long
'Altera uma Solicitacao de Servico

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lNumIntDocAux As Long
Dim iIndice As Integer
Dim objGarantiaProduto As ClassGarantiaProduto
Dim alComando(1 To 5) As Long
Dim objGarantia_BD As New ClassGarantia
Dim objSolicSRV As New ClassSolicSRV
Dim objProduto As New ClassProduto
Dim objItensSolicSrv As ClassItensSolicSRV
Dim objGarantiaNumSerie_BD As ClassGarantiaNumSerie
Dim objGarantiaNumSerie As ClassGarantiaNumSerie
Dim iAchou As Integer
Dim objGarantiaProduto_BD As ClassGarantiaProduto
Dim sNumSerie As String
Dim iPrazo As Integer

On Error GoTo Erro_Garantia_Altera_BD

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183949
    Next

    objGarantia_BD.iFilialEmpresa = objGarantia.iFilialEmpresa
    objGarantia_BD.lCodigo = objGarantia.lCodigo

    lErro = CF("Garantia_Le", objGarantia_BD)
    If lErro <> SUCESSO And lErro <> 183568 Then gError 183950

    'se encontrou alguma solicitacao envolvendo a garantia e produto em questao
    lErro = CF("ItensSolicSRV_Le_Garantia", objSolicSRV, objGarantia_BD.lNumIntDoc, objGarantia_BD.sProduto)
    If lErro <> SUCESSO And lErro <> 183946 Then gError 183951
    
    If lErro = SUCESSO Then
    
        'se o produto foi alterado ==> erro (nao pode alterar)
        If objGarantia_BD.sProduto <> objGarantia.sProduto Then gError 183952
    
        objProduto.sCodigo = objGarantia.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 183953
            
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_OP Then
        
            For Each objItensSolicSrv In objSolicSRV.colItens
            
                If objItensSolicSrv.sLote = objGarantia_BD.sLote And objItensSolicSrv.iFilialOP = objGarantia_BD.iFilialOP Then
                
                    If objGarantia.sLote <> objGarantia_BD.sLote Or objGarantia.iFilialOP <> objGarantia_BD.iFilialOP Then gError 183954
                    
                End If
                
            Next
        
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then

            'pesquisa as solicitacoes de servico
            For Each objItensSolicSrv In objSolicSRV.colItens
            
                'pesquisa os numeros de serie que estao no BD
                For Each objGarantiaNumSerie_BD In objGarantia_BD.colNumSerie
            
                    'se algum num. serie participa de uma solicitacao ==> nao pode ser excluido
                    If objItensSolicSrv.sLote = objGarantiaNumSerie_BD.sNumSerie Then
                    
                        iAchou = 0
                    
                        For Each objGarantiaNumSerie In objGarantia.colNumSerie
                        
                            If objGarantiaNumSerie.sNumSerie = objGarantiaNumSerie_BD.sNumSerie Then
                                iAchou = 1
                                Exit For
                            End If
                            
                        Next
                        
                        'se nao achou o num. serie ==> que foi excluido mas nao podia pois é usado em SolicSRV ==> erro
                        If iAchou = 0 Then gError 183955
                            
                    End If
                
                Next
                
            Next

        End If

    End If

    lErro = Comando_ExecutarPos(alComando(1), "UPDATE Garantia SET Produto = ?, DataVenda = ?, Quantidade = ?, Lote = ?, FilialOP = ?, Serie = ?, NumNotaFiscal = ?, TipoGarantia = ?, GarantiaTotal = ?, GarantiaTotalPrazo = ?, Ativo = ?, Fornecedor = ?, FilialFornecedor = ?,CliFabr = ?, FilialCliFabr = ?", lComando, _
    objGarantia.sProduto, objGarantia.dtDataVenda, objGarantia.dQuantidade, objGarantia.sLote, objGarantia.iFilialOP, objGarantia.sSerie, objGarantia.lNumNotaFiscal, objGarantia.lTipoGarantia, objGarantia.iGarantiaTotal, objGarantia.iGarantiaTotalPrazo, objGarantia.iAtivo, objGarantia.lFornecedor, objGarantia.iFilialFornecedor, objGarantia.lCliFabr, objGarantia.iFilialCliFabr)
    If lErro <> AD_SQL_SUCESSO Then gError 183956

    'Remove de colProduto do BD os itens presente na colProduto da tela
    For Each objGarantiaProduto In objGarantia.colProduto
        For iIndice = objGarantia_BD.colProduto.Count To 1 Step -1
            If objGarantiaProduto.sProduto = objGarantia_BD.colProduto(iIndice).sProduto Then
                objGarantia_BD.colProduto.Remove iIndice
                Exit For
            End If
        Next
    Next

    'Remove de colNumSerie do BD os itens presente na colNumSerie da tela
    For Each objGarantiaNumSerie In objGarantia.colNumSerie
        For iIndice = objGarantia_BD.colNumSerie.Count To 1 Step -1
            If objGarantiaNumSerie.sNumSerie = objGarantia_BD.colNumSerie(iIndice).sNumSerie Then
                objGarantia_BD.colNumSerie.Remove iIndice
                Exit For
            End If
        Next
    Next

    For Each objGarantiaProduto In objGarantia.colProduto

        lErro = Comando_ExecutarPos(alComando(2), "SELECT Prazo FROM GarantiaProduto WHERE NumIntGarantia = ? AND Produto = ?", 0, iPrazo, objGarantia.lNumIntDoc, objGarantiaProduto.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 183957

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183958

        If lErro = AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE GarantiaProduto SET Prazo = ?", alComando(2), objGarantiaProduto.iPrazo)
            If lErro <> AD_SQL_SUCESSO Then gError 183959
            
        Else
    
            lErro = Comando_Executar(alComando(4), "INSERT INTO GarantiaProduto (NumIntGarantia, Produto, Prazo) VALUES (?,?,?)", objGarantia.lNumIntDoc, objGarantiaProduto.sProduto, objGarantiaProduto.iPrazo)
            If lErro <> AD_SQL_SUCESSO Then gError 183960
            
        End If

    Next

    'Exclui os ítens de Solicitacao que foram excluídos na tela
    For Each objGarantiaProduto_BD In objGarantia_BD.colProduto

        lErro = Comando_Executar(alComando(5), "DELETE FROM GarantiaProduto WHERE NumIntGarantia = ? AND Produto = ?", objGarantiaProduto_BD.lNumIntGarantia, objGarantiaProduto_BD.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 183961
    
    Next
    
    For Each objGarantiaNumSerie In objGarantia.colNumSerie

        sNumSerie = String(STRING_LOTE_RASTREAMENTO, 0)

        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumSerie FROM GarantiaNumSerie WHERE NumIntGarantia = ? AND NumSerie = ?", 0, sNumSerie, objGarantia.lNumIntDoc, objGarantiaNumSerie.sNumSerie)
        If lErro <> AD_SQL_SUCESSO Then gError 183962

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183963

        If lErro = AD_SQL_SEM_DADOS Then
        
            lErro = Comando_Executar(alComando(4), "INSERT INTO GarantiaNumSerie (NumIntGarantia, NumSerie) VALUES (?,?)", objGarantia.lNumIntDoc, objGarantiaNumSerie.sNumSerie)
            If lErro <> AD_SQL_SUCESSO Then gError 183964
            
        End If
            
    Next

    'Exclui os ítens de GarantiaNumSerie que foram excluídos na tela
    For Each objGarantiaNumSerie_BD In objGarantia_BD.colNumSerie

        lErro = Comando_Executar(alComando(5), "DELETE FROM GarantiaNumSerie WHERE NumIntGarantia = ? AND NumSerie = ?", objGarantiaNumSerie_BD.lNumIntGarantia, objGarantiaNumSerie_BD.sNumSerie)
        If lErro <> AD_SQL_SUCESSO Then gError 183966
    
    Next
    
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Garantia_Altera_BD = SUCESSO

    Exit Function

Erro_Garantia_Altera_BD:

    Garantia_Altera_BD = gErr

    Select Case gErr

        Case 183949
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183950, 183951, 183953
        
        Case 183952
            Call Rotina_Erro(vbOKOnly, "ERRO_PROD_EM_SOLICSRV_GARANTIA", gErr, objGarantia_BD.sProduto)

        Case 183954
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_EM_SOLICSRV_GARANTIA", gErr, objGarantia_BD.sProduto, objGarantia_BD.sLote, objGarantia_BD.iFilialOP)

        Case 183955
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMSERIE_EM_SOLICSRV_GARANTIA", gErr, objGarantia_BD.sProduto, objGarantiaNumSerie_BD.sNumSerie)

        Case 183956
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_GARANTIA", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo)
        
        Case 183957, 183958
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIAPRODUTO1", gErr, objGarantia.lNumIntDoc, objGarantiaProduto.sProduto)

        Case 183959
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_GARANTIAPRODUTO", gErr, objGarantia.lNumIntDoc, objGarantiaProduto.sProduto)

        Case 183960
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GARANTIAPRODUTO", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo, objGarantiaProduto.sProduto)

        Case 183961
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GARANTIAPRODUTO", gErr, objGarantiaProduto_BD.lNumIntGarantia, objGarantiaProduto_BD.sProduto)

        Case 183962, 183963
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIANUMSERIE", gErr, objGarantia.lNumIntDoc)

        Case 183964
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GARANTIANUMSERIE", gErr, objGarantia.lNumIntDoc, objGarantiaNumSerie.sNumSerie)

        Case 183966
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GARANTIANUMSERIE", gErr, objGarantiaNumSerie_BD.lNumIntGarantia, objGarantiaNumSerie_BD.sNumSerie)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183967)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Garantia_Exclui(ByVal objGarantia As ClassGarantia) As Long
'Exclui uma garantia

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objGarantia_BD As New ClassGarantia
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_Garantia_Exclui
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 183973

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183974
    Next

    objGarantia_BD.iFilialEmpresa = objGarantia.iFilialEmpresa
    objGarantia_BD.lCodigo = objGarantia.lCodigo

    lErro = CF("Garantia_Le", objGarantia_BD)
    If lErro <> SUCESSO And lErro <> 183568 Then gError 183975

    'se encontrou alguma solicitacao envolvendo a garantia e produto em questao
    lErro = CF("ItensSolicSRV_Le_Garantia", objSolicSRV, objGarantia_BD.lNumIntDoc, objGarantia_BD.sProduto)
    If lErro <> SUCESSO And lErro <> 183946 Then gError 183976

    If lErro = SUCESSO Then gError 183977

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM Garantia WHERE FilialEmpresa = ? AND Codigo = ?", 0, lNumIntDoc, objGarantia.iFilialEmpresa, objGarantia.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 183978
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183979
        
    If lErro = AD_SQL_SEM_DADOS Then gError 183980

    lErro = Comando_Executar(alComando(2), "DELETE FROM GarantiaProduto WHERE NumIntGarantia=?", lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 183981

    lErro = Comando_Executar(alComando(3), "DELETE FROM GarantiaNumSerie WHERE NumIntGarantia=?", lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 183982

    lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM Garantia", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 183983

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 183984

    Garantia_Exclui = SUCESSO
    
    Exit Function
    
Erro_Garantia_Exclui:
    
    Garantia_Exclui = gErr
    
    Select Case gErr
    
        Case 183973
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 183974
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183975, 183976

        Case 183977
            Call Rotina_Erro(vbOKOnly, "ERRO_GARANTIA_EM_SOLICITACAOSRV", gErr)

        Case 183978, 183979
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GARANTIA", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo)
        
        Case 183980
            Call Rotina_Erro(vbOKOnly, "ERRO_GARANTIA_NAO_CADASTRADA", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo)
        
        Case 183981
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GARANTIAPRODUTO1", gErr, lNumIntDoc)
        
        Case 183982
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GARANTIANUMSERIE1", gErr, lNumIntDoc)
        
        Case 183983
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GARANTIA", gErr, objGarantia.iFilialEmpresa, objGarantia.lCodigo)

        Case 183984
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183985)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function TipoGarantia_Exclui(ByVal objTipoGarantia As ClassTipoGarantia) As Long
'Exclui um tipo de garantia

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_TipoGarantia_Exclui
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 186044

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186045
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM TipoGarantia WHERE Codigo = ?", 0, lNumIntDoc, objTipoGarantia.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 186046
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186047
        
    If lErro = AD_SQL_SEM_DADOS Then gError 186048

    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM TipoGarantia", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 186049

    lErro = Comando_Executar(alComando(3), "DELETE FROM TipoGarantiaProduto WHERE NumIntTipoGarantia=?", lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 186113
    
    lErro = CF("Exclusao_Valida", VALIDAEXCLUSAO_CODIGO_TIPOGARANTIA, objTipoGarantia.lCodigo)
    If lErro <> SUCESSO Then gError 198292

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 186050

    TipoGarantia_Exclui = SUCESSO
    
    Exit Function
    
Erro_TipoGarantia_Exclui:
    
    TipoGarantia_Exclui = gErr
    
    Select Case gErr
    
        Case 186044
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 186045
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186046, 186047
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOGARANTIA", gErr, objTipoGarantia.lCodigo)
        
        Case 186048
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOGARANTIA_NAO_CADASTRADA", gErr, objTipoGarantia.lCodigo)
        
        Case 186049
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOGARANTIA", gErr, objTipoGarantia.lCodigo)

        Case 186050
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 186113
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOGARANTIAPRODUTO", gErr, objTipoGarantia.lCodigo)
        
        Case 198292
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186051)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function TipoGarantia_Grava(ByVal objTipoGarantia As ClassTipoGarantia) As Long
'grava na tabela Garantia

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 1) As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer

On Error GoTo Erro_TipoGarantia_Grava
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 186107

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186097
    Next

    'Faz o lock nos objetos vinculados a TipoGarantia
    lErro = CF("TipoGarantiaObjVinc_Lock", objTipoGarantia)
    If lErro <> SUCESSO Then gError 186098

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM TipoGarantia WHERE Codigo = ?", 0, lNumIntDoc, objTipoGarantia.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 186099
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186100
        
    If lErro = AD_SQL_SEM_DADOS Then

        'Insere tipogarantia novo
        lErro = TipoGarantia_Insere_BD(objTipoGarantia)
        If lErro <> SUCESSO Then gError 186101

    Else
    
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 186102
    
        objTipoGarantia.lNumIntDoc = lNumIntDoc
    
        'Altera a garantia
        lErro = TipoGarantia_Altera_BD(objTipoGarantia, alComando(1))
        If lErro <> SUCESSO Then gError 186103
    
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 186104

    lErro = CF("WorkFlow_Trata_Transacao", "SRV", objTipoGarantia.objTela, gsUsuario, objTipoGarantia)
    If lErro <> SUCESSO Then gError 186105
          
    TipoGarantia_Grava = SUCESSO
    
    Exit Function
    
Erro_TipoGarantia_Grava:
    
    TipoGarantia_Grava = gErr
    
    Select Case gErr
    
        Case 186097
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186098, 186101, 186103, 186105
        
        Case 186099, 186100
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOGARANTIA", gErr, objTipoGarantia.lCodigo)
        
        Case 186102
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOGARANTIA", gErr, objTipoGarantia.lCodigo)
        
        Case 186104
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 186107
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186106)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function TipoGarantiaObjVinc_Lock(ByVal objTipoGarantia As ClassTipoGarantia) As Long
'Faz "lock" nos objetos associados a TipoGarantia

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim objTipoGarantiaProduto As ClassTipoGarantiaProduto

On Error GoTo Erro_TipoGarantiaObjVinc_Lock

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186093
    Next
    
    'Para cada item do Tipo Garantia
    For Each objTipoGarantiaProduto In objTipoGarantia.colTipoGarantiaProduto

        objProduto.sCodigo = objTipoGarantiaProduto.sProduto
        
        'Locka  o produto
        lErro = CF("Produto_Lock_Shared", alComando(1), objProduto)
        If lErro <> SUCESSO Then gError 186094
        
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 186095

    Next
    

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TipoGarantiaObjVinc_Lock = SUCESSO

    Exit Function

Erro_TipoGarantiaObjVinc_Lock:

    TipoGarantiaObjVinc_Lock = gErr

    Select Case gErr

        Case 186093
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186094

        Case 186095
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186096)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function TipoGarantia_Insere_BD(ByVal objTipoGarantia As ClassTipoGarantia) As Long
'Insere nova Garantia

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim alComando(1 To 3) As Long
Dim objTipoGarantiaProduto As ClassTipoGarantiaProduto

On Error GoTo Erro_TipoGarantia_Insere_BD

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186108
    Next

    'Gera o número interno automático da Solicitacao
    lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_TIPOGARANTIA", lNumIntDoc)
    If lErro <> SUCESSO Then gError 186109

    lErro = Comando_Executar(alComando(1), "INSERT INTO TipoGarantia (NumIntDoc, Codigo, Descricao, PrazoPadrao, GarantiaTotal, GarantiaTotalPrazo) VALUES (?,?,?,?,?,?)", lNumIntDoc, objTipoGarantia.lCodigo, objTipoGarantia.sDescricao, objTipoGarantia.iPrazoPadrao, objTipoGarantia.iGarantiaTotal, objTipoGarantia.iGarantiaTotalPrazo)
    If lErro <> AD_SQL_SUCESSO Then gError 186110

    For Each objTipoGarantiaProduto In objTipoGarantia.colTipoGarantiaProduto

        lErro = Comando_Executar(alComando(2), "INSERT INTO TipoGarantiaProduto (NumIntTipoGarantia, Produto, Prazo) VALUES (?,?,?)", lNumIntDoc, objTipoGarantiaProduto.sProduto, objTipoGarantiaProduto.iPrazo)
        If lErro <> AD_SQL_SUCESSO Then gError 186111

    Next
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TipoGarantia_Insere_BD = SUCESSO

    Exit Function

Erro_TipoGarantia_Insere_BD:

    TipoGarantia_Insere_BD = gErr

    Select Case gErr

        Case 186108
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186109

        Case 186110
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOGARANTIA", gErr, objTipoGarantia.lCodigo)

        Case 186111
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOGARANTIAPRODUTO", gErr, objTipoGarantia.lCodigo, objTipoGarantiaProduto.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186112)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function TipoGarantia_Altera_BD(objTipoGarantia As ClassTipoGarantia, ByVal lComando As Long) As Long
'Altera uma Solicitacao de Servico

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoGarantiaProduto As ClassTipoGarantiaProduto
Dim alComando(1 To 2) As Long
Dim iAchou As Integer
Dim iPrazo As Integer

On Error GoTo Erro_TipoGarantia_Altera_BD

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186114
    Next

    lErro = Comando_ExecutarPos(alComando(1), "UPDATE TipoGarantia SET Descricao = ?, PrazoPadrao = ?, GarantiaTotal = ?, GarantiaTotalPrazo = ?", lComando, _
    objTipoGarantia.sDescricao, objTipoGarantia.iPrazoPadrao, objTipoGarantia.iGarantiaTotal, objTipoGarantia.iGarantiaTotalPrazo)
    If lErro <> AD_SQL_SUCESSO Then gError 186115

    lErro = Comando_Executar(alComando(2), "DELETE FROM TipoGarantiaProduto WHERE NumIntTipoGarantia=?", objTipoGarantia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 186116

    For Each objTipoGarantiaProduto In objTipoGarantia.colTipoGarantiaProduto

        lErro = Comando_Executar(alComando(2), "INSERT INTO TipoGarantiaProduto (NumIntTipoGarantia, Produto, Prazo) VALUES (?,?,?)", objTipoGarantia.lNumIntDoc, objTipoGarantiaProduto.sProduto, objTipoGarantiaProduto.iPrazo)
        If lErro <> AD_SQL_SUCESSO Then gError 186117

    Next


    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TipoGarantia_Altera_BD = SUCESSO

    Exit Function

Erro_TipoGarantia_Altera_BD:

    TipoGarantia_Altera_BD = gErr

    Select Case gErr

        Case 186114
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186115
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_TIPOGARANTIA", gErr, objTipoGarantia.lCodigo)
        
        Case 186116
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOGARANTIAPRODUTO", gErr, objTipoGarantia.lCodigo)

        Case 186117
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOGARANTIAPRODUTO", gErr, objTipoGarantia.lCodigo, objTipoGarantiaProduto.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186118)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function OrcamentoSRV_Grava(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Grava o Orcamento de SRV

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_OrcamentoSRV_Grava

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 186792

    'Chama a rotina de gravação em transação
    lErro = CF("OrcamentoSRV_Grava_EmTrans", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186793
    
    'Faz o commit na transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 186794

    lErro = CF("WorkFlow_Trata_Transacao", "SRV", objOrcamentoVenda.objTela, gsUsuario)
    If lErro <> SUCESSO Then gError 186795
    
    OrcamentoSRV_Grava = SUCESSO

    Exit Function

Erro_OrcamentoSRV_Grava:

    OrcamentoSRV_Grava = gErr

    Select Case gErr

        Case 186792
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 186793, 186795
        
        Case 186794
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186796)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Public Function OrcamentoSRV_Grava_EmTrans(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Grava o Orcamento de Venda - deve ser chamada dentro de uma Transação

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0) As Long
Dim objOrcamentoVenda_BD As New ClassOrcamentoVenda
Dim lNumIntPedVenda As Long
Dim lNumIntNFiscal As Long

On Error GoTo Erro_OrcamentoSRV_Grava_EmTrans
    
    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186885
    Next

    'Faz o lock nos objetos vinculados ao Orcamento
    lErro = CF("OVObjetosVinculados_Lock", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186886

    'Faz o lock nos objetos vinculados ao Orcamento
    lErro = CF("OrcSRVObjetosVinculados_Lock", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186887

    'Lê OrcamentoVenda
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntPedSRV, NumIntNFiscal FROM OrcamentoSRV WHERE Codigo = ? AND FilialEmpresa = ? ", 0, lNumIntPedVenda, lNumIntNFiscal, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 186888

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 186889

    'Se não existir
    If lErro <> AD_SQL_SUCESSO Then

        'Insere o orcamento novo no Banco de Dados
        lErro = CF("OrcamentoSRV_Insere_BD", objOrcamentoVenda)
        If lErro <> SUCESSO Then gError 186890

    'Se existir
    Else
        
        'Se o orçamento estiver vinculado a um Pedido de Venda, não poderá ser alterado -> Erro.
        If lNumIntPedVenda <> 0 Then gError 186891
        
        'Se o orçamento estiver vinculado a uma Nota Fiscal, não poderá ser alterado -> Erro.
        If lNumIntNFiscal <> 0 Then gError 186892
        
        'Altera o Orcamento de Venda e armazena os ítens do Pedido no BD em objOrcamentoVenda_BD
        lErro = CF("OrcamentoSRV_Altera_BD", objOrcamentoVenda, objOrcamentoVenda_BD, alComando(0))
        If lErro <> SUCESSO Then gError 101213

    End If

'    '############################################
'    'Inserido por Wagner 22/03/2006
'    lErro = CF("ProjetoInfo_Grava", objOrcamentoVenda.objProjetoInfo, PRJ_CR_TIPO_OV, objOrcamentoVenda.lNumIntDoc)
'    If lErro <> SUCESSO Then gError 181706
'    '############################################

    objOrcamentoVenda.objInfoAdic.lNumIntDoc = objOrcamentoVenda.lNumIntDoc
    objOrcamentoVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_OSRV
    lErro = CF("InfoAdicionais_Grava", objOrcamentoVenda.objInfoAdic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrcamentoSRV_Grava_EmTrans = SUCESSO

    Exit Function

Erro_OrcamentoSRV_Grava_EmTrans:

    OrcamentoSRV_Grava_EmTrans = gErr

    Select Case gErr

        Case 186885
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186886, 186887, 186890, 101213

        Case 186888, 186889
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186891
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_VINCULADO_PEDVENDA", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
            
        Case 186892
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_VINCULADO_NFISCAL", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186893)

    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function OrcamentoSRV_Insere_BD(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Insere o novo Orcamento de Venda no Banco de Dados

Dim lErro As Long
Dim dQuantNaoCancelada As Double
Dim lNumIntDoc As Long
Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long

On Error GoTo Erro_OrcamentoSRV_Insere_BD

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186819
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 186820
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 186821
        
    'Gera o número interno automático do Orcamento de Venda
    lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_ORCAMENTO_SRV", lNumIntDoc)
    If lErro <> SUCESSO Then gError 186822

    objOrcamentoVenda.lNumIntDoc = lNumIntDoc

    'Insere o novo Orcamento SRV no BD
    lErro = Comando_Executar(lComando, "INSERT INTO OrcamentoSRV (NaturezaOp, NomeCli, NomeFilialCli, Vendedor, FilialEmpresa, Codigo, Cliente, Filial, CondicaoPagto, PercAcrescFinanceiro, DataEmissao, DataReferencia, ValorTotal, ValorProdutos, ValorFrete, ValorDesconto, ValorSeguro, ValorOutrasDespesas, TabelaPreco, CobrancaAutomatica, NumIntDoc, PrazoValidade, Usuario, CodigoBase, Status, MotivoPerda, NumIntSolicSRV, ValorDescontoTit, ValorItens, ValorDescontoItens) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objOrcamentoVenda.sNaturezaOp, objOrcamentoVenda.sNomeCli, objOrcamentoVenda.sNomeFilialCli, objOrcamentoVenda.iVendedor, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, objOrcamentoVenda.lCliente, objOrcamentoVenda.iFilial, objOrcamentoVenda.iCondicaoPagto, objOrcamentoVenda.dPercAcrescFinanceiro, objOrcamentoVenda.dtDataEmissao, objOrcamentoVenda.dtDataReferencia, objOrcamentoVenda.dValorTotal, objOrcamentoVenda.dValorProdutos, objOrcamentoVenda.dValorFrete, _
        objOrcamentoVenda.dValorDesconto, objOrcamentoVenda.dValorSeguro, objOrcamentoVenda.dValorOutrasDespesas, objOrcamentoVenda.iTabelaPreco, objOrcamentoVenda.iCobrancaAutomatica, objOrcamentoVenda.lNumIntDoc, objOrcamentoVenda.iPrazoValidade, objOrcamentoVenda.sUsuario, objOrcamentoVenda.lCodigoBase, objOrcamentoVenda.lStatus, objOrcamentoVenda.lMotivoPerda, objOrcamentoVenda.lNumIntSolicSRV, objOrcamentoVenda.dValorDescontoTit, objOrcamentoVenda.dValorItens, objOrcamentoVenda.dValorDescontoItens)
    If lErro <> AD_SQL_SUCESSO Then gError 186823

    'grava a tributacao da nota como um todo
    lErro = CF("OrcamentoSRV_GravaTributacao", objOrcamentoVenda, lComando2, lComando3)
    If lErro <> SUCESSO Then gError 186824

    'Insere as Parcelas do Pedido no BD
    lErro = CF("ParcelasOrcSRV_Grava", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186825

    'Insere os itens do Pedido no BD
    lErro = CF("ItensOrcamentoSRV_Grava", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186826
    
    lErro = CF("BloqueiosGen_Gera", MAPBLOQGEN_TIPOTELA_ORCAMENTOSRV, objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 191625

    OrcamentoSRV_Insere_BD = SUCESSO

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

Erro_OrcamentoSRV_Insere_BD:

    OrcamentoSRV_Insere_BD = gErr

    Select Case gErr

        Case 186819 To 186821
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186822, 186824, 186825, 186826, 191625

        Case 186823
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORCAMENTOSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186827)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function


Public Function OrcamentoSRV_GravaTributacao(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long

On Error GoTo Erro_OrcamentoSRV_GravaTributacao

'    With objOrcamentoVenda.objTributacaoOV
'        lErro = Comando_Executar(lComando1, "INSERT INTO TributacaoOrcSRV (FilialEmpresa, CodOrcamento, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 186803
'
'    'inserir os complementos
'    lErro = CF("OrcamentoSRV_GravaTributacaoCompl", objOrcamentoVenda, objOrcamentoVenda.objTributacaoOV.objTributacaoDesconto, lComando2)
'    If lErro <> SUCESSO Then gError 186804
'
'    lErro = CF("OrcamentoSRV_GravaTributacaoCompl", objOrcamentoVenda, objOrcamentoVenda.objTributacaoOV.objTributacaoFrete, lComando2)
'    If lErro <> SUCESSO Then gError 186805
'
'    lErro = CF("OrcamentoSRV_GravaTributacaoCompl", objOrcamentoVenda, objOrcamentoVenda.objTributacaoOV.objTributacaoOutras, lComando2)
'    If lErro <> SUCESSO Then gError 186806
'
'    lErro = CF("OrcamentoSRV_GravaTributacaoCompl", objOrcamentoVenda, objOrcamentoVenda.objTributacaoOV.objTributacaoSeguro, lComando2)
'    If lErro <> SUCESSO Then gError 186807

    objOrcamentoVenda.objTributacao.lNumIntDoc = objOrcamentoVenda.lNumIntDoc
    objOrcamentoVenda.objTributacao.iTipoDoc = TIPODOC_TRIB_OSRV

    lErro = CF("TributacaoDoc_Grava", objOrcamentoVenda.objTributacao)
    If lErro <> SUCESSO Then gError 186804

    OrcamentoSRV_GravaTributacao = SUCESSO

    Exit Function

Erro_OrcamentoSRV_GravaTributacao:

    OrcamentoSRV_GravaTributacao = gErr

    Select Case gErr

'        Case 186803
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBUTACAOORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186804 'To 186807

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186808)

    End Select

    Exit Function

End Function

'Public Function OrcamentoSRV_GravaTributacaoCompl(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal objTribComplOV As ClassTributacaoItemPV, ByVal lComando As Long) As Long
'
'Dim lErro As Long, iItem As Integer
'
'On Error GoTo Erro_OrcamentoSRV_GravaTributacaoCompl
'
'    iItem = objTribComplOV.lNumIntDoc
'
'    With objTribComplOV
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoComplOrcSRV (FilialEmpresa, CodOrcamento, Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPICodProduto, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, iItem, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, "", .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 186801
'
'    OrcamentoSRV_GravaTributacaoCompl = SUCESSO
'
'    Exit Function
'
'Erro_OrcamentoSRV_GravaTributacaoCompl:
'
'    OrcamentoSRV_GravaTributacaoCompl = gErr
'
'    Select Case gErr
'
'        Case 186801
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBCOMPLORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186802)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function ParcelasOrcSRV_Grava(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Insere as no BD Parcelas do orcamento passado

Dim lErro As Long
Dim lComando As Long
Dim objParcelaOV As ClassParcelaOV

On Error GoTo Erro_ParcelasOrcSRV_Grava

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186809
    
    'Para cada parcela da coleção
    For Each objParcelaOV In objOrcamentoVenda.colParcela

        'Insere em ParcelasOV
        lErro = Comando_Executar(lComando, "INSERT INTO ParcelasOrcSRV (FilialEmpresa,OrcamentoSRV,NumParcela,DataVencimento,Valor,Desconto1Codigo,Desconto1Ate,Desconto1Valor,Desconto2Codigo,Desconto2Ate,Desconto2Valor,Desconto3Codigo,Desconto3Ate,Desconto3Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, objParcelaOV.iNumParcela, objParcelaOV.dtDataVencimento, objParcelaOV.dValor, objParcelaOV.iDesconto1Codigo, objParcelaOV.dtDesconto1Ate, objParcelaOV.dDesconto1Valor, _
        objParcelaOV.iDesconto2Codigo, objParcelaOV.dtDesconto2Ate, objParcelaOV.dDesconto2Valor, objParcelaOV.iDesconto3Codigo, objParcelaOV.dtDesconto3Ate, objParcelaOV.dDesconto3Valor)
        If lErro <> AD_SQL_SUCESSO Then gError 186810

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ParcelasOrcSRV_Grava = SUCESSO

    Exit Function

Erro_ParcelasOrcSRV_Grava:

    ParcelasOrcSRV_Grava = gErr

    Select Case gErr

        Case 186809
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 186810
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PARCELASORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186811)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensOrcamentoSRV_Grava(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Insere os itens do orcamento de Venda passado no BD

Dim alComando(0 To 4) As Long
Dim lErro As Long
Dim objItemOV As ClassItemOV
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim objItemOP As New ClassItemOP
Dim sProduto As String
Dim iStatus As Integer
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_ItensOrcamentoSRV_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186812
    Next

    If objOrcamentoVenda.lNumIntSolicSRV > 0 Then
    
        objSolicSRV.lNumIntDoc = objOrcamentoVenda.lNumIntSolicSRV
        
        lErro = CF("SolicitacaoSRV_Le_NumIntDoc", objSolicSRV)
        If lErro <> SUCESSO Then gError 186992

    End If
    
    'Para cada item da coleção
    For Each objItemOV In objOrcamentoVenda.colItensSRV

'        sProduto = String(STRING_PRODUTO, 0)
'
'        'Localiza o item no Banco de Dados
'        lErro = Comando_ExecutarPos(alComando(0), "SELECT Produto, NumIntDoc FROM ItensOrcamentoVenda WHERE CodOrcamento = ? AND FilialEmpresa = ? AND Produto = ? ", 0, sProduto, lNumIntDoc, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa, objItemOV.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 101136
'
'        lErro = Comando_BuscarPrimeiro(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101137
'        If lErro <> AD_SQL_SUCESSO Then

            'Obtém o número interno automático do item
            lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_ITEM_ORCAMENTOSRV", lNumIntDoc)
            If lErro <> SUCESSO Then gError 186813

            objItemOV.iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
            objItemOV.lCodOrcamento = objOrcamentoVenda.lCodigo
            objItemOV.lNumIntDoc = lNumIntDoc
            objItemOV.lNumIntDocOV = objOrcamentoVenda.lNumIntDoc

            'Insere o item no BD
            '#####################################
            'ALTERADO POR WAGNER
            lErro = Comando_Executar(alComando(0), "INSERT INTO ItensOrcamentoSRV (NumIntDoc,FilialEmpresa,CodOrcamento,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao, VersaoKit, VersaoKitBase, Status, MotivoPerda, Observacao, Peca) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDoc, objItemOV.iFilialEmpresa, objItemOV.lCodOrcamento, objItemOV.sProduto, objItemOV.dQuantidade, objItemOV.dPrecoUnitario, objItemOV.dPrecoTotal, objItemOV.iClasseUM, objItemOV.sUnidadeMed, objItemOV.dValorDesconto, objItemOV.dtDataEntrega, objItemOV.sDescricao, objItemOV.sVersaoKit, objItemOV.sVersaoKitBase, objItemOV.lStatus, objItemOV.lMotivoPerda, objItemOV.sObservacao, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 186814
            '#####################################
        
'        'Se for uma alteração
'        Else
'
'            objItemOV.lNumIntDoc = lNumIntDoc
'
'            'Faz a atualização no Banco de Dados
'            '#####################################
'            'ALTERADO POR WAGNER
'            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensOrcamentoVenda SET Produto =?, Quantidade = ?, PrecoUnitario = ?, PrecoTotal = ?, ClasseUM = ?, UnidadeMed = ?, ValorDesconto =?, DataEntrega = ?, Descricao = ?, VersaoKit = ? , VersaoKitBase = ? ", alComando(0), objItemOV.sProduto, objItemOV.dQuantidade, objItemOV.dPrecoUnitario, objItemOV.dPrecoTotal, objItemOV.iClasseUM, objItemOV.sUnidadeMed, objItemOV.dValorDesconto, objItemOV.dtDataEntrega, objItemOV.sDescricao, objItemOV.sVersaoKit, objItemOV.sVersaoKitBase)
'            If lErro <> AD_SQL_SUCESSO Then gError 101140
'            '#####################################
'
'            lErro = CF("ItemOV_ExcluiTributacao", objItemOV, alComando(2), alComando(3))
'            If lErro <> SUCESSO Then gError 101141
'
'            lErro = ItensOVGrade_Exclui(objItemOV)
'            If lErro <> SUCESSO Then gError 86361
'
'        End If

        lErro = CF("ItemOrcSRV_GravaTributacao", objItemOV, alComando(4))
        If lErro <> SUCESSO Then gError 186815
        
        lErro = CF("ItemOrcSRV_InsereProdSolicSRV", objItemOV, objOrcamentoVenda.colProdSolicSRV)
        If lErro <> SUCESSO Then gError 186975
        
    Next

    For Each objItemOV In objOrcamentoVenda.colItens

        'Obtém o número interno automático do item
        lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_ITEM_ORCAMENTOSRV", lNumIntDoc)
        If lErro <> SUCESSO Then gError 186913

        objItemOV.iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
        objItemOV.lCodOrcamento = objOrcamentoVenda.lCodigo
        objItemOV.lNumIntDoc = lNumIntDoc
        objItemOV.lNumIntDocOV = objOrcamentoVenda.lNumIntDoc

        'Insere o item no BD
        '#####################################
        'ALTERADO POR WAGNER
        lErro = Comando_Executar(alComando(0), "INSERT INTO ItensOrcamentoSRV (NumIntDoc,FilialEmpresa,CodOrcamento,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao, VersaoKit, VersaoKitBase, Status, MotivoPerda, Observacao, Peca) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDoc, objItemOV.iFilialEmpresa, objItemOV.lCodOrcamento, objItemOV.sProduto, objItemOV.dQuantidade, objItemOV.dPrecoUnitario, objItemOV.dPrecoTotal, objItemOV.iClasseUM, objItemOV.sUnidadeMed, objItemOV.dValorDesconto, objItemOV.dtDataEntrega, objItemOV.sDescricao, objItemOV.sVersaoKit, objItemOV.sVersaoKitBase, objItemOV.lStatus, objItemOV.lMotivoPerda, objItemOV.sObservacao, 1)
        If lErro <> AD_SQL_SUCESSO Then gError 186914
        '#####################################
        
        lErro = CF("ItemOrcSRV_GravaTributacao", objItemOV, alComando(4))
        If lErro <> SUCESSO Then gError 186915
        
    Next



'    '############################################
'    'Inserido por Wagner 22/03/2006
'    lErro = CF("OVItens_Grava_Customizado", objOrcamentoVenda)
'    If lErro <> SUCESSO Then gError 177046
'    '############################################

    'Fecha os comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensOrcamentoSRV_Grava = SUCESSO

    Exit Function

Erro_ItensOrcamentoSRV_Grava:

    ItensOrcamentoSRV_Grava = gErr

    Select Case gErr

        Case 186812
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186813, 186815, 186913, 186915, 186975, 186977

        Case 186814, 186914
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSORCSRV", gErr, objItemOV.iFilialEmpresa, objItemOV.lCodOrcamento)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186816)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ItemOrcSRV_GravaTributacao(ByVal objItemOV As ClassItemOV, ByVal lComando As Long) As Long

Dim lErro As Long

On Error GoTo Erro_ItemOrcSRV_GravaTributacao
'
'    With objItemOV.objTributacaoItemOV
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoItemOrcSRV (NumIntDoc, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objItemOV.lNumIntDoc, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 186817

    objItemOV.objTributacao.lNumIntDoc = objItemOV.lNumIntDocOV
    objItemOV.objTributacao.lNumIntDocItem = objItemOV.lNumIntDoc
    objItemOV.objTributacao.iItem = objItemOV.iItem
    objItemOV.objTributacao.dDescontoGrid = objItemOV.dValorDesconto
    objItemOV.objTributacao.dPrecoUnitario = objItemOV.dPrecoUnitario
    objItemOV.objTributacao.dPrecoUnitarioMoeda = objItemOV.dPrecoUnitario
    objItemOV.objTributacao.dQuantidade = objItemOV.dQuantidade
    objItemOV.objTributacao.sProduto = objItemOV.sProduto
    objItemOV.objTributacao.sProdutoDescricao = objItemOV.sDescricao
    objItemOV.objTributacao.sUnidadeMed = objItemOV.sUnidadeMed
    objItemOV.objTributacao.iTipoDoc = TIPODOC_TRIB_OSRV

    lErro = CF("TributacaoDocItem_Grava", objItemOV.objTributacao)
    If lErro <> SUCESSO Then gError 186817

    ItemOrcSRV_GravaTributacao = SUCESSO

    Exit Function

Erro_ItemOrcSRV_GravaTributacao:

    ItemOrcSRV_GravaTributacao = gErr

    Select Case gErr

        Case 186817
            'Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBITEMORCSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186818)

    End Select

    Exit Function

End Function


Public Function OrcamentoSRV_Altera_BD(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal objOrcamentoVenda_BD As ClassOrcamentoVenda, ByVal lComando1 As Long) As Long
'Faz as alterações no Orcamento de SRV

Dim lErro As Long
Dim objItemOV As New ClassItemOV
Dim objBloqueioPV_BD As New ClassBloqueioPV
Dim objBloqueioPV1 As ClassBloqueioPV
Dim bAchouBloqueioBD As Boolean
Dim vbMsgRes As VbMsgBoxResult
Dim objBloqueioPV As New ClassBloqueioPV
Dim iIndice As Integer
Dim iBloqueioEstoqueAdicionado As Integer
Dim dPedidoFaltaReservar As Double
Dim bPossuiBloqueioTotal As Boolean
Dim objMovEstoque As New ClassMovEstoque
Dim dQuantNaoCancelada As Double
Dim iCreditoAprovado As Integer
Dim objItemOV_BD As ClassItemOV
Dim bAdicionaBloqueio As Boolean
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando As Long

On Error GoTo Erro_OrcamentoSRV_Altera_BD

    lComando = Comando_Abrir
    If lComando = 0 Then gError 186869
    
    lComando2 = Comando_Abrir
    If lComando2 = 0 Then gError 186870
    
    lComando3 = Comando_Abrir
    If lComando3 = 0 Then gError 186871
    
    'Inicialização de flags
    bPossuiBloqueioTotal = False

    objOrcamentoVenda_BD.lCodigo = objOrcamentoVenda.lCodigo
    objOrcamentoVenda_BD.iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa

    'Lê e faz o "lock" no orçamento
    lErro = CF("OrcamentoSRV_Lock", objOrcamentoVenda_BD, lComando1)
    If lErro <> SUCESSO And lErro <> 186830 Then gError 186872
    
    If lErro <> SUCESSO Then gError 186873

    objOrcamentoVenda.lNumIntDoc = objOrcamentoVenda_BD.lNumIntDoc

    'Lê e locka os itens do orcamento
    lErro = CF("ItensOrcamentoSRV_Le_Lock", objOrcamentoVenda_BD)
    If lErro <> SUCESSO Then gError 186874
    
    lErro = CF("OrcamentoSRV_ExcluiTributacao", objOrcamentoVenda, lComando, lComando1, lComando2, lComando3)
    If lErro <> SUCESSO Then gError 186880
    
    lErro = CF("ItensOrcamentoSRV_Exclui", objOrcamentoVenda_BD)
    If lErro <> SUCESSO Then gError 186875
   
    lErro = CF("ItensOrcamentoSRV_Grava", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186876
    
    'Apaga os itens da tabela de parcelas
    lErro = CF("ParcelasOrcSRV_Exclui", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186877
    
    'grava os novos itens de parcelas no BD
    lErro = CF("ParcelasOrcSRV_Grava", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186878
    
    '#################################################
    'ALTERADO POR WAGNER
    lErro = Comando_ExecutarPos(lComando2, "UPDATE OrcamentoSRV SET NaturezaOp = ?, NomeFilialCli = ?, NomeCli = ?, Vendedor=?, FilialEmpresa=?, Codigo=?, Cliente=?, Filial=?, CondicaoPagto=?, PercAcrescFinanceiro=?, DataEmissao=?, DataReferencia=?, ValorTotal=?, ValorProdutos=?, ValorFrete=?, ValorDesconto=?, ValorSeguro=?, ValorOutrasDespesas=?, TabelaPreco=?, CobrancaAutomatica=?, NumIntDoc=?, PrazoValidade=?, Usuario = ?, CodigoBase = ?, Status=?, MotivoPerda=?, NumIntSolicSRV=?, ValorDescontoTit = ?, ValorItens = ?, ValorDescontoItens = ? ", lComando1, _
    objOrcamentoVenda.sNaturezaOp, objOrcamentoVenda.sNomeFilialCli, objOrcamentoVenda.sNomeCli, objOrcamentoVenda.iVendedor, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, objOrcamentoVenda.lCliente, objOrcamentoVenda.iFilial, objOrcamentoVenda.iCondicaoPagto, objOrcamentoVenda.dPercAcrescFinanceiro, objOrcamentoVenda.dtDataEmissao, objOrcamentoVenda.dtDataReferencia, objOrcamentoVenda.dValorTotal, objOrcamentoVenda.dValorProdutos, objOrcamentoVenda.dValorFrete, objOrcamentoVenda.dValorDesconto, objOrcamentoVenda.dValorSeguro, objOrcamentoVenda.dValorOutrasDespesas, objOrcamentoVenda.iTabelaPreco, objOrcamentoVenda.iCobrancaAutomatica, objOrcamentoVenda.lNumIntDoc, objOrcamentoVenda.iPrazoValidade, objOrcamentoVenda.sUsuario, objOrcamentoVenda.lCodigoBase, objOrcamentoVenda.lStatus, objOrcamentoVenda.lMotivoPerda, objOrcamentoVenda.lNumIntSolicSRV, _
    objOrcamentoVenda.dValorDescontoTit, objOrcamentoVenda.dValorItens, objOrcamentoVenda.dValorDescontoItens)
    If lErro <> AD_SQL_SUCESSO Then gError 186879
    '#################################################
    
'    lErro = CF("OrcamentoSRV_ExcluiTributacao", objOrcamentoVenda, lComando, lComando1, lComando2, lComando3)
'    If lErro <> SUCESSO Then gError 186880
    
    lErro = CF("OrcamentoSRV_GravaTributacao", objOrcamentoVenda, lComando1, lComando2)
    If lErro <> SUCESSO Then gError 186881
    
'    For iIndice = 1 To objOrcamentoVenda.colItens.Count
'
'        Set objItemOV = objOrcamentoVenda.colItens(iIndice)
'
'        lErro = CF("ItemOrcSRV_ExcluiTributacao", objItemOV, lComando2, lComando3)
'        If lErro <> SUCESSO Then gError 186882
'
'        lErro = CF("ItemOrcSRV_GravaTributacao", objItemOV, lComando2)
'        If lErro <> SUCESSO Then gError 186883
'
'    Next

    lErro = CF("BloqueiosGen_Altera", MAPBLOQGEN_TIPOTELA_ORCAMENTOSRV, objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 191990
    
    OrcamentoSRV_Altera_BD = SUCESSO

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando)
    
    Exit Function

Erro_OrcamentoSRV_Altera_BD:

    OrcamentoSRV_Altera_BD = gErr

    Select Case gErr

        Case 186869, 186870, 186871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 186872, 186874 To 186878, 186880 To 186883, 188044, 191990
        
        Case 186873
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_NAO_CADASTRADO", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
        
        Case 186879
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ORCAMENTOSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186884)

    End Select

    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Public Function ItensOrcamentoSRV_Exclui(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Exclui os itens presentes na coleção de itens do orcamento de servico passado

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objItemOV As ClassItemOV
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_ItensOrcamentoSRV_Exclui

    lComando = Comando_Abrir
    If lComando = 0 Then gError 186840
    
    lComando1 = Comando_Abrir
    If lComando1 = 0 Then gError 186841

    'Para cada item na coleção
    For Each objItemOV In objOrcamentoVenda.colItens

        'Busca o ítem no BD
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM ItensOrcamentoSRV WHERE NumIntDoc = ?", 0, lNumIntDoc, objItemOV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 186842

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186843
        
        If lErro = AD_SQL_SEM_DADOS Then gError 186844

        'Exclui o item da tabela de itens de Orcamento de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ItensOrcamentoSRV", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 186845
        
        lErro = CF("ItemOrcSRV_ExcluiTributacao", objItemOV, lComando, lComando1)
        If lErro <> SUCESSO Then gError 186846

    Next

    'Para cada item na coleção
    For Each objItemOV In objOrcamentoVenda.colItensSRV

        'Busca o ítem no BD
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM ItensOrcamentoSRV WHERE NumIntDoc = ?", 0, lNumIntDoc, objItemOV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 186842

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186843
        
        If lErro = AD_SQL_SEM_DADOS Then gError 186844

        'Exclui o item da tabela de itens de Orcamento de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ItensOrcamentoSRV", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 186845
        
        lErro = CF("ItemOrcSRV_ExcluiTributacao", objItemOV, lComando, lComando1)
        If lErro <> SUCESSO Then gError 186846

        lErro = CF("ItemOrcSRV_ExcluiProdSolicSRV", objItemOV)
        If lErro <> SUCESSO Then gError 186976

    Next

    Call Comando_Fechar(lComando)
    
    Call Comando_Fechar(lComando1)

    ItensOrcamentoSRV_Exclui = SUCESSO

    Exit Function

Erro_ItensOrcamentoSRV_Exclui:

    ItensOrcamentoSRV_Exclui = gErr

    Select Case gErr

        Case 186840, 186841
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186842, 186843
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORCAMENTOSRV1", gErr, objItemOV.lNumIntDoc)

        Case 186844
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMORCSRV_NAO_ENCONTRADO", gErr, objItemOV.lNumIntDoc)

        Case 186845
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSORCSRV", gErr, objItemOV.lNumIntDoc)

        Case 186846, 186976

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186847)

    End Select

    Call Comando_Fechar(lComando)
    
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Public Function ParcelasOrcSRV_Exclui(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Exclui do BD as Parcelas do Orcamento

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iNumParcela As Integer

On Error GoTo Erro_ParcelasOrcSRV_Exclui

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186848

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 186849

    'Busca as Parcelas do Orcamento de Venda
    lErro = Comando_ExecutarPos(lComando, "SELECT NumParcela FROM ParcelasOrcSRV WHERE FilialEmpresa = ? AND OrcamentoSRV = ?", 0, iNumParcela, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 186850

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186851

    'Para cada parcela encontrada
    Do While lErro <> AD_SQL_SEM_DADOS

        'Exclui a parcela da tabela de parcelas de Orcamento de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ParcelasOrcSRV", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 186852

        'Busca a Próxima
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186853

    Loop

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ParcelasOrcSRV_Exclui = SUCESSO

    Exit Function

Erro_ParcelasOrcSRV_Exclui:

    ParcelasOrcSRV_Exclui = gErr

    Select Case gErr

        Case 186848, 186849
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186850, 186851, 186853
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186852
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PARCELASORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186854)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Public Function OrcamentoSRV_ExcluiTributacao(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long) As Long

Dim lErro As Long
'Dim iItem As Integer
'Dim iFilialEmpresa As Integer

On Error GoTo Erro_OrcamentoSRV_ExcluiTributacao
'
'    lErro = Comando_ExecutarPos(lComando, "SELECT Item FROM TributacaoComplOrcSRV WHERE FilialEmpresa = ? AND Codorcamento = ? ", 0, iItem, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 186855
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186856
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TributacaoComplOrcSRV", lComando)
'        If lErro <> AD_SQL_SUCESSO Then gError 186857
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186858
'
'    Loop
'
'    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa FROM TributacaoOrcSRV WHERE FilialEmpresa = ? AND Codorcamento = ? ", 0, iFilialEmpresa, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 186859
'
'    lErro = Comando_BuscarPrimeiro(lComando2)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186860
'
'    If lErro <> AD_SQL_SUCESSO Then gError 186861
'
'    lErro = Comando_ExecutarPos(lComando3, "DELETE FROM TributacaoOrcSRV", lComando2)
'    If lErro <> AD_SQL_SUCESSO Then gError 186862

    lErro = CF("TributacaoDoc_Exclui", TIPODOC_TRIB_OSRV, objOrcamentoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 186862

    OrcamentoSRV_ExcluiTributacao = SUCESSO

    Exit Function

Erro_OrcamentoSRV_ExcluiTributacao:

    OrcamentoSRV_ExcluiTributacao = gErr

    Select Case gErr

'        Case 186855, 186856, 186858
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBCOMPLORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'
'        Case 186857
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBCOMPLORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'
'        Case 186859, 186860
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'
'        Case 186861
'            Call Rotina_Erro(vbOKOnly, "ERRO_TRIBORCSRV_NAO_ENCONTRADA", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186862
            'Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBUTACAOORCSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186863)

    End Select

    Exit Function

End Function

Public Function ItemOrcSRV_ExcluiTributacao(ByVal objItemOV As ClassItemOV, ByVal lComando As Long, ByVal lComando1 As Long) As Long

Dim lErro As Long
'Dim iNaturezaOpManual As Integer

On Error GoTo Erro_ItemOrcSRV_ExcluiTributacao

'    lErro = Comando_ExecutarPos(lComando, "SELECT NaturezaOpManual FROM TributacaoItemOrcSRV WHERE NumIntDoc = ? ", 0, iNaturezaOpManual, objItemOV.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 186864
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186865
'
'    If lErro <> AD_SQL_SUCESSO Then gError 186866
'
'    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TributacaoItemOrcSRV", lComando)
'    If lErro <> AD_SQL_SUCESSO Then gError 186867

    lErro = CF("TributacaoDocItem_Exclui", TIPODOC_TRIB_OSRV, objItemOV.lNumIntDoc)
    If lErro <> SUCESSO Then gError 186867

    ItemOrcSRV_ExcluiTributacao = SUCESSO

    Exit Function

Erro_ItemOrcSRV_ExcluiTributacao:

    ItemOrcSRV_ExcluiTributacao = gErr

    Select Case gErr
'
'        Case 186864, 186865
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOITEMORCSRV", gErr, objItemOV.lNumIntDoc)
'
'        Case 186866
'            Call Rotina_Erro(vbOKOnly, "ERRO_TRIBUTACAOITEMORCSRV_NAO_CADASTRADO", gErr, objItemOV.lNumIntDoc)

        Case 186867
            'Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBUTACAOITEMORCSRV", gErr, objItemOV.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186868)

    End Select

    Exit Function

End Function

Public Function OrcamentoSRV_Exclui(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Exclui o Orcamento de Venda passado

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim objItemOV As ClassItemOV

On Error GoTo Erro_OrcamentoSRV_Exclui

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 186901

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186902
    Next

    'Faz lock no orcamento
    lErro = CF("OrcamentoSRV_Lock", objOrcamentoVenda, alComando(0))
    If lErro <> SUCESSO And lErro <> 186830 Then gError 186903
    
    If lErro <> SUCESSO Then gError 186904

    'Lê e faz lock nos itens do Pedido
    lErro = CF("ItensOrcamentoSRV_Le_Lock", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186905
    
    'Exclui as Parcelas do Orcamento
    lErro = CF("ParcelasOrcSRV_Exclui", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186906

    'Exclui os itens do Orcamento
    lErro = CF("ItensOrcamentoSRV_Exclui", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186907

    'Exclui o Orcamento do BD
    lErro = CF("OrcamentoSRV_Exclui_BD", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 186908
    
    'Exclui a Tributacao e seus complementos
    lErro = CF("OrcamentoSRV_ExcluiTributacao", objOrcamentoVenda, alComando(0), alComando(1), alComando(2), alComando(3))
    If lErro <> SUCESSO Then gError 186909

'    '############################################
'    'Inserido por Wagner 22/03/2006
'    lErro = CF("ProjetoInfo_Exclui", PRJ_CR_TIPO_OV, objOrcamentoVenda.lNumIntDoc)
'    If lErro <> SUCESSO Then gError 181764
'    '############################################

    'Exclui os bloqueios do Pedido
    lErro = CF("BloqueiosGen_Exclui", MAPBLOQGEN_TIPOTELA_ORCAMENTOSRV, objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 193165
    
    lErro = CF("InfoAdicionais_Exclui", TIPODOC_INFOADIC_PSRV, objOrcamentoVenda.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 193165
    
    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 186910

    lErro = CF("WorkFlow_Trata_Transacao", "SRV", objOrcamentoVenda.objTela, gsUsuario)
    If lErro <> SUCESSO Then gError 186911

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrcamentoSRV_Exclui = SUCESSO

    Exit Function

Erro_OrcamentoSRV_Exclui:

    OrcamentoSRV_Exclui = gErr

    Select Case gErr

        Case 186901
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 186902
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186903, 186905 To 186909, 186911, 193165

        Case 186904
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)

        Case 186910
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186912)

    End Select

    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function OrcamentoSRV_Exclui_BD(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Exclui o Orcamento de SRV da tabela de OrcamentoSRV

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_OrcamentoSRV_Exclui_BD

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186894

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 186895
    
    'Busca o Orcamento no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM OrcamentoSRV WHERE FilialEmpresa = ? AND Codigo = ? ", 0, lNumIntDoc, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 186896

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186897
    
    If lErro = AD_SQL_SEM_DADOS Then gError 186898

    'Exclui o Orcamento da Tabela de Orcamentos
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM OrcamentoSRV", lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 186899

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    OrcamentoSRV_Exclui_BD = SUCESSO

    Exit Function

Erro_OrcamentoSRV_Exclui_BD:

    OrcamentoSRV_Exclui_BD = gErr

    Select Case gErr
    
        Case 186894, 186895
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 186896, 186897
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186898
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_NAO_CADASTRADO", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case 186899
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTOSRV", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186900)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Public Function ItemOrcSRV_ExcluiProdSolicSRV(ByVal objItemOV As ClassItemOV) As Long
'Insere as no BD Parcelas do orcamento passado

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_ItemOrcSRV_ExcluiProdSolicSRV

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186972
    Next
    
    'Busca o Orcamento no BD
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM ProdSolicSRV WHERE NumIntItensOrcSRV = ?", 0, lNumIntDoc, objItemOV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 188221

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188222
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Exclui os itens de ProdSolicSRV relacionados ao ItemOV
        lErro = Comando_Executar(alComando(3), "DELETE From ProdutoSRV WHERE NumIntProdSolicSRV = ?", lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 188224

        'Exclui os itens de ProdSolicSRV relacionados ao ItemOV
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ProdSolicSRV", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 188223

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188225

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOrcSRV_ExcluiProdSolicSRV = SUCESSO

    Exit Function

Erro_ItemOrcSRV_ExcluiProdSolicSRV:

    ItemOrcSRV_ExcluiProdSolicSRV = gErr

    Select Case gErr

        Case 186972
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 188821, 188222, 188225
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODSOLICSRV", objItemOV.lNumIntDoc)
        
        Case 188223
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRODSOLICSRV", gErr, objItemOV.lNumIntDoc)

        Case 188224
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRODUTOSRV", gErr, objItemOV.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188226)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function ItemOrcSRV_InsereProdSolicSRV(ByVal objItemOV As ClassItemOV, ByVal colProdSolicSRV As Collection) As Long
'Insere no BD o relacionamento de servicos do orcamento com solicitacao de servico

Dim lErro As Long
Dim lComando As Long
Dim objProdSolicSRV As ClassProdSolicSRV
Dim objItensSolicSrv As ClassItensSolicSRV
Dim lNumIntDoc As Long

On Error GoTo Erro_ItemOrcSRV_InsereProdSolicSRV

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186978
    
    'Para cada parcela da coleção
    For Each objProdSolicSRV In colProdSolicSRV

        If objProdSolicSRV.sServicoOrcSRV = objItemOV.sProduto Then

            'Gera o número interno automático da tabela ProdSolicSRV
            lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_PRODSOLICSRV", lNumIntDoc)
            If lErro <> SUCESSO Then gError 188129
            
            objProdSolicSRV.lNumIntDoc = lNumIntDoc

            'Insere em ProdSolicSRV
            lErro = Comando_Executar(lComando, "INSERT INTO ProdSolicSRV (NumIntDoc, NumIntItensOrcSRV, Produto,Quantidade, Lote, FilialOP, Garantia, Contrato) VALUES (?,?,?,?,?,?,?,?)", objProdSolicSRV.lNumIntDoc, objItemOV.lNumIntDoc, objProdSolicSRV.sProduto, objProdSolicSRV.dQuantidade, objProdSolicSRV.sLote, objProdSolicSRV.iFilialOP, objProdSolicSRV.lGarantia, objProdSolicSRV.sContrato)
            If lErro <> AD_SQL_SUCESSO Then gError 186979
                    
            lErro = CF("ProdutoSRV_Insere", objProdSolicSRV)
            If lErro <> SUCESSO Then gError 188220
                    
        End If

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItemOrcSRV_InsereProdSolicSRV = SUCESSO

    Exit Function

Erro_ItemOrcSRV_InsereProdSolicSRV:

    ItemOrcSRV_InsereProdSolicSRV = gErr

    Select Case gErr

        Case 186978
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 186979
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRODSOLICSRV", gErr, objItemOV.lNumIntDoc, objItensSolicSrv.lNumIntDoc)

        Case 188129, 188220

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186980)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ProdutoSRV_Insere(ByVal objProdSolicSRV As ClassProdSolicSRV) As Long
'Insere no BD as pecas usadas para fazerem um servico

Dim lErro As Long
Dim lComando As Long
Dim objProdutoSRV As ClassProdutoSRV

On Error GoTo Erro_ProdutoSRV_Insere

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 188217
    
    'Para cada parcela da coleção
    For Each objProdutoSRV In objProdSolicSRV.colProdutoSRV

        'Insere em ProdutoSRV
        lErro = Comando_Executar(lComando, "INSERT INTO ProdutoSRV (NumIntProdSolicSRV, Produto, Quantidade, Garantia, Contrato) VALUES (?,?,?,?,?)", objProdSolicSRV.lNumIntDoc, objProdutoSRV.sProduto, objProdutoSRV.dQuantidade, objProdutoSRV.lGarantia, objProdutoSRV.sContrato)
        If lErro <> AD_SQL_SUCESSO Then gError 188218
                    
    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ProdutoSRV_Insere = SUCESSO

    Exit Function

Erro_ProdutoSRV_Insere:

    ProdutoSRV_Insere = gErr

    Select Case gErr

        Case 188217
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 188218
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRODUTOSRV", gErr, objProdSolicSRV.lNumIntDoc, objProdutoSRV.sProduto, objProdutoSRV.lGarantia, objProdutoSRV.sContrato)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188219)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PedidoServico_Automatico(lNumPedido As Long) As Long
'Gera o próximo Número de um Pedido De Venda

Dim lErro As Long

On Error GoTo Erro_PedidoServico_Automatico

    lErro = CF("Config_ObterAutomatico", "SRVConfig", "NUM_PROX_PEDIDOSERVICO", "PedidoServico", "Codigo", lNumPedido)
    If lErro <> SUCESSO Then gError 188770

    PedidoServico_Automatico = SUCESSO

    Exit Function

Erro_PedidoServico_Automatico:

    PedidoServico_Automatico = gErr

    Select Case gErr

        Case 188770

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188771)

    End Select

    Exit Function

End Function

Function PedidoServico_AtualizaObsBloq_Trans(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'atualiza observacoes de bloqueios que já estavam no BD

Dim lErro As Long, sObservacao As String, iIndice As Integer
Dim objBloqueioPV As New ClassBloqueioPV, alComando(1 To 2) As Long

On Error GoTo Erro_PedidoServico_AtualizaObsBloq_Trans

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191484
    Next

    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
    
        If objBloqueioPV.iSequencial = 0 Then Exit For
    
        sObservacao = String(STRING_BLOQUEIOSPV_OBSERVACAO, 0)
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Observacao FROM BloqueiosPedidoSRV WHERE FilialEmpresa = ? AND PedidoSRV = ? AND Sequencial = ?", 0, sObservacao, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objBloqueioPV.iSequencial)
        If lErro <> AD_SQL_SUCESSO Then gError 191485
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191486
        
        If lErro <> AD_SQL_SUCESSO Then gError 191487
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE BloqueiosPedidoSRV SET Observacao = ?", alComando(1), objBloqueioPV.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 191488
        
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PedidoServico_AtualizaObsBloq_Trans = SUCESSO
     
    Exit Function
    
Erro_PedidoServico_AtualizaObsBloq_Trans:

    PedidoServico_AtualizaObsBloq_Trans = gErr
     
    Select Case gErr
          
        Case 191484
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 191485, 191486
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIO_PSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objBloqueioPV.iSequencial)
        
        Case 191487
            Call Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIO_PSRV_NAO_CADASTRADO", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objBloqueioPV.iSequencial)
            
        Case 191488
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_BLOQUEIO_PSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objBloqueioPV.iSequencial)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191489)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function PedidoServico_AtualizaObsBloq(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long, lTransacao As Long

On Error GoTo Erro_PedidoServico_AtualizaObsBloq

    'Abre a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 191526

    'Realiza a Gravação do BloqueioPV
    lErro = PedidoServico_AtualizaObsBloq_Trans(objPedidoVenda)
    If lErro <> SUCESSO Then gError 191527
    
    'Faz o Commit da Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 191528

    PedidoServico_AtualizaObsBloq = SUCESSO
     
    Exit Function
    
Erro_PedidoServico_AtualizaObsBloq:

    PedidoServico_AtualizaObsBloq = gErr
     
    Select Case gErr
          
        Case 191526
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 191527
        
        Case 191528
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191529)
     
    End Select
     
    Call Transacao_Rollback
    
    Exit Function

End Function

Function PedidoServico_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Grava o Pedido de Venda

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 29) As Long
Dim lTransacao As Long
Dim lNumIntDoc As Long
Dim objPedVenda_BD As New ClassPedidoDeVenda
Dim tPedVenda_BD As typePedidoVenda
Dim dQuantNaoCancelada As Double
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_PedidoServico_Grava

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 191940

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191941
    Next

    'Faz o lock nos objetos vinculados ao Pedido
    lErro = CF("PVObjetosVinculados_Lock", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191942

    'Faz o lock nos objetos vinculados ao Pedido
    lErro = CF("PSRVObjetosVinculados_Lock", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191943

    'Verifica se o Pedido passado está baixado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PedidoServico WHERE Codigo = ? AND FilialEmpresa = ? AND Status = ?", 0, lNumIntDoc, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa, STATUS_BAIXADO)
    If lErro <> AD_SQL_SUCESSO Then gError 191944

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 191945

    'Se estiver --> gErro. Sai.
    If lErro = AD_SQL_SUCESSO Then gError 191946

    'Lê o Pedido
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, DataEmissao, Cliente, Filial, FilialEmpresaFaturamento FROM PedidoServico WHERE Codigo = ? AND FilialEmpresa = ?", 0, tPedVenda_BD.lNumIntDoc, tPedVenda_BD.dtDataEmissao, tPedVenda_BD.lCliente, tPedVenda_BD.iFilial, tPedVenda_BD.iFilialEmpresaFaturamento, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 191947

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 191948

    'Se não existir
    If lErro <> AD_SQL_SUCESSO Then
        
        'Insere o Pedido novo no Banco de Dados
        lErro = PedidoServico_Insere_BD(objPedidoVenda)
        If lErro <> SUCESSO Then gError 191949
        
        'Grava os dados estatísticos - INSERÇÃO
        lErro = CF("PedidoSRV_Grava_Estatisticas", objPedidoVenda)
        If lErro <> SUCESSO Then gError 191950
        
    
    'Se existir
    Else

        objPedVenda_BD.lNumIntDoc = tPedVenda_BD.lNumIntDoc
        objPedVenda_BD.dtDataEmissao = tPedVenda_BD.dtDataEmissao
        objPedidoVenda.lNumIntDoc = tPedVenda_BD.lNumIntDoc
        objPedVenda_BD.lCliente = tPedVenda_BD.lCliente
        objPedVenda_BD.iFilial = tPedVenda_BD.iFilial
        objPedVenda_BD.iFilialEmpresaFaturamento = tPedVenda_BD.iFilialEmpresaFaturamento

        'Altera o Pedido de Venda e armazena os ítens do Pedido no BD em objPedVenda_BD
        lErro = PedidoSRV_Altera_BD(objPedidoVenda, objPedVenda_BD, alComando)
        If lErro <> SUCESSO Then gError 191951
                
        'Grava os dados estatísticos - ALTERAÇÃO
        lErro = CF("PedidoSRV_Grava_Estatisticas", objPedidoVenda, objPedVenda_BD)
        If lErro <> SUCESSO Then gError 191952
        
    End If

    'Verifica se tem quantidade não cancelada
    lErro = PedSRV_QuantNaoCancelada(objPedidoVenda, dQuantNaoCancelada)
    If lErro <> SUCESSO Then gError 191953

    'Se nao tiver quantidade nao cancelada, baixa o Pedido
    If dQuantNaoCancelada = 0 Then

        lErro = CF("PedidoServico_Baixar", objPedidoVenda)
        If lErro <> SUCESSO Then gError 191954
         
    End If
    
    'Vincula ou desvincula um Orçamento de Venda a um Pedido de Venda.
    lErro = PedidoServico_AlteraVinculo_Orcamento(objPedidoVenda) 'Inserida por Leo em 26/04/02
    If lErro <> SUCESSO Then gError 191955
    
'    '############################################
'    'Inserido por Wagner 22/03/2006
'    lErro = CF("ProjetoInfo_Grava", objPedidoVenda.objProjetoInfo, PRJ_CR_TIPO_PV, objPedidoVenda.lNumIntDoc)
'    If lErro <> SUCESSO Then gError 181706
'    '############################################

    objPedidoVenda.objInfoAdic.lNumIntDoc = objPedidoVenda.lNumIntDoc
    objPedidoVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_PSRV
    lErro = CF("InfoAdicionais_Grava", objPedidoVenda.objInfoAdic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Faz o commit na transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 191956

    lErro = CF("WorkFlow_Trata_Transacao", "SRV", objPedidoVenda.objTela, gsUsuario, objPedidoVenda)
    If lErro <> SUCESSO Then gError 178179
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoServico_Grava = SUCESSO

    Exit Function

Erro_PedidoServico_Grava:

    PedidoServico_Grava = gErr

    Select Case gErr

        Case 191940
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 191941
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 191942, 191943, 191949 To 191955

        Case 191944, 191945, 191947, 191948
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSERVICO", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)

        Case 191946
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOSERVICO_BAIXADO_ALTERACAO", gErr, objPedidoVenda.lCodigo)

        Case 191956
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191957)

    End Select

    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function PedidoServico_Insere_BD(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere o novo Pedido de Venda no Banco de Dados

Dim lErro As Long
Dim dQuantNaoCancelada As Double
Dim lNumIntDoc As Long
Dim objMovEstoque As New ClassMovEstoque
Dim objCliente As New ClassCliente 'Inserido por Wagner
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemPV As ClassItemPedido
Dim objBloqueioPV As ClassBloqueioPV
Dim iIndice As Integer, bLiberado As Boolean


On Error GoTo Erro_PedidoServico_Insere_BD

    'zera os campos de quantfaturada, quantop, quantsc para garantir que quando
    'pedido e' novo, mesmo que ele seja oriundo de um pedido baixado, estas quantidades estejam zeradas
    For Each objItemPV In objPedidoVenda.colItensPedido

        objItemPV.dQuantFaturada = 0
        objItemPV.dQuantOP = 0
        objItemPV.dQuantSC = 0
        
    Next

    For Each objItemPV In objPedidoVenda.colItensPedidoSRV

        objItemPV.dQuantFaturada = 0
        objItemPV.dQuantOP = 0
        objItemPV.dQuantSC = 0
        
    Next

'     'Para cada Bloqueio do Pedido
'    For iIndice = objPedidoVenda.colBloqueiosPV.Count To 1 Step -1
'        Set objBloqueioPV = objPedidoVenda.colBloqueiosPV.Item(iIndice)
'        If objBloqueioPV.dtDataLib <> DATA_NULA Then
'            objPedidoVenda.colBloqueiosPV.Remove iIndice
'        End If
'    Next
'
'    'Inclui no BD um bloqueio de pagamento antecipado, se necessario
'    lErro = CF("BloqueioPagtoAnt_Adiciona", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 191609
'
'    'Inclui no BD um bloqueio de dias de atraso, se necessario
'    lErro = CF("BloqueioDiasAtraso_Adiciona", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 191610
'
'    objCliente.lCodigo = objPedidoVenda.lCliente
'    'le os dados do cliente
'    lErro = CF("Cliente_Le", objCliente)
'    If lErro <> SUCESSO And lErro <> 12293 Then gError 191611
'
'    If objCliente.iAtivo <> CLIENTE_ATIVO Then gError 191612
'
'    'verifica se eh para criar ou nao bloqueio de limite de credito
'    If gobjCRFAT.iCreditoVerificaLimite = VERIFICAR_LIMITE_CREDITO Then
'        'Inclui no BD um bloqueio de Credito, se o cliente não tiver crédito
'        lErro = CF("BloqueioCredito_Adiciona", objPedidoVenda)
'        If lErro <> SUCESSO Then gError 191613
'    End If
'
'    'Adiciona um Bloqueio na coleção de Bloqueio de acordo com o a quantidade que falta reservar
'    lErro = CF("BloqueioReserva_Adiciona1", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 191614

    If gobjCRFAT.iCliAtrasoNaoGravaPV <> 0 Then
    
        If objPedidoVenda.lNumIntOrcVenda <> 0 Then
        
            bLiberado = True
        
        End If
        
        If bLiberado = False Then
        
            lErro = CF("PV_Grava_AnalisaCliente", objPedidoVenda)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
    End If
    
    'Retorna a soma das quantidades não canceladas de todos os itens
    lErro = PedSRV_QuantNaoCancelada(objPedidoVenda, dQuantNaoCancelada)
    If lErro <> SUCESSO Then gError 191615
    
    'Gera o número interno automático do Pedido de Venda
    lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_PEDIDOSRV", lNumIntDoc)
    If lErro <> SUCESSO Then gError 191616

    objPedidoVenda.lNumIntDoc = lNumIntDoc
    
    'Insere o Pedido na Tabela de Pedidos
    lErro = CF("PedidoServico_Grava_BD", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191617

    'Insere as comissões do Pedido no BD
    lErro = CF("ComissoesPedidoSRV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191618
    
    'Insere as Parcelas do Pedido no BD
    lErro = CF("ParcelasPedidoSRV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191619

    lErro = ProdutoFilial_Adiciona_QuantPedidaSRV(objPedidoVenda)
    If lErro <> SUCESSO Then gError 191620

    'Insere os itens do Pedido no BD (pecas)
    lErro = CF("ItensPSRV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191621

    'Insere os itens do Pedido no BD (servicos)
    lErro = CF("ItensPSRV_Grava1", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191621

'    'Insere as reservas do Pedido no BD
'    lErro = CF("ReservasPSRV_Grava", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 191622
'
'    'Cria um Movimento de Estoque
'    lErro = CF("MovEstoque_Cria", objMovEstoque, objPedidoVenda)
'    If lErro <> SUCESSO Then gError 191623
'
'    'Grava o Movimento de Estoque
'    lErro = CF("MovEstoque_Grava", objMovEstoque)
'    If lErro <> SUCESSO Then gError 191624
    
'    'Insere no BD os Bloqueios do Pedido de Venda
'    'Aqui o Pedido já está gravado e portanto acessivel no BD
'    lErro = CF("BloqueiosPedidoSRV_Grava", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 191625

    lErro = CF("BloqueiosGen_Gera", MAPBLOQGEN_TIPOTELA_PEDIDOSRV, objPedidoVenda)
    If lErro <> SUCESSO Then gError 191625
    
    PedidoServico_Insere_BD = SUCESSO

    Exit Function

Erro_PedidoServico_Insere_BD:

    PedidoServico_Insere_BD = gErr

    Select Case gErr

        Case 191609 To 191611, 191613 To 191625, ERRO_SEM_MENSAGEM

        Case 191612
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INATIVO", gErr, objPedidoVenda.lCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191626)

    End Select

    Exit Function

End Function

Function ItensPSRV_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere os itens do Pedido de Venda passado no BD (pecas)

Dim alComando(0 To 4) As Long
Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim objItemOP As New ClassItemOP
Dim sProduto As String
Dim iStatus As Integer

On Error GoTo Erro_ItensPSRV_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191547
    Next

    'Para cada item da coleção
    For Each objItemPV In objPedidoVenda.colItensPedido

        sProduto = String(STRING_PRODUTO, 0)
        
        lErro = CF("ItensPSRVAux_Grava", alComando(0), sProduto, iStatus, lNumIntDoc, objPedidoVenda, objItemPV)
        If lErro <> SUCESSO Then gError 191548

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191549
        
        If lErro <> AD_SQL_SUCESSO Then

            'Obtém o número interno automático do item
            lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_ITEM_PEDIDOSRV", lNumIntDoc)
            If lErro <> SUCESSO Then gError 191550

            objItemPV.iFilialEmpresa = giFilialEmpresa
            objItemPV.lCodPedido = objPedidoVenda.lCodigo
            objItemPV.lNumIntDoc = lNumIntDoc
            objItemPV.iPeca = 1
            
            If objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada <= 0 Then objItemPV.iStatus = STATUS_ATENDIDO
            
            lErro = CF("ItensPSRV_Insere", alComando(0), objItemPV)
            If lErro <> SUCESSO Then gError 191551
        
        'Se for uma alteração
        Else
            
            objItemPV.lNumIntDoc = lNumIntDoc
            
'            'Se houve alteracao de Produto
'            If sProduto <> objItemPV.sProduto Then
'
'                'Verifica se ítem estava vinculado a Ordem Produção
'                lErro = CF("ItemOP_Le_ItemPSRV", objItemOP, objItemPV)
'                If lErro <> SUCESSO And lErro <> 191559 Then gError 191552
'
'                If lErro = SUCESSO Then gError 191553
'
'            End If
            
            objItemPV.iStatus = iStatus
            If objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada <= 0 Then objItemPV.iStatus = STATUS_ATENDIDO
            
            'Faz a atualização no Banco de Dados
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensPedidoSRV SET Produto =?, Quantidade = ?, PrecoUnitario = ?, PrecoTotal = ?, ClasseUM = ?, UnidadeMed = ?, ValorDesconto =?, DataEntrega = ?, Descricao = ?, QuantCancelada = ?, QuantReservada = ?, QuantFaturada = ?, Status =? ", alComando(0), objItemPV.sProduto, objItemPV.dQuantidade, objItemPV.dPrecoUnitario, objItemPV.dPrecoTotal, objItemPV.iClasseUM, objItemPV.sUnidadeMed, objItemPV.dValorDesconto, objItemPV.dtDataEntrega, objItemPV.sDescricao, objItemPV.dQuantCancelada, objItemPV.dQuantReservada, objItemPV.dQuantFaturada, objItemPV.iStatus)
            If lErro <> AD_SQL_SUCESSO Then gError 191554

            lErro = CF("ItemPSRV_ExcluiTributacao", objItemPV, alComando(2), alComando(3))
            If lErro <> SUCESSO Then gError 191555
            
        End If

        objItemPV.lNumIntDocPV = objPedidoVenda.lNumIntDoc
        
        lErro = CF("ItemPSRV_GravaTributacao", objItemPV, alComando(4))
        If lErro <> SUCESSO Then gError 191556

    Next

    'Fecha os comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensPSRV_Grava = SUCESSO

    Exit Function

Erro_ItensPSRV_Grava:

    ItensPSRV_Grava = gErr

    Select Case gErr

        Case 191547
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 191548, 191549
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDOSRV1", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objItemPV.sProduto)

        Case 191550, 191551, 191552, 191555, 191556

        Case 191553
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PSRV_VINCULADO_ITEM_OP", gErr, objItemOP.lNumIntDoc)

        Case 191554
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDOSRV", gErr, objItemPV.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191555)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensPSRV_Grava1(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere os itens do Pedido de Venda passado no BD (servicos)

Dim alComando(0 To 4) As Long
Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim objItemOP As New ClassItemOP
Dim sProduto As String
Dim iStatus As Integer

On Error GoTo Erro_ItensPSRV_Grava1

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191673
    Next

    'Para cada item da coleção
    For Each objItemPV In objPedidoVenda.colItensPedidoSRV

        sProduto = String(STRING_PRODUTO, 0)
        
        lErro = CF("ItensPSRVAux_Grava", alComando(0), sProduto, iStatus, lNumIntDoc, objPedidoVenda, objItemPV)
        If lErro <> SUCESSO Then gError 191674

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191675
        
        If lErro <> AD_SQL_SUCESSO Then

            'Obtém o número interno automático do item
            lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_ITEM_PEDIDOSRV", lNumIntDoc)
            If lErro <> SUCESSO Then gError 191676

            objItemPV.iFilialEmpresa = giFilialEmpresa
            objItemPV.lCodPedido = objPedidoVenda.lCodigo
            objItemPV.lNumIntDoc = lNumIntDoc
            objItemPV.iPeca = 0
            
            If objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada <= 0 Then objItemPV.iStatus = STATUS_ATENDIDO
            
            lErro = CF("ItensPSRV_Insere", alComando(0), objItemPV)
            If lErro <> SUCESSO Then gError 191676
        
        'Se for uma alteração
        Else
            
            objItemPV.lNumIntDoc = lNumIntDoc
            
'            'Se houve alteracao de Produto
'            If sProduto <> objItemPV.sProduto Then
'
'                'Verifica se ítem estava vinculado a Ordem de Servico
'                lErro = CF("ItemOP_Le_ItemPSRV", objItemOP, objItemPV)
'                If lErro <> SUCESSO And lErro <> 46074 Then gError 191552
'
'                If lErro = SUCESSO Then gError 191553
'
'            End If
            
            objItemPV.iStatus = iStatus
            If objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada <= 0 Then objItemPV.iStatus = STATUS_ATENDIDO
            
            'Faz a atualização no Banco de Dados
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensPedidoSRV SET Produto =?, Quantidade = ?, PrecoUnitario = ?, PrecoTotal = ?, ClasseUM = ?, UnidadeMed = ?, ValorDesconto =?, DataEntrega = ?, Descricao = ?, QuantCancelada = ?, QuantReservada = ?, QuantFaturada = ?, Status =? ", alComando(0), objItemPV.sProduto, objItemPV.dQuantidade, objItemPV.dPrecoUnitario, objItemPV.dPrecoTotal, objItemPV.iClasseUM, objItemPV.sUnidadeMed, objItemPV.dValorDesconto, objItemPV.dtDataEntrega, objItemPV.sDescricao, objItemPV.dQuantCancelada, objItemPV.dQuantReservada, objItemPV.dQuantFaturada, objItemPV.iStatus)
            If lErro <> AD_SQL_SUCESSO Then gError 191677

            lErro = CF("ItemPSRV_ExcluiTributacao", objItemPV, alComando(2), alComando(3))
            If lErro <> SUCESSO Then gError 191678
            
        End If
        
        objItemPV.lNumIntDocPV = objPedidoVenda.lNumIntDoc
        
        lErro = CF("ItemPSRV_GravaTributacao", objItemPV, alComando(4))
        If lErro <> SUCESSO Then gError 191679

    Next

    'Fecha os comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensPSRV_Grava1 = SUCESSO

    Exit Function

Erro_ItensPSRV_Grava1:

    ItensPSRV_Grava1 = gErr

    Select Case gErr

        Case 191673
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 191674, 191675
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDOSRV1", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objItemPV.sProduto)

        Case 191676, 191678, 191679

'        Case 191553
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PSRV_VINCULADO_ITEM_OP", gErr, objItemOP.lNumIntDoc)
'
        Case 191677
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDOSRV", gErr, objItemPV.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191680)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensPSRVAux_Grava(ByVal lComando As Long, sProduto As String, iStatus As Integer, lNumIntDoc As Long, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objItemPV As ClassItemPedido) As Long
'cria ordens de produção a partir de pedidos de venda

Dim lErro As Long

On Error GoTo Erro_ItensPSRVAux_Grava

    'Localiza o item no Banco de Dados
    lErro = Comando_ExecutarPos(lComando, "SELECT Produto, Status, NumIntDoc FROM ItensPedidoSRV WHERE CodPedido = ? AND FilialEmpresa = ? AND Produto = ? ", 0, sProduto, iStatus, lNumIntDoc, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa, objItemPV.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 191537
        
    ItensPSRVAux_Grava = SUCESSO

    Exit Function

Erro_ItensPSRVAux_Grava:

    ItensPSRVAux_Grava = gErr

    Select Case gErr
    
        Case 191537

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191538)

    End Select

    Exit Function

End Function

Function ItensPSRV_Insere(ByVal lComando As Long, ByVal objItemPV As ClassItemPedido) As Long


Dim lErro As Long

On Error GoTo Erro_ItensPSRV_Insere

    'Insere o item no BD
    lErro = Comando_Executar(lComando, "INSERT INTO ItensPedidoSRV (NumIntDoc,FilialEmpresa,CodPedido,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC, Peca) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objItemPV.lNumIntDoc, objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.sProduto, objItemPV.dQuantidade, objItemPV.dPrecoUnitario, objItemPV.dPrecoTotal, objItemPV.iClasseUM, objItemPV.sUnidadeMed, objItemPV.dValorDesconto, objItemPV.dtDataEntrega, objItemPV.sDescricao, objItemPV.sLote, objItemPV.iStatus, objItemPV.dValorAbatComissao, objItemPV.dQuantCancelada, objItemPV.dQuantReservada, objItemPV.dQuantFaturada, objItemPV.dQuantOP, objItemPV.dQuantSC, objItemPV.iPeca)
    If lErro <> AD_SQL_SUCESSO Then gError 191539
        
    ItensPSRV_Insere = SUCESSO

    Exit Function

Erro_ItensPSRV_Insere:

    ItensPSRV_Insere = gErr

    Select Case gErr
    
        Case 191539
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPEDIDOSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191540)

    End Select

    Exit Function

End Function

Function ItemPSRV_ExcluiTributacao(ByVal objItemPV As ClassItemPedido, ByVal lComando As Long, ByVal lComando1 As Long) As Long

Dim lErro As Long
Dim iNaturezaOpManual As Integer

On Error GoTo Erro_ItemPSRV_ExcluiTributacao

'    lErro = Comando_ExecutarPos(lComando, "SELECT NaturezaOpManual FROM TributacaoItemPedidoSRV WHERE NumIntDoc = ? ", 0, iNaturezaOpManual, objItemPV.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 191541
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191542
'
'    If lErro <> AD_SQL_SUCESSO Then gError 191543
'
'    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TributacaoItemPedidoSRV", lComando)
'    If lErro <> AD_SQL_SUCESSO Then gError 191544

    lErro = CF("TributacaoDocItem_Exclui", TIPODOC_TRIB_PSRV, objItemPV.lNumIntDoc)
    If lErro <> SUCESSO Then gError 191544

    ItemPSRV_ExcluiTributacao = SUCESSO

    Exit Function

Erro_ItemPSRV_ExcluiTributacao:

    ItemPSRV_ExcluiTributacao = gErr

    Select Case gErr

'        Case 191541 To 191543
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBITEMPEDIDOSRV", gErr)

        Case 191544
            'Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBITEMPEDIDOSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191544)

    End Select

    Exit Function

End Function

Function ItemPSRV_GravaTributacao(ByVal objItemPV As ClassItemPedido, ByVal lComando As Long) As Long

Dim lErro As Long

On Error GoTo Erro_ItemPSRV_GravaTributacao

'    With objItemPV.objTributacaoItemPV
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoItemPedidoSRV (NumIntDoc, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objItemPV.lNumIntDoc, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 191545

    objItemPV.objTributacao.lNumIntDoc = objItemPV.lNumIntDocPV
    objItemPV.objTributacao.lNumIntDocItem = objItemPV.lNumIntDoc
    objItemPV.objTributacao.iItem = objItemPV.iItem
    objItemPV.objTributacao.dDescontoGrid = objItemPV.dValorDesconto
    objItemPV.objTributacao.dPrecoUnitario = objItemPV.dPrecoUnitario
    objItemPV.objTributacao.dPrecoUnitarioMoeda = objItemPV.dPrecoUnitario
    objItemPV.objTributacao.dQuantidade = objItemPV.dQuantidade - objItemPV.dQuantCancelada
    objItemPV.objTributacao.sProduto = objItemPV.sProduto
    objItemPV.objTributacao.sProdutoDescricao = objItemPV.sDescricao
    objItemPV.objTributacao.sUnidadeMed = objItemPV.sUnidadeMed
    objItemPV.objTributacao.iTipoDoc = TIPODOC_TRIB_PSRV

    lErro = CF("TributacaoDocItem_Grava", objItemPV.objTributacao)
    If lErro <> SUCESSO Then gError 191545
    
    ItemPSRV_GravaTributacao = SUCESSO

    Exit Function

Erro_ItemPSRV_GravaTributacao:

    ItemPSRV_GravaTributacao = gErr

    Select Case gErr

        Case 191545
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBITEMPSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191546)

    End Select

    Exit Function

End Function

Function PedidoServico_Grava_BD(ByVal objPedidoVenda As ClassPedidoDeVenda, Optional ByVal lComando1 As Long = 0) As Long
'Insere o Pedido de SErvico passado na Tabela de Pedidos de Servico
'ou atualiza caso o Pedido já exista

Dim lErro As Long
Dim alComando(0 To 6) As Long
Dim iIndice As Integer

On Error GoTo Erro_PedidoServico_Grava_BD

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191580
    Next

    'Verifica se é uma Inclusão
    If lComando1 = 0 Then
        'Insere o novo Pedido no BD
        lErro = CF("PedidoServico_Insere", objPedidoVenda, alComando(0))
        If lErro <> SUCESSO Then gError 191581
        
    Else 'lComando1 foi passado
        'Altera o Pedido de Venda
        lErro = Comando_ExecutarPos(alComando(0), "UPDATE PedidoServico SET DataEmissao = ?, FilialEntrega = ?, CodTransportadora = ?,CondicaoPagto =?, PercAcrescFinanceiro= ?, DataReferencia =?, MensagemNota =?,NaturezaOp = ?, ValorProdutos =?,ValorTotal =?, ValorFrete=?, ValorDesconto=?, ValorSeguro =?, ValorOutrasDespesas =?, PedidoCliente = ?, CanalVenda = ?,TabelaPreco =?, FaturaIntegral = ?, ComissaoAutomatica = ?, CobrancaAutomatica = ?,Placa = ? ,PlacaUF = ? ,VolumeQuant = ?, VolumeEspecie = ? ,VolumeMarca = ?,PesoLiq = ? ,PesoBruto = ?, VolumeNumero = ?, FreteRespons = ?, FilialEmpresaFaturamento = ?, CodTranspRedesp = ?, DetPagFrete = ?, DataEntrega = ?, PedidoRepresentante = ?, NumIntSolicSRV = ?, ValorDescontoTit = ?, ValorItens = ?, ValorDescontoItens = ? ", lComando1, _
            objPedidoVenda.dtDataEmissao, objPedidoVenda.iFilialEntrega, objPedidoVenda.iCodTransportadora, objPedidoVenda.iCondicaoPagto, objPedidoVenda.dPercAcrescFinanceiro, objPedidoVenda.dtDataReferencia, objPedidoVenda.sMensagemNota, objPedidoVenda.sNaturezaOp, objPedidoVenda.dValorProdutos, objPedidoVenda.dValorTotal, objPedidoVenda.dValorFrete, objPedidoVenda.dValorDesconto, objPedidoVenda.dValorSeguro, objPedidoVenda.dValorOutrasDespesas, objPedidoVenda.sPedidoCliente, objPedidoVenda.iCanalVenda, objPedidoVenda.iTabelaPreco, objPedidoVenda.iFaturaIntegral, objPedidoVenda.iComissaoAutomatica, objPedidoVenda.iCobrancaAutomatica, objPedidoVenda.sPlaca, objPedidoVenda.sPlacaUF, objPedidoVenda.lVolumeQuant, objPedidoVenda.lVolumeEspecie, objPedidoVenda.lVolumeMarca, objPedidoVenda.dPesoLiq, objPedidoVenda.dPesoBruto, objPedidoVenda.sVolumeNumero, objPedidoVenda.iFreteRespons, _
            objPedidoVenda.iFilialEmpresaFaturamento, objPedidoVenda.iCodTranspRedesp, objPedidoVenda.iDetPagFrete, objPedidoVenda.dtDataEntrega, objPedidoVenda.sPedidoRepresentante, objPedidoVenda.lNumIntSolicSRV, objPedidoVenda.dValorDescontoTit, objPedidoVenda.dValorItens, objPedidoVenda.dValorDescontoItens)
        If lErro <> AD_SQL_SUCESSO Then gError 191582
        
        lErro = CF("PedidoServico_ExcluiTributacao", objPedidoVenda)
        If lErro <> SUCESSO Then gError 191583
    End If

    lErro = CF("PedidoServico_GravaTributacao", objPedidoVenda, alComando(3), alComando(4))
    If lErro <> SUCESSO Then gError 191584

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoServico_Grava_BD = SUCESSO

    Exit Function

Erro_PedidoServico_Grava_BD:

    PedidoServico_Grava_BD = gErr

    Select Case gErr

        Case 191580
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 191581, 191583, 191584

        Case 191582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOSERVICO", gErr, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191585)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoServico_Insere(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long) As Long
'insere o psrv na tabela

Dim lErro As Long

On Error GoTo Erro_PedidoServico_Insere
            
    With objPedidoVenda
        lErro = Comando_Executar(lComando, "INSERT INTO PedidoServico (FilialEmpresa,Codigo,FilialEmpresaFaturamento,Cliente,Filial,FilialEntrega,CodTransportadora,CondicaoPagto,PercAcrescFinanceiro,DataEmissao,DataReferencia,MensagemNota,NaturezaOp,ValorTotal,ValorFrete,ValorDesconto,ValorSeguro,ValorOutrasDespesas,PedidoCliente,CanalVenda,TabelaPreco,ProxSeqBloqueio,FaturaIntegral,ComissaoAutomatica,CobrancaAutomatica,NumIntDoc,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,PesoLiq,PesoBruto,ValorProdutos,VolumeNumero,FreteRespons, CodTranspRedesp, DetPagFrete, DataEntrega, PedidoRepresentante, Status, NumIntSolicSRV, ValorDescontoTit, ValorItens, ValorDescontoItens) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresaFaturamento, objPedidoVenda.lCliente, objPedidoVenda.iFilial, _
            .iFilialEntrega, .iCodTransportadora, .iCondicaoPagto, .dPercAcrescFinanceiro, .dtDataEmissao, .dtDataReferencia, .sMensagemNota, .sNaturezaOp, .dValorTotal, .dValorFrete, .dValorDesconto, .dValorSeguro, .dValorOutrasDespesas, .sPedidoCliente, .iCanalVenda, .iTabelaPreco, .iProxSeqBloqueio, .iFaturaIntegral, .iComissaoAutomatica, .iCobrancaAutomatica, .lNumIntDoc, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .dPesoLiq, .dPesoBruto, .dValorProdutos, .sVolumeNumero, .iFreteRespons, .iCodTranspRedesp, .iDetPagFrete, .dtDataEntrega, .sPedidoRepresentante, STATUS_LANCADO, .lNumIntSolicSRV, .dValorDescontoTit, .dValorItens, .dValorDescontoItens)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 191561
       
    PedidoServico_Insere = SUCESSO
        
    Exit Function
   
Erro_PedidoServico_Insere:
    
    PedidoServico_Insere = gErr
    
    Select Case gErr
    
        Case 191561
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PEDIDOSRV", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191562)
            
    End Select

    Exit Function

End Function

Function PedidoServico_ExcluiTributacao(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
'Dim iItem As Integer
'Dim iFilialEmpresa As Integer
'Dim alComando(1 To 4) As Long
'Dim iIndice As Integer

On Error GoTo Erro_PedidoServico_ExcluiTributacao

'    'Abre o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 191580
'    Next
'
'    lErro = Comando_ExecutarPos(alComando(1), "SELECT Item FROM TributacaoComplPSRV WHERE FilialEmpresa = ? AND CodPedido = ? ", 0, iItem, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 191563
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191564
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM TributacaoComplPSRV", alComando(1))
'        If lErro <> AD_SQL_SUCESSO Then gError 191565
'
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191566
'
'    Loop
'
'    lErro = Comando_ExecutarPos(alComando(3), "SELECT FilialEmpresa FROM TributacaoPedSRV WHERE FilialEmpresa = ? AND CodPedido = ? ", 0, iFilialEmpresa, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 191567
'
'    lErro = Comando_BuscarPrimeiro(alComando(3))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191568
'
'    If lErro <> AD_SQL_SUCESSO Then gError 191569
'
'    lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM TributacaoPedSRV", alComando(3))
'    If lErro <> AD_SQL_SUCESSO Then gError 191570
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next

    lErro = CF("TributacaoDoc_Exclui", TIPODOC_TRIB_PSRV, objPedidoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 191570

    PedidoServico_ExcluiTributacao = SUCESSO

    Exit Function

Erro_PedidoServico_ExcluiTributacao:

    PedidoServico_ExcluiTributacao = gErr

    Select Case gErr

'        Case 191563, 191564, 191566
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBCOMPLPEDIDOSRV", gErr)
'
'        Case 191566
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBCOMPLPEDIDOSRV", gErr, objPedidoVenda.lCodigo)
'
'        Case 191567, 191568
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBPEDIDOSRV", gErr)
'
'        Case 191569
'            Call Rotina_Erro(vbOKOnly, "ERRO_TRIBPEDIDOSRV_NAO_ENCONTRADA", gErr, objPedidoVenda.lCodigo)

        Case 191570
            'Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBPEDIDOSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191571)

    End Select
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next

    Exit Function

End Function

Function PedidoServico_GravaTributacao(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long

On Error GoTo Erro_PedidoServico_GravaTributacao

'    With objPedidoVenda.objTributacaoPV
'
'        lErro = Comando_Executar(lComando1, "INSERT INTO TributacaoPedSRV (FilialEmpresa, CodPedido, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 191574
'
'    'inserir os complementos
'    lErro = CF("PedidoServico_GravaTributacaoCompl", objPedidoVenda, objPedidoVenda.objTributacaoPV.objTributacaoDesconto, lComando2)
'    If lErro <> SUCESSO Then gError 191575
'
'    lErro = CF("PedidoServico_GravaTributacaoCompl", objPedidoVenda, objPedidoVenda.objTributacaoPV.objTributacaoFrete, lComando2)
'    If lErro <> SUCESSO Then gError 191576
'
'    lErro = CF("PedidoServico_GravaTributacaoCompl", objPedidoVenda, objPedidoVenda.objTributacaoPV.objTributacaoOutras, lComando2)
'    If lErro <> SUCESSO Then gError 191577
'
'    lErro = CF("PedidoServico_GravaTributacaoCompl", objPedidoVenda, objPedidoVenda.objTributacaoPV.objTributacaoSeguro, lComando2)
'    If lErro <> SUCESSO Then gError 191578
'XXXX

    objPedidoVenda.objTributacao.lNumIntDoc = objPedidoVenda.lNumIntDoc
    objPedidoVenda.objTributacao.iTipoDoc = TIPODOC_TRIB_PSRV

    lErro = CF("TributacaoDoc_Grava", objPedidoVenda.objTributacao)
    If lErro <> SUCESSO Then gError 191575

    PedidoServico_GravaTributacao = SUCESSO

    Exit Function

Erro_PedidoServico_GravaTributacao:

    PedidoServico_GravaTributacao = gErr

    Select Case gErr

'        Case 191574
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBPSRV", gErr)

        Case 191575 'To 191578

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191579)

    End Select

    Exit Function

End Function

'Function PedidoServico_GravaTributacaoCompl(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objTribComplPV As ClassTributacaoItemPV, ByVal lComando As Long) As Long
'
'Dim lErro As Long, iItem As Integer
'
'On Error GoTo Erro_PedidoServico_GravaTributacaoCompl
'
'    iItem = objTribComplPV.lNumIntDoc
'
'
'    With objTribComplPV
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoComplPSRV (FilialEmpresa, CodPedido, Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPICodProduto, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, iItem, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, "", .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 191572
'
'    PedidoServico_GravaTributacaoCompl = SUCESSO
'
'    Exit Function
'
'Erro_PedidoServico_GravaTributacaoCompl:
'
'    PedidoServico_GravaTributacaoCompl = gErr
'
'    Select Case gErr
'
'        Case 191572
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBCOMPLPSRV", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191573)
'
'    End Select
'
'    Exit Function
'
'End Function

Function ComissoesPedidoSRV_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere no Bd as Comissões do Pedido de Venda passado

Dim lErro As Long
Dim lComando As Long
Dim objComissaoPV As ClassComissaoPedVendas

On Error GoTo Erro_ComissoesPedidoSRV_Grava

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 191586

    'Para cada Comissão da coleção
    For Each objComissaoPV In objPedidoVenda.colComissoes
        'Insere um registro em ComissoesPV
        'Tulio - Incluido campo Direto/Indireto na gravacao...
        lErro = Comando_Executar(lComando, "INSERT INTO ComissoesPedidoSRV (FilialEmpresa,PedidoSRV,CodVendedor,ValorBase,Percentual,Valor,PercentualEmissao,ValorEmissao,Indireta, Seq) VALUES (?,?,?,?,?,?,?,?,?,?) ", objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objComissaoPV.iCodVendedor, objComissaoPV.dValorBase, objComissaoPV.dPercentual, objComissaoPV.dValor, objComissaoPV.dPercentualEmissao, objComissaoPV.dValorEmissao, objComissaoPV.iIndireta, objComissaoPV.iSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 191587

    Next

    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ComissoesPedidoSRV_Grava = SUCESSO

    Exit Function

Erro_ComissoesPedidoSRV_Grava:

    ComissoesPedidoSRV_Grava = gErr
    
    Select Case gErr

        Case 191586
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 191587
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMISSOESPSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191588)

    End Select

    Exit Function

End Function

Function ParcelasPedidoSRV_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere as no BD Parcelas do Pedido passado

Dim lErro As Long
Dim lComando As Long
Dim objParcelaPV As ClassParcelaPedidoVenda

On Error GoTo Erro_ParcelasPedidoSRV_Grava

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 191589

    'Para cada parcela da coleção
    For Each objParcelaPV In objPedidoVenda.colParcelas

        'Insere em ParcelasPedidoDeVenda
        lErro = Comando_Executar(lComando, "INSERT INTO ParcelasPedidoSRV (FilialEmpresa,PedidoSRV,NumParcela,DataVencimento,Valor,Desconto1Codigo,Desconto1Ate,Desconto1Valor,Desconto2Codigo,Desconto2Ate,Desconto2Valor,Desconto3Codigo,Desconto3Ate,Desconto3Valor, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objParcelaPV.iNumParcela, objParcelaPV.dtDataVencimento, objParcelaPV.dValor, objParcelaPV.iDesconto1Codigo, objParcelaPV.dtDesconto1Ate, objParcelaPV.dDesconto1Valor, _
        objParcelaPV.iDesconto2Codigo, objParcelaPV.dtDesconto2Ate, objParcelaPV.dDesconto2Valor, objParcelaPV.iDesconto3Codigo, objParcelaPV.dtDesconto3Ate, objParcelaPV.dDesconto3Valor, objParcelaPV.iTipoPagto, objParcelaPV.iCodConta, objParcelaPV.dtDataCredito, objParcelaPV.dtDataEmissaoCheque, objParcelaPV.iBancoCheque, objParcelaPV.sAgenciaCheque, objParcelaPV.sContaCorrenteCheque, objParcelaPV.lNumeroCheque, objParcelaPV.dtDataDepositoCheque, objParcelaPV.iAdmMeioPagto, objParcelaPV.iParcelamento, objParcelaPV.sNumeroCartao, objParcelaPV.dtValidadeCartao, objParcelaPV.sAprovacaoCartao, objParcelaPV.dtDataTransacaoCartao)
        If lErro <> AD_SQL_SUCESSO Then gError 191590

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ParcelasPedidoSRV_Grava = SUCESSO

    Exit Function

Erro_ParcelasPedidoSRV_Grava:

    ParcelasPedidoSRV_Grava = gErr

    Select Case gErr

        Case 191589
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

        Case 191590
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PARCELASPSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191591)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Function ReservasPSRV_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Grava as Reservas
'
'Dim lErro As Long
'Dim lComando As Long
'Dim lComando1 As Long
'Dim objReserva As New ClassReserva
'Dim objItem As ClassItemPedido
'Dim lNumIntDoc As Long
'Dim lCodigo As Long
'Dim objItemRomaneio As ClassItemRomaneioGrade
'Dim objReservaItem As ClassReservaItem
'
'On Error GoTo Erro_ReservasPSRV_Grava
'
'    'Abrir comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 191592
'
'    lComando1 = Comando_Abrir()
'    If lComando1 = 0 Then gError 191593
'
'    'Para cada item do pedido
'    For Each objItem In objPedidoVenda.colItensPedido
'
'        'Zera a variável de número interno
'        lNumIntDoc = 0
'
'        'Gera o numero internos para as reservas
'        lErro = CF("Config_ObterNumInt", "MATConfig", "NUM_PROX_INT_RESERVA", lNumIntDoc, objItem.ColReserva.Count)
'        If lErro <> SUCESSO Then gError 191594
'
'        If objItem.colItensRomaneioGrade.Count = 0 Then
'
'            'Para cada reserva de cada item
'            For Each objReserva In objItem.ColReserva
'
'                'Obtém o Código Automático
'                lErro = CF("Reserva_Automatico_Trans", lCodigo)
'                If lErro <> SUCESSO Then gError 191595
'
'                 With objReserva
'                    .lNumIntDoc = lNumIntDoc
'                    .iTipoDoc = TIPO_PEDIDO_SRV_COD
'                    .dtDataReserva = gdtDataHoje
'                    .sCodUsuario = gsUsuario
'                    .lDocOrigem = objPedidoVenda.lCodigo
'                    .lNumIntOrigem = objItem.lNumIntDoc
'                    .lCodigo = lCodigo
'                    .sProduto = objItem.sProduto
'                    .iFilialEmpresa = objItem.iFilialEmpresa
'
'
'                    'Insere a reserva no Banco de Dados
'                    lErro = Comando_Executar(lComando, "INSERT INTO Reserva (FilialEmpresa,Codigo,Produto,Almoxarifado,TipoDoc,DocOrigem,NumIntOrigem,Quantidade,DataReserva,DataValidade,CodUsuario,Responsavel,NumIntDoc) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", .iFilialEmpresa, .lCodigo, .sProduto, .iAlmoxarifado, .iTipoDoc, .lDocOrigem, .lNumIntOrigem, .dQuantidade, .dtDataReserva, .dtDataValidade, .sCodUsuario, .sResponsavel, .lNumIntDoc)
'                    If lErro <> AD_SQL_SUCESSO Then gError 191596
'
'                End With
'                lNumIntDoc = lNumIntDoc + 1
'            Next
'
'        End If
'
'
'    Next
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    ReservasPSRV_Grava = SUCESSO
'
'    Exit Function
'
'Erro_ReservasPSRV_Grava:
'
'    ReservasPSRV_Grava = gErr
'
'    Select Case gErr
'
'        Case 191592, 191593
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 191594, 191595
'
'        Case 191596
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RESERVASPSRV", gErr, objPedidoVenda.lCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191597)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    Exit Function
'
'End Function

'Function BloqueiosPedidoSRV_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Insere no BD os Bloqueios do Pedido de Venda
'
'Dim lComando As Long
'Dim lErro As Long
'Dim objBloqueioPV As ClassBloqueioPV
'Dim iSequencial As Integer
'
'On Error GoTo Erro_BloqueiosPedidoSRV_Grava
'
'    'Abre o Comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 191605
'
'     'Para cada Bloqueio do Pedido
'    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
'
'        'Gera o sequencial automático para o Bloqueio
'        lErro = CF("BloqueioPSRV_Automatico", objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa, iSequencial)
'        If lErro <> SUCESSO Then gError 191606
'
'        'Atribui Sequencial ANTES de INSERIR
'        objBloqueioPV.iSequencial = iSequencial
'
'        'Insere o bloqueio na tabela de Bloqueios de Pedido de Venda
'        lErro = Comando_Executar(lComando, "INSERT INTO BloqueiosPedidoSRV(FilialEmpresa,PedidoSRV,Sequencial,TipoDeBloqueio,CodUsuario,Responsavel,Data,CodUsuarioLib,ResponsavelLib,DataLib,Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?)", objBloqueioPV.iFilialEmpresa, objBloqueioPV.lPedidoDeVendas, objBloqueioPV.iSequencial, objBloqueioPV.iTipoDeBloqueio, objBloqueioPV.sCodUsuario, objBloqueioPV.sResponsavel, objBloqueioPV.dtData, objBloqueioPV.sCodUsuarioLib, objBloqueioPV.sResponsavelLib, objBloqueioPV.dtDataLib, objBloqueioPV.sObservacao)
'        If lErro <> AD_SQL_SUCESSO Then gError 191607
'
'    Next
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    BloqueiosPedidoSRV_Grava = SUCESSO
'
'    Exit Function
'
'Erro_BloqueiosPedidoSRV_Grava:
'
'    BloqueiosPedidoSRV_Grava = gErr
'
'    Select Case gErr
'
'        Case 191605
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 191606
'
'        Case 191607
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BLOQUEIOPSRV", gErr, objBloqueioPV.iFilialEmpresa, objBloqueioPV.lPedidoDeVendas, iSequencial)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191608)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Function BloqueioPSRV_Automatico(ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer, iProxSeqBloqueio As Integer) As Long
''retorna o número da próximo Sequencial de Bloqueio disponivel
''OBS: ESTA FUNÇÃO DEVE SER CHAMADA DE DENTRO DE TRANSAÇÃO
'
'Dim lErro As Long
'Dim lComando As Long
'Dim lComando2 As Long
'Dim iFim_de_Arquivo As Integer
'Dim iSequencial_Automatico As Integer
'Dim iConteudo As Integer
'Dim iConteudo2 As Integer
'
'On Error GoTo Erro_BloqueioPSRV_Automatico
'
'    'Inicializa comandos
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 191598
'
'    lComando2 = Comando_Abrir()
'    If lComando2 = 0 Then gError 191599
'
'    'Pesquisa o número do Bloqueio gerado automaticamente pelo sistema
'    lErro = Comando_ExecutarPos(lComando, "SELECT ProxSeqBloqueio FROM PedidoServico WHERE Codigo = ? AND FilialEmpresa = ?", 0, iConteudo, lCodigo, iFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 191600
'
'    'Lê o registro com o número do Sequencial automático
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO Then gError 191601
'
'    'Loca o registro que guarda o numero do Sequencial gerado automaticamente
'    lErro = Comando_LockExclusive(lComando)
'    If lErro <> AD_SQL_SUCESSO Then gError 191602
'
'    If iConteudo = 0 Then iConteudo = 1 'p/garantir que comece no 1 em vez zero
'
'    iSequencial_Automatico = iConteudo
'
'    'Valor de retorno da funcao
'    iProxSeqBloqueio = iSequencial_Automatico
'
'    iSequencial_Automatico = iSequencial_Automatico + 1
'
'    iConteudo2 = iSequencial_Automatico
'
'    'Atualiza o numero do Sequencial gerado automaticamente
'    lErro = Comando_ExecutarPos(lComando2, "UPDATE PedidoServico SET ProxSeqBloqueio = ?", lComando, iConteudo2)
'    If lErro <> AD_SQL_SUCESSO Then gError 191603
'
'    'Fecha comandos
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando2)
'
'    BloqueioPSRV_Automatico = SUCESSO
'
'    Exit Function
'
'Erro_BloqueioPSRV_Automatico:
'
'    BloqueioPSRV_Automatico = gErr
'
'    Select Case gErr
'
'        Case 191598, 191599
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 191600, 191601
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSERVICO", gErr, iFilialEmpresa, lCodigo)
'
'        Case 191602
'            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOSRV", gErr, lCodigo)
'
'        Case 191603
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOSERVICO", gErr, lCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191604)
'
'    End Select
'
'    'Fecha comandos
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando2)
'
'    Exit Function
'
'End Function

Function PedidoSRV_Grava_Estatisticas(ByVal objPedVenda As ClassPedidoDeVenda, Optional ByVal objPedVenda_BD As ClassPedidoDeVenda) As Long
'Grava a estatística dos produtos Pedidos

Dim lErro As Long

On Error GoTo Erro_PedidoSRV_Grava_Estatisticas
           
    If objPedVenda_BD Is Nothing Then
        
        'Grava a estatistica do Dia do produto
        lErro = SldDiaFat_Grava_EstPSRV(objPedVenda)
        If lErro <> SUCESSO Then gError 191668
    
        'Grava a estatística do Mês do produto
        lErro = SldMesFat_Grava_EstPSRV(objPedVenda)
        If lErro <> SUCESSO Then gError 191669

    Else
            
        'Grava a estatística do Dia do produto
        lErro = SldDiaFat_Grava_EstPSRV(objPedVenda, objPedVenda_BD)
        If lErro <> SUCESSO Then gError 191670
    
        'Grava a estatística do Mês do produto
        lErro = SldMesFat_Grava_EstPSRV(objPedVenda, objPedVenda_BD)
        If lErro <> SUCESSO Then gError 191671
    
    End If
            
    PedidoSRV_Grava_Estatisticas = SUCESSO
    
    Exit Function
    
Erro_PedidoSRV_Grava_Estatisticas:
    
    PedidoSRV_Grava_Estatisticas = Err

    Select Case Err
        
        Case 191668 To 191671
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 191672)

    End Select
    
    Exit Function

End Function

Private Function SldDiaFat_Grava_EstPSRV(ByVal objPedVenda As ClassPedidoDeVenda, Optional ByVal objPedVenda_BD As ClassPedidoDeVenda) As Long
'Grava a estatística do dia do Produto Pedido

Dim lErro As Long
Dim objItemPV As New ClassItemPedido
Dim objItemPV_BD As New ClassItemPedido
Dim objSldDiaFat As New ClassSldDiaFat
Dim dQuantUMVenda As Double
Dim dFator As Double
Dim alComando(1 To 5) As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_SldDiaFat_Grava_EstPSRV
        
    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191640
    Next
          
    'Para cada item do PedidoVenda
    For Each objItemPV In objPedVenda.colItensPedido
        
        lErro = SldDiaFat_Grava_EstPSRV1(alComando, objItemPV, objPedVenda)
        If lErro <> SUCESSO Then gError 191641
    
    Next
    
    'Para cada item do PedidoVenda
    For Each objItemPV In objPedVenda.colItensPedidoSRV
        
        lErro = SldDiaFat_Grava_EstPSRV1(alComando, objItemPV, objPedVenda)
        If lErro <> SUCESSO Then gError 191642
    
    Next
    
    'Se o parametro foi preenchido é um pedido alteração
    If Not (objPedVenda_BD Is Nothing) Then
            
        For Each objItemPV_BD In objPedVenda_BD.colItensPedido
        
            lErro = SldDiaFat_Grava_EstPSRV2(alComando, objItemPV_BD, objPedVenda_BD)
            If lErro <> SUCESSO Then gError 191643
            
        Next
        
        For Each objItemPV_BD In objPedVenda_BD.colItensPedidoSRV
        
            lErro = SldDiaFat_Grava_EstPSRV2(alComando, objItemPV_BD, objPedVenda_BD)
            If lErro <> SUCESSO Then gError 191644
            
        Next
        
    End If
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    SldDiaFat_Grava_EstPSRV = SUCESSO
    
    Exit Function
    
Erro_SldDiaFat_Grava_EstPSRV:
    
    SldDiaFat_Grava_EstPSRV = gErr

    Select Case gErr
    
        Case 191640
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 191641 To 191644
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191645)

    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function SldDiaFat_Grava_EstPSRV1(alComando() As Long, ByVal objItemPV As ClassItemPedido, ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Grava a estatística do dia do Produto Pedido

Dim lErro As Long
Dim objSldDiaFat As New ClassSldDiaFat
Dim dQuantUMVenda As Double
Dim dFator As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_SldDiaFat_Grava_EstPSRV1
        
    objProduto.sCodigo = objItemPV.sProduto
    
    'Lê o produto para pegar a UM_Venda
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 191627

    If lErro = 28030 Then gError 191628
        
    'Lê o fator de conversão da UM_Venda
    lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 191629

    'Converte a UM da quantidade
    dQuantUMVenda = ((objItemPV.dQuantidade - objItemPV.dQuantCancelada) * dFator)
    
    objSldDiaFat.iFilialEmpresa = giFilialEmpresa
    objSldDiaFat.dtData = objPedVenda.dtDataEmissao
    objSldDiaFat.sProduto = objItemPV.sProduto
    
    'Lê todos os atributos na tabela SldDiaFat e faz lock exclusive
    lErro = CF("SldDiaFat_Le_Lock", alComando(1), objSldDiaFat)
    If lErro <> SUCESSO And lErro <> 52729 Then gError 191630
                
    'Se não encontrou
    If lErro = 52729 Then
        
        'Insere na tabela de SldDiaFat um novo registro
        lErro = Comando_Executar(alComando(5), "INSERT INTO SldDiaFat (FilialEmpresa, Data, Produto, QuantPedidaSRV) VALUES (?,?,?,?)", giFilialEmpresa, objPedVenda.dtDataEmissao, objItemPV.sProduto, dQuantUMVenda)
        If lErro <> AD_SQL_SUCESSO Then gError 191631
    
    Else
        'Se não atualiza
        
        objSldDiaFat.dQuantPedidaSRV = objSldDiaFat.dQuantPedidaSRV + dQuantUMVenda
        
        'Atualiza a tabela de SldDiaFat
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE SldDiaFat SET QuantPedidaSRV =?", alComando(1), objSldDiaFat.dQuantPedidaSRV)
        If lErro <> AD_SQL_SUCESSO Then gError 191632
    
    End If

    SldDiaFat_Grava_EstPSRV1 = SUCESSO
    
    Exit Function
    
Erro_SldDiaFat_Grava_EstPSRV1:
    
    SldDiaFat_Grava_EstPSRV1 = gErr

    Select Case gErr
    
        Case 191627, 191629, 191630
        
        Case 191628
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 191631
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SLDDIAFAT", gErr, objSldDiaFat.iFilialEmpresa, objSldDiaFat.sProduto, objSldDiaFat.dtData)
        
        Case 191632
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDDIAFAT", gErr, objSldDiaFat.iFilialEmpresa, objSldDiaFat.sProduto, objSldDiaFat.dtData)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191633)

    End Select

End Function

Private Function SldDiaFat_Grava_EstPSRV2(alComando() As Long, ByVal objItemPV_BD As ClassItemPedido, ByVal objPedVenda_BD As ClassPedidoDeVenda) As Long
'Grava a estatística do dia do Produto Pedido

Dim lErro As Long
Dim objSldDiaFat As New ClassSldDiaFat
Dim dQuantUMVenda As Double
Dim dFator As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_SldDiaFat_Grava_EstPSRV2

    objSldDiaFat.iFilialEmpresa = giFilialEmpresa
    objSldDiaFat.dtData = objPedVenda_BD.dtDataEmissao
    objSldDiaFat.sProduto = objItemPV_BD.sProduto
    
    'Lê todos os atributos na tabela SldDiaFat e faz lock exclusive
    lErro = CF("SldDiaFat_Le_Lock", alComando(3), objSldDiaFat)
    If lErro <> SUCESSO Then gError 191634
                
    objProduto.sCodigo = objItemPV_BD.sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 191635
    
    If lErro = 28030 Then gError 191636
            
    'Calcula o fator de conversão da UM_Venda
    lErro = CF("UM_Conversao_Trans", objItemPV_BD.iClasseUM, objItemPV_BD.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 191637

    'Converte a quantidade
    dQuantUMVenda = ((objItemPV_BD.dQuantidade - objItemPV_BD.dQuantCancelada) * dFator)
        
    objSldDiaFat.dQuantPedidaSRV = objSldDiaFat.dQuantPedidaSRV - dQuantUMVenda
    
    'Atualiza a tabela de SldDiaFat para alteração
    lErro = Comando_ExecutarPos(alComando(4), "UPDATE SldDiaFat SET QuantPedidaSRV =?", alComando(3), objSldDiaFat.dQuantPedidaSRV)
    If lErro <> AD_SQL_SUCESSO Then gError 191638

    SldDiaFat_Grava_EstPSRV2 = SUCESSO
    
    Exit Function
    
Erro_SldDiaFat_Grava_EstPSRV2:
    
    SldDiaFat_Grava_EstPSRV2 = gErr

    Select Case gErr
    
        Case 191634, 191635, 191637
        
        Case 191636
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 191638
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDDIAFAT", gErr, objSldDiaFat.iFilialEmpresa, objSldDiaFat.sProduto, objSldDiaFat.dtData)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191639)

    End Select

    Exit Function

End Function

Private Function SldMesFat_Grava_EstPSRV(ByVal objPedVenda As ClassPedidoDeVenda, Optional ByVal objPedVenda_BD As ClassPedidoDeVenda) As Long
'Grava a estatística do Mês do produto

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim objSldMesFat As New ClassSldMesFat
Dim iMes As Integer
Dim dFator As Double
Dim dQuantUMVenda As Double
Dim sMes As String
Dim objItemPV_BD As ClassItemPedido
Dim alComando(1 To 5) As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_SldMesFat_Grava_EstPSRV

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191662
    Next

    'Para cada item do Pedido de Venda
    For Each objItemPV In objPedVenda.colItensPedido
        
        lErro = SldMesFat_Grava_EstPSRV1(alComando, objItemPV, objPedVenda)
        If lErro <> SUCESSO Then gError 191663
        
        
    Next
    
    'Para cada item do Pedido de Venda
    For Each objItemPV In objPedVenda.colItensPedidoSRV
        
        lErro = SldMesFat_Grava_EstPSRV1(alComando, objItemPV, objPedVenda)
        If lErro <> SUCESSO Then gError 191664
        
        
    Next
    
    'Se o parametro foi preenchido é um pedido alteração
    If Not (objPedVenda_BD Is Nothing) Then
    
        'Para cada objItemPV
        For Each objItemPV_BD In objPedVenda_BD.colItensPedido
            
            lErro = SldMesFat_Grava_EstPSRV2(alComando, objItemPV_BD, objPedVenda_BD)
            If lErro <> SUCESSO Then gError 191665
    
        Next
        
        'Para cada objItemPV
        For Each objItemPV_BD In objPedVenda_BD.colItensPedidoSRV
            
            lErro = SldMesFat_Grava_EstPSRV2(alComando, objItemPV_BD, objPedVenda_BD)
            If lErro <> SUCESSO Then gError 191666
    
        Next
        
        
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    SldMesFat_Grava_EstPSRV = SUCESSO

    Exit Function

Erro_SldMesFat_Grava_EstPSRV:

    SldMesFat_Grava_EstPSRV = gErr

    Select Case gErr
        
        Case 191662
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 191663 To 191666
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191667)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function SldMesFat_Grava_EstPSRV1(alComando() As Long, ByVal objItemPV As ClassItemPedido, ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Grava a estatística do Mês do produto

Dim lErro As Long
Dim objSldMesFat As New ClassSldMesFat
Dim iMes As Integer
Dim dFator As Double
Dim dQuantUMVenda As Double
Dim sMes As String
Dim objItemPV_BD As ClassItemPedido
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_SldMesFat_Grava_EstPSRV1

        
    objSldMesFat.iFilialEmpresa = giFilialEmpresa
    objSldMesFat.iAno = Year(objPedVenda.dtDataEmissao)
    objSldMesFat.sProduto = objItemPV.sProduto
    
    'Lê todos os atributos na tabela SldMesFat e faz lock exclusive
    lErro = CF("SldMesFat_Le_Lock", alComando(1), objSldMesFat)
    If lErro <> SUCESSO And lErro <> 52733 Then gError 191652
            
    If lErro = 52733 Then
        
        iMes = Month(objPedVenda.dtDataEmissao)
    
        objProduto.sCodigo = objItemPV.sProduto
        
        'Lê o produto para pegar a UMVenda
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 191653
        
        If lErro = 28030 Then gError 191654
                
        'Calcula o fator da UM_Venda
        lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError 191655
    
        'Converte a quantidade
        dQuantUMVenda = ((objItemPV.dQuantidade - objItemPV.dQuantCancelada) * dFator)
                    
        sMes = CStr(iMes)

        'Atualiza a tabela de SldMesFat para inserção
        lErro = Comando_Executar(alComando(2), "INSERT INTO SldMesFat (FilialEmpresa, Ano, Produto, QuantPedidaSRV" + sMes + ") VALUES (?,?,?,?)", giFilialEmpresa, objSldMesFat.iAno, objSldMesFat.sProduto, dQuantUMVenda)
        If lErro <> AD_SQL_SUCESSO Then gError 191656

    Else
                
        iMes = Month(objPedVenda.dtDataEmissao)
        
        objProduto.sCodigo = objItemPV.sProduto
        
        'Lê o produto para pegar a UMVenda
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 191657
        
        If lErro = 28030 Then gError 191658
                
        'Calcula o fator para a UM_Venda
        lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError 191659
    
        'converte a quantidade
        dQuantUMVenda = ((objItemPV.dQuantidade - objItemPV.dQuantCancelada) * dFator)
            
        objSldMesFat.dQuantPedidaSRV(iMes) = objSldMesFat.dQuantPedidaSRV(iMes) + dQuantUMVenda
        
        sMes = CStr(iMes)

        'Atualiza a tabela de SldMesFat para inserção
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE SldMesFat SET QuantPedidaSRV" + sMes + "=?", alComando(1), objSldMesFat.dQuantPedidaSRV(iMes))
        If lErro <> AD_SQL_SUCESSO Then gError 191660
    
    End If
        
    SldMesFat_Grava_EstPSRV1 = SUCESSO

    Exit Function

Erro_SldMesFat_Grava_EstPSRV1:

    SldMesFat_Grava_EstPSRV1 = gErr

    Select Case gErr
        
        Case 191652, 191653, 191655, 191657, 191659
        
        Case 191654, 191658
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 191656
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SLDMESFAT", gErr, objSldMesFat.iAno, objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto)
        
        Case 191660
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESFAT", gErr, objSldMesFat.iAno, objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191661)

    End Select

    Exit Function
    
End Function

Private Function SldMesFat_Grava_EstPSRV2(alComando() As Long, ByVal objItemPV_BD As ClassItemPedido, ByVal objPedVenda_BD As ClassPedidoDeVenda) As Long
'Grava a estatística do Mês do produto

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim objSldMesFat As New ClassSldMesFat
Dim iMes As Integer
Dim dFator As Double
Dim dQuantUMVenda As Double
Dim sMes As String
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_SldMesFat_Grava_EstPSRV2

    objSldMesFat.iFilialEmpresa = giFilialEmpresa
    objSldMesFat.iAno = Year(objPedVenda_BD.dtDataEmissao)
    objSldMesFat.sProduto = objItemPV_BD.sProduto
    
    'Lê todos os atributos na tabela SldMesFat e faza lock exclusive
    lErro = CF("SldMesFat_Le_Lock", alComando(3), objSldMesFat)
    If lErro <> SUCESSO Then gError 191646
    
    iMes = Month(objPedVenda_BD.dtDataEmissao)
    
    objProduto.sCodigo = objItemPV_BD.sProduto
    
    'Lê o produto para pegar a UMVenda
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 191647
    
    If lErro = 28030 Then gError 191648
            
    'Calcula o fator da UM_Venda
    lErro = CF("UM_Conversao_Trans", objItemPV_BD.iClasseUM, objItemPV_BD.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 191649

    'converte a quantidade
    dQuantUMVenda = ((objItemPV_BD.dQuantidade - objItemPV_BD.dQuantCancelada) * dFator)
        
    objSldMesFat.dQuantPedidaSRV(iMes) = objSldMesFat.dQuantPedidaSRV(iMes) - dQuantUMVenda
    
    sMes = CStr(iMes)

    'Atualiza a tabela de SldMesFat para alteração
    lErro = Comando_ExecutarPos(alComando(4), "UPDATE SldMesFat SET QuantPedidaSRV" + sMes + "=?", alComando(3), objSldMesFat.dQuantPedidaSRV(iMes))
    If lErro <> AD_SQL_SUCESSO Then gError 191650
    

    SldMesFat_Grava_EstPSRV2 = SUCESSO

    Exit Function

Erro_SldMesFat_Grava_EstPSRV2:

    SldMesFat_Grava_EstPSRV2 = gErr

    Select Case gErr
        
        Case 191646, 191647, 191649
        
        Case 191648
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 191650
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESFAT", gErr, objSldMesFat.iAno, objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191651)

    End Select

    Exit Function
    
End Function

Private Function PedidoSRV_Altera_BD(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objPedVenda_BD As ClassPedidoDeVenda, alComando() As Long) As Long
'Faz as alterações no Pedido de Servico
'Passa os dados de ítens de PedidoVenda no BD para objPedVenda_BD

Dim lErro As Long
Dim objPedidoVenda_BD As New ClassPedidoDeVenda
Dim objItemPV As ClassItemPedido
Dim objBloqueioPV_BD As New ClassBloqueioPV
Dim objBloqueioPV1 As ClassBloqueioPV
Dim bAchouBloqueioBD As Boolean
Dim vbMsgRes As VbMsgBoxResult
Dim objBloqueioPV As New ClassBloqueioPV
Dim iIndice As Integer
Dim iBloqueioEstoqueAdicionado As Integer
Dim dPedidoFaltaReservar As Double
Dim bPossuiBloqueioTotal As Boolean
Dim objMovEstoque As New ClassMovEstoque
Dim dQuantNaoCancelada As Double
Dim iCreditoAprovado As Integer
Dim objItemPV_BD As ClassItemPedido
Dim bAdicionaBloqueio As Boolean
Dim iAtraso As Integer, colBloqBD As New Collection
Dim bPossuiBloqueioCredAtivo As Boolean
Dim objCliente As New ClassCliente 'Inserido Por Wagner

On Error GoTo Erro_PedidoSRV_Altera_BD

    'Inicialização de flags
    bPossuiBloqueioTotal = False

    objPedidoVenda_BD.lCodigo = objPedidoVenda.lCodigo
    objPedidoVenda_BD.iFilialEmpresa = objPedidoVenda.iFilialEmpresa

    'Lê e faz o "lock" no Pedido
    lErro = CF("PedidoServico_Lock", objPedidoVenda_BD, alComando(0))
    If lErro <> SUCESSO And lErro <> 191683 Then gError 191973

    If lErro = 191683 Then gError 191974

    objPedidoVenda.lNumIntDoc = objPedidoVenda_BD.lNumIntDoc

    'Lê e locka os itens do Pedido
    lErro = CF("ItensPSRV_Le_Lock", objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 191975
    
    objPedidoVenda_BD.iFilialEmpresaFaturamento = objPedVenda_BD.iFilialEmpresaFaturamento
    objPedidoVenda_BD.lCliente = objPedVenda_BD.lCliente
    objPedidoVenda_BD.iFilial = objPedVenda_BD.iFilial
    
'    '====BLOQUEIOS ==============================================
'
'    lErro = PedidoServico_AtualizaObsBloq_Trans(objPedidoVenda)
'    If lErro <> SUCESSO Then gError 191976
'
'    'remove da colecao de bloqueios os que já estavam no BD
'    'incluindo-os na colecao temporaria
'    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
'
'        If objBloqueioPV.iSequencial = 0 Then Exit For
'
'        Call colBloqBD.Add(objBloqueioPV)
'        Call objPedidoVenda.colBloqueiosPV.Remove(1)
'
'    Next
'
'    '??? verificar se pedido tem bloqueio por atraso liberado
'
'    'Verifica se o pedido tem credito aprovado
'    lErro = CF("BloqueiosPedidoSRV_Atraso_Testa", objPedidoVenda, iAtraso)
'    If lErro <> SUCESSO Then gError 191977
'
'    '??? se atraso ainda nao foi liberado, exclui bloqueios de atraso, se existirem
'
'    'Se o bloqueio por atraso ainda não foi aprovado
'    If iAtraso <> BLOQUEIO_POR_ATRASO_LIBERADO Then
'
'        lErro = CF("BloqueiosPedidoSRV_Exclui_Tipo", objPedidoVenda, BLOQUEIO_DIAS_ATRASO)
'        If lErro <> SUCESSO Then gError 191978
'
'        '??? chama BloqueioDiasAtraso_Adiciona
'
'        lErro = CF("BloqueioDiasAtraso_Adiciona", objPedidoVenda)
'        If lErro <> SUCESSO Then gError 191979
'
'        'Inclui no BD um bloqueio de dias de atraso, se necessario
'
'    End If
'
'    lErro = CF("BloqueiosPedidoSRV_Le", objPedidoVenda_BD)
'    If lErro <> SUCESSO Then gError 191980
'
'    objCliente.lCodigo = objPedidoVenda.lCliente
'    'le os dados do cliente
'    lErro = CF("Cliente_Le", objCliente)
'    If lErro <> SUCESSO And lErro <> 12293 Then gError 191981
'
'    If objCliente.iAtivo <> CLIENTE_ATIVO Then gError 191982
'
'    'tulio130103
'    'se o cliente nao usar o bloqueio de credito, bypassar os testes
'    If gobjCRFAT.iCreditoVerificaLimite = VERIFICAR_LIMITE_CREDITO Then
'
'        bPossuiBloqueioCredAtivo = False
'
'        For Each objBloqueioPV_BD In objPedidoVenda_BD.colBloqueiosPV
'            If objBloqueioPV_BD.iTipoDeBloqueio = BLOQUEIO_CREDITO And objBloqueioPV_BD.dtDataLib = DATA_NULA Then
'                bPossuiBloqueioCredAtivo = True
'            End If
'        Next
'
'        If bPossuiBloqueioCredAtivo Then
'            'Verifica se o pedido tem credito aprovado
'            lErro = CF("BloqueiosPSRV_Credito_Aprovado_Testa", objPedidoVenda, iCreditoAprovado)
'            If lErro <> SUCESSO Then gError 191983
'
'            bAdicionaBloqueio = False
'
'            'Se o crédito ainda não foi aprovado
'            If (iCreditoAprovado <> BLOQUEIO_CREDITO_LIBERADO) Then bAdicionaBloqueio = True
'
'        End If
'
'        'Se Valor Total aumentou faz adição (condicional) de bloqueio de crédito
'        If (Not bAdicionaBloqueio) And (objPedidoVenda.dValorTotal - objPedidoVenda_BD.dValorTotal) > DELTA_VALORMONETARIO Then bAdicionaBloqueio = True
'
'        If bAdicionaBloqueio Then
'            'Inclui no BD um bloqueio de Credito, se o cliente não tiver crédito
'            lErro = CF("BloqueioCredito_Adiciona", objPedidoVenda)
'            If lErro <> SUCESSO Then gError 191984
'
'        End If
'        If (iCreditoAprovado <> BLOQUEIO_CREDITO_LIBERADO) Then
'            'Exclui bloqueios de Credito com data de liberacao = DATA_NULA
'            lErro = CF("BloqueiosPSRV_Exclui_Credito", objPedidoVenda)
'            If lErro <> SUCESSO Then gError 191985
'            bAdicionaBloqueio = True
'        End If
'
'    End If
'
'    objBloqueioPV_BD.lPedidoDeVendas = objPedidoVenda.lCodigo
'    objBloqueioPV_BD.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
'
'    'Lê os bloqueio de estoque correspondente ao pedido que não esteja liberado
'    lErro = CF("BloqueioEstoquePSRV_Le", objBloqueioPV_BD)
'    If lErro <> SUCESSO And lErro <> 191706 Then gError 191986
'
'    If lErro = SUCESSO Then bAchouBloqueioBD = True
'    If lErro = 46032 Then bAchouBloqueioBD = False
'
'    'Se encontrou e no BD existe bloqueio total
'    If bAchouBloqueioBD And objBloqueioPV_BD.iTipoDeBloqueio = BLOQUEIO_TOTAL Then
'
'        'Avisa que as reservas serão canceladas
'        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_BLOQUEIO_TOTAL2")
'        If vbMsgRes = vbNo Then gError 191987
'
'        'Limpa as reservas
'        For Each objItemPV In objPedidoVenda.colItensPedido
'            Set objItemPV.ColReserva = New ColReserva
'        Next
'
'        iIndice = 0
'
'        'Exclui o Bloqueio total da coleção de bloqueios vinda da Tela
'        For Each objBloqueioPV1 In objPedidoVenda.colBloqueiosPV
'            iIndice = iIndice + 1
'            If objBloqueioPV1.iTipoDeBloqueio = BLOQUEIO_TOTAL Then
'                objPedidoVenda.colBloqueiosPV.Remove iIndice
'                Exit For
'            End If
'        Next
'
'    Else 'nao encontrou bloqueio TOTAL de reserva
'
'        'Adiciona um Bloqueio na coleção de Bloqueio de acordo com o a quantidade que falta reservar
'        lErro = CF("BloqueioReserva_Adiciona2", objPedidoVenda, objBloqueioPV_BD, objBloqueioPV, iBloqueioEstoqueAdicionado, dPedidoFaltaReservar)
'        If lErro <> SUCESSO Then gError 191988
'
'        'Verifica se na colecao de bloqueios existe algum bloqueio total
'        For Each objBloqueioPV1 In objPedidoVenda.colBloqueiosPV
'            If objBloqueioPV1.iTipoDeBloqueio = BLOQUEIO_TOTAL Then
'                bPossuiBloqueioTotal = True
'                Exit For
'            End If
'        Next
'
'        If bAchouBloqueioBD And (dPedidoFaltaReservar = 0 Or objBloqueioPV.iTipoDeBloqueio <> objBloqueioPV_BD.iTipoDeBloqueio Or bPossuiBloqueioTotal) Then
'            'Libera o Bloqueios de Estoque de Pedido
'            lErro = CF("BloqueioEstoquePSRV_Libera", objBloqueioPV_BD)
'            If lErro <> SUCESSO Then gError 191989
'
'        End If
'
'    End If
'
'    'copiar observacoes de bloqueios automaticos que foram recriados
'    Call CF("PedidoVenda_BloqAutoAtualiza", objPedidoVenda, colBloqBD)

    lErro = CF("BloqueiosGen_Altera", MAPBLOQGEN_TIPOTELA_PEDIDOSRV, objPedidoVenda)
    If lErro <> SUCESSO Then gError 191990
'
'    'Grava os bloqueios do Pedido de Venda
'    lErro = CF("BloqueiosPedidoSRV_Grava", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 191990
'    '=========FIM BLOQUEIOS=============================================
        
    'Altera estatística de quantidade pedida do Produto
    lErro = ProdutoFilial_Altera_QuantPedidaSRV(objPedidoVenda, objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 191991

    'Passa as informações de objPedidoVenda_BD para objPedVenda_BD
    For Each objItemPV_BD In objPedidoVenda_BD.colItensPedido
        With objItemPV_BD
        objPedVenda_BD.colItensPedido.Add .iFilialEmpresa, .lCodPedido, .iItem, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .ColReserva, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, .dQuantAFaturar, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, .lNumIntDoc
        End With
    Next
    
    For Each objItemPV_BD In objPedidoVenda_BD.colItensPedidoSRV
        With objItemPV_BD
        objPedVenda_BD.colItensPedidoSRV.Add .iFilialEmpresa, .lCodPedido, .iItem, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .ColReserva, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, .dQuantAFaturar, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, .lNumIntDoc
        End With
    Next
    
    'Verifica se a quantidade faturada do Item do BD é a mesma da tela
    lErro = ItensPSRV_QuantFaturadaAlterada(objPedidoVenda, objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 191992

    'Altera os dados do Pedido no BD
    lErro = CF("PedidoServico_Grava_BD", objPedidoVenda, alComando(0))
    If lErro <> SUCESSO Then gError 191993

    'Atualiza as comissões do Pedido passado
    lErro = CF("ComissoesPedidoSRV_Altera", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191994

    'Atualiza as parcelas do Pedido de Venda
    lErro = CF("ParcelasPedidoSRV_Altera", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191995
    
    'Insere os itens do Pedido no BD (peca)
    lErro = CF("ItensPSRV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191996
    
    'Insere os itens do Pedido no BD (servicos)
    lErro = CF("ItensPSRV_Grava1", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191997

'    'Lê as reservas do Pedido existentes no BD
'    lErro = CF("ReservasPedido_Le", objPedidoVenda_BD)
'    If lErro <> SUCESSO Then gError 191998
'
'    'Cria um Movimento de Estoque de Cancelamento de Reserva
'    lErro = CF("MovEstoqueExclusao_Cria", objMovEstoque, objPedidoVenda_BD)
'    If lErro <> SUCESSO Then gError 191999
'
'    'Grava o Movimento de Estoque de exclusao de reservas
'    lErro = CF("MovEstoque_Grava", objMovEstoque)
'    If lErro <> SUCESSO Then gError 193000
'
'    'Exclui reservas da tabela de reservas e inclui as vindas da tela
'    lErro = CF("ReservasPSRV_Altera", objPedidoVenda, alComando(1), alComando(2))
'    If lErro <> SUCESSO Then gError 193001
'
'    'Cria o Movimento de Estoque para inclusão das reservas
'    lErro = CF("MovEstoque_Cria", objMovEstoque, objPedidoVenda)
'    If lErro <> SUCESSO Then gError 193002
'
'    'Grava o Movimento de Estoque
'    lErro = CF("MovEstoque_Grava", objMovEstoque)
'    If lErro <> SUCESSO Then gError 193003

    'Remove de colItens do BD os itens presente na colItens da tela
    For Each objItemPV In objPedidoVenda.colItensPedido
        For iIndice = objPedidoVenda_BD.colItensPedido.Count To 1 Step -1
            If objItemPV.lNumIntDoc = objPedidoVenda_BD.colItensPedido(iIndice).lNumIntDoc Then
                objPedidoVenda_BD.colItensPedido.Remove iIndice
                Exit For
            End If
        Next
    Next

    For Each objItemPV In objPedidoVenda.colItensPedidoSRV
        For iIndice = objPedidoVenda_BD.colItensPedidoSRV.Count To 1 Step -1
            If objItemPV.lNumIntDoc = objPedidoVenda_BD.colItensPedidoSRV(iIndice).lNumIntDoc Then
                objPedidoVenda_BD.colItensPedidoSRV.Remove iIndice
                Exit For
            End If
        Next
    Next


    'Verifica para cada item se existe vínculo com um item de Ordem de Produção
    lErro = CF("ItensPSRV_ItensOP_Vinculados", objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 193004

    'Exclui os ítens de Pedido de venda que foram excluídos na tela
    lErro = CF("ItensPSRV_Exclui", objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 193005

    PedidoSRV_Altera_BD = SUCESSO

    Exit Function

Erro_PedidoSRV_Altera_BD:

    PedidoSRV_Altera_BD = gErr

    Select Case gErr

        Case 191973, 191975 To 191981, 191983 To 191999, 193000 To 193005

        Case 191974
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SERVICO_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)

        Case 191982
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INATIVO", gErr, objPedidoVenda.lCliente)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193006)

    End Select

    Exit Function

End Function
'
'Public Function BloqueiosPedidoSRV_Exclui_Tipo(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal iTipoBloqueio As Integer) As Long
''Exclui os Bloqueios do tipo recebido como parâmetro não liberados ligados ao Pedido de Venda passado
''Função Genérica
'
'Dim lErro As Long
'Dim lComando As Long
'Dim lComando1 As Long
'Dim iTipoBloqAux As Integer
'
'On Error GoTo Erro_BloqueiosPedidoSRV_Exclui_Tipo
'
'    'Abre os comandos
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 191691
'
'    lComando1 = Comando_Abrir()
'    If lComando1 = 0 Then gError 191692
'
'    'Busca os Bloqueios de Crédito ligados ao Pedido
'    lErro = Comando_ExecutarPos(lComando, "SELECT TipoDeBloqueio FROM BloqueiosPedidoSRV WHERE FilialEmpresa = ? AND PedidoSRV = ? AND TipoDeBloqueio = ? AND DataLib = ?", 0, iTipoBloqAux, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, iTipoBloqueio, DATA_NULA)
'    If lErro <> AD_SQL_SUCESSO Then gError 191693
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191694
'
'    'Para cada Bloqueio encontrado
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        'Exclui o bloqueio da tabela de Bloqueios de Pedido de Venda
'        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM BloqueiosPedidoSRV", lComando)
'        If lErro <> AD_SQL_SUCESSO Then gError 191695
'
'        'Busca o próximo
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191696
'
'    Loop
'
'    'Fecha os comandos
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    BloqueiosPedidoSRV_Exclui_Tipo = SUCESSO
'
'    Exit Function
'
'Erro_BloqueiosPedidoSRV_Exclui_Tipo:
'
'    BloqueiosPedidoSRV_Exclui_Tipo = gErr
'
'    Select Case gErr
'
'        Case 191691, 191692
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 191693, 191694, 191696
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'
'        Case 191695
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BLOQUEIOPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191696)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    Exit Function
'
'End Function
'
'Function BloqueiosPSRV_Exclui_Credito(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Exclui os Bloqueios de Creditos não liberados ligados ao Pedido de Venda passado
'
''Função alterada por leo em 26/02/02. Agora só chama uma função genérica que exclui os bloqueios do tipo passado como parâmetro
'Dim lErro As Long
'
'On Error GoTo Erro_BloqueiosPSRV_Exclui_Credito
'
'    'Função que exclui os Bloqueios do Tipo passada como parâmetro.
'    lErro = CF("BloqueiosPedidoSRV_Exclui_Tipo", objPedidoVenda, BLOQUEIO_CREDITO)
'    If lErro <> SUCESSO Then gError 94388
'
'    BloqueiosPSRV_Exclui_Credito = SUCESSO
'
'    Exit Function
'
'Erro_BloqueiosPSRV_Exclui_Credito:
'
'    BloqueiosPSRV_Exclui_Credito = gErr
'
'    Select Case gErr
'
'        Case 94388
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152557)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function BloqueioEstoquePSRV_Libera(ByVal objBloqueioPV As ClassBloqueioPV) As Long
''Libera o Bloqueio de Pedido de Venda passado em objBloqueioPV
'
'Dim lErro As Long
'Dim lComando As Long
'Dim lComando1 As Long
'Dim iFilialEmpresa As Integer
'
'On Error GoTo Erro_BloqueioEstoquePSRV_Libera
'
'    'Abre os comandos
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 191708
'
'    lComando1 = Comando_Abrir()
'    If lComando1 = 0 Then gError 191709
'
'    'Lê o bloqueio de Pedido de Venda
'    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa FROM BloqueiosPedidoSRV WHERE FilialEmpresa = ? AND PedidoSRV = ? AND Sequencial = ? AND (TipoDeBloqueio = ? OR TipoDeBloqueio = ?) AND DataLib = ?", 0, iFilialEmpresa, objBloqueioPV.iFilialEmpresa, objBloqueioPV.lPedidoDeVendas, objBloqueioPV.iSequencial, BLOQUEIO_PARCIAL, BLOQUEIO_NAO_RESERVA, DATA_NULA)
'    If lErro <> AD_SQL_SUCESSO Then gError 191710
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191711
'
'    If lErro = AD_SQL_SUCESSO Then
'
'        'Data e responsável pela liberação do Bloqueio
'        objBloqueioPV.dtDataLib = gdtDataHoje
'        objBloqueioPV.sResponsavelLib = BLOQUEIO_AUTO_RESP
'        objBloqueioPV.sCodUsuarioLib = gsUsuario
'
'        'Libera o bloqueio
'        lErro = Comando_ExecutarPos(lComando1, "UPDATE BloqueiosPedidoSRV SET DataLib = ?, ResponsavelLib = ?, CodUsuarioLib = ? ", lComando, objBloqueioPV.dtDataLib, objBloqueioPV.sResponsavelLib, objBloqueioPV.sCodUsuarioLib)
'        If lErro <> AD_SQL_SUCESSO Then gError 191712
'    End If
'
'    'Finaliza os comandos
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    BloqueioEstoquePSRV_Libera = SUCESSO
'
'    Exit Function
'
'Erro_BloqueioEstoquePSRV_Libera:
'
'    BloqueioEstoquePSRV_Libera = gErr
'
'    Select Case gErr
'
'        Case 191708, 191709
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 191710, 191711
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIO_PSRV", gErr, objBloqueioPV.iFilialEmpresa, objBloqueioPV.lPedidoDeVendas, objBloqueioPV.iSequencial)
'
'        Case 191712
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_BLOQUEIO_PSRV", gErr, objBloqueioPV.iFilialEmpresa, objBloqueioPV.lPedidoDeVendas, objBloqueioPV.iSequencial)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191713)
'
'    End Select
'
'    'Fecha os comandos
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    Exit Function
'
'End Function

Function ProdutoFilial_Altera_QuantPedidaSRV(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objPedidoVenda_Anterior As ClassPedidoDeVenda) As Long
'Atualiza a quantidade pedida em produtos filial na alteração do pedido de servico passado

Dim lErro As Long
Dim objPedidoVenda_BD As New ClassPedidoDeVenda

On Error GoTo Erro_ProdutoFilial_Altera_QuantPedidaSRV

    objPedidoVenda_BD.lNumIntDoc = objPedidoVenda.lNumIntDoc
    objPedidoVenda_BD.lCodigo = objPedidoVenda.lCodigo
    objPedidoVenda_BD.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
    
    'Lê o Pedido de Venda no BD
    lErro = CF("PedidoServico_Le_Itens", objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 191743
    
    'Subtrai a quantidade pedida dos itens lidos do BD
    lErro = ProdutoFilial_Subtrai_QuantPedidaSRV(objPedidoVenda_BD, objPedidoVenda_Anterior.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 191744
    
    'Adiciona a quantidade pedida dos itens da tela
    lErro = ProdutoFilial_Adiciona_QuantPedidaSRV(objPedidoVenda)
    If lErro <> SUCESSO Then gError 191745
    
    ProdutoFilial_Altera_QuantPedidaSRV = SUCESSO
    
    Exit Function
    
Erro_ProdutoFilial_Altera_QuantPedidaSRV:

    ProdutoFilial_Altera_QuantPedidaSRV = gErr
    
    Select Case gErr
    
        Case 191743, 191744, 191745
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191746)
            
    End Select
    
    Exit Function
    
End Function

Private Function ProdutoFilial_Subtrai_QuantPedidaSRV(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal iFilialEmpresa As Integer) As Long
'Atualiza Subtraindo a quantPedida do Produto em UM de Venda na tabela ProdutosFilial
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objItemPedVenda As ClassItemPedido
 
On Error GoTo Erro_ProdutoFilial_Subtrai_QuantPedidaSRV

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191719
    Next
    
    For Each objItemPedVenda In objPedidoVenda.colItensPedido
    
        'Subtrai a quantidade pedida dos itens lidos do BD
        lErro = ProdutoFilial_Subtrai_QuantPedidaSRV1(alComando, objItemPedVenda, objPedidoVenda.iFilialEmpresa)
        If lErro <> SUCESSO Then gError 191720
    
    Next
    
    For Each objItemPedVenda In objPedidoVenda.colItensPedidoSRV
    
        'Subtrai a quantidade pedida dos itens lidos do BD
        lErro = ProdutoFilial_Subtrai_QuantPedidaSRV1(alComando, objItemPedVenda, objPedidoVenda.iFilialEmpresa)
        If lErro <> SUCESSO Then gError 191721
    
    Next
    
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ProdutoFilial_Subtrai_QuantPedidaSRV = SUCESSO
    
    Exit Function
    
Erro_ProdutoFilial_Subtrai_QuantPedidaSRV:

    ProdutoFilial_Subtrai_QuantPedidaSRV = gErr
    
    Select Case gErr
    
        Case 191719
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                        
        Case 191720, 191721
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191722)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function ProdutoFilial_Subtrai_QuantPedidaSRV1(alComando() As Long, ByVal objItemPedVenda As ClassItemPedido, ByVal iFilialEmpresa As Integer) As Long
'Atualiza Subtraindo a quantPedida do Produto em UM de Venda na tabela ProdutosFilial
 
Dim lErro As Long
Dim dQuantPedida As Double
Dim dQuantItem As Double
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim sProduto As String
 
On Error GoTo Erro_ProdutoFilial_Subtrai_QuantPedidaSRV1

    sProduto = objItemPedVenda.sProduto
    dQuantItem = (objItemPedVenda.dQuantidade - objItemPedVenda.dQuantCancelada - objItemPedVenda.dQuantFaturada)

    objProduto.sCodigo = sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 191723
    
    If lErro <> SUCESSO Then gError 191724

    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPedVenda.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 191725
    
    'Calcula a quantidade a ser Subtraida a quantidade pedida do produto
    dQuantItem = dQuantItem * dFator
                
    'Busca o Registro de Produto do Item e FiliaisProduto
    lErro = Comando_ExecutarPos(alComando(0), "SELECT QuantPedida FROM ProdutosFilial WHERE Produto = ? AND FilialEmpresa = ?", 0, dQuantPedida, sProduto, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 191726

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191727
    
    If lErro <> AD_SQL_SUCESSO Then gError 191728
                  
    'Atualiza a quantidade pedida em ProdutosFilial para o Produto
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE ProdutosFilial SET QuantPedida = QuantPedida - ?", alComando(0), dQuantItem)
    If lErro <> AD_SQL_SUCESSO Then gError 191729
    
    ProdutoFilial_Subtrai_QuantPedidaSRV1 = SUCESSO
    
    Exit Function
    
Erro_ProdutoFilial_Subtrai_QuantPedidaSRV1:

    ProdutoFilial_Subtrai_QuantPedidaSRV1 = gErr
    
    Select Case gErr
    
        Case 191723, 191725
        
        Case 191724
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)
    
        Case 191726, 191727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err, iFilialEmpresa, sProduto)
        
        Case 191728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", Err, sProduto, iFilialEmpresa)
        
        Case 191729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PRODUTOSFILIAL", Err, iFilialEmpresa, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 191730)

    End Select
    
    Exit Function
    
End Function

Private Function ProdutoFilial_Adiciona_QuantPedidaSRV(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Atualiza adicionando a quantPedida do Produto em UM de Venda na tabela ProdutosFilial
 
Dim lErro As Long
Dim dQuantPedida As Double
Dim dQuantItem As Double
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objItemPedVenda As ClassItemPedido
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim iContador As Integer
Dim iQuantTotal As Integer

Dim sProduto As String
Dim sUnidadeMed As String
 
On Error GoTo Erro_ProdutoFilial_Adiciona_QuantPedidaSRV

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191739
    Next
    
    For Each objItemPedVenda In objPedidoVenda.colItensPedido
    
        lErro = ProdutoFilial_Adiciona_QuantPedidaSRV1(alComando(), objItemPedVenda, objPedidoVenda)
        If lErro <> SUCESSO Then gError 191740
    
    Next
    
    For Each objItemPedVenda In objPedidoVenda.colItensPedidoSRV
    
        lErro = ProdutoFilial_Adiciona_QuantPedidaSRV1(alComando(), objItemPedVenda, objPedidoVenda)
        If lErro <> SUCESSO Then gError 191741
    
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ProdutoFilial_Adiciona_QuantPedidaSRV = SUCESSO
    
    Exit Function
    
Erro_ProdutoFilial_Adiciona_QuantPedidaSRV:

    ProdutoFilial_Adiciona_QuantPedidaSRV = gErr
    
    Select Case gErr
    
        Case 191739
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 191740, 191741
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 191742)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function ProdutoFilial_Adiciona_QuantPedidaSRV1(alComando() As Long, ByVal objItemPedVenda As ClassItemPedido, ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Atualiza adicionando a quantPedida do Produto em UM de Venda na tabela ProdutosFilial
 
Dim lErro As Long
Dim dQuantPedida As Double
Dim dQuantItem As Double
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim sProduto As String
Dim sUnidadeMed As String
 
On Error GoTo Erro_ProdutoFilial_Adiciona_QuantPedidaSRV1

    sProduto = objItemPedVenda.sProduto
    dQuantItem = (objItemPedVenda.dQuantidade - objItemPedVenda.dQuantCancelada - objItemPedVenda.dQuantFaturada)
    
    objProduto.sCodigo = sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 191731
    
    If lErro <> SUCESSO Then gError 191732

    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPedVenda.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 191733
    
    'Calcula a quantidade a ser adicionada a quantidade pedida do produto
    dQuantItem = dQuantItem * dFator
                
    'Busca o Registro de Produto do Item e FiliaisProduto na Filial de Faturamento
    lErro = Comando_ExecutarPos(alComando(0), "SELECT QuantPedida FROM ProdutosFilial WHERE Produto = ? AND FilialEmpresa = ?", 0, dQuantPedida, sProduto, objPedidoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 191734

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191735
    
    If lErro <> AD_SQL_SUCESSO Then gError 191736
                  
    'Atualiza a quantidade pedida em ProdutosFilial para o Produto e Filial
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE ProdutosFilial SET QuantPedida = QuantPedida + ?", alComando(0), dQuantItem)
    If lErro <> AD_SQL_SUCESSO Then gError 191737
        
    ProdutoFilial_Adiciona_QuantPedidaSRV1 = SUCESSO
    
    Exit Function
    
Erro_ProdutoFilial_Adiciona_QuantPedidaSRV1:

    ProdutoFilial_Adiciona_QuantPedidaSRV1 = gErr
    
    Select Case gErr
    
        Case 191731, 191733
    
        Case 191732
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                        
        Case 191734, 191735
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", gErr, objItemPedVenda.iFilialEmpresa, objItemPedVenda.sProduto)
        
        Case 191736
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE_FILIALFATURAMENTO", gErr, objItemPedVenda.sProduto, objItemPedVenda.iFilialEmpresa)
        
        Case 191737
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PRODUTOSFILIAL", gErr, objItemPedVenda.iFilialEmpresa, objItemPedVenda.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191738)

    End Select
    
    Exit Function
    
End Function

Function ComissoesPedidoSRV_Altera(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Atualiza as comissões do Pedido de Venda passado

Dim lErro As Long

On Error GoTo Erro_ComissoesPedidoSRV_Altera

    'Exclui do BD as comissões do Pedido
    lErro = CF("ComissoesPedidoSRV_Exclui", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191747

    'Inclui no BD as comissões em objPedidoVenda
    lErro = CF("ComissoesPedidoSRV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191748

    ComissoesPedidoSRV_Altera = SUCESSO

    Exit Function

Erro_ComissoesPedidoSRV_Altera:

    ComissoesPedidoSRV_Altera = gErr

    Select Case gErr

        Case 191747, 191748

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191749)

    End Select

    Exit Function

End Function

Function ComissoesPedidoSRV_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui do BD as Comissões do Pedido de Venda passado

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iCodVendedor As Integer

On Error GoTo Erro_ComissoesPedidoSRV_Exclui

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 191750

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 191751

    'Busca as comissões vinculadas ao Pedido de Venda passado
    lErro = Comando_ExecutarPos(lComando, "SELECT CodVendedor FROM ComissoesPedidoSRV WHERE FilialEmpresa =? AND PedidoSRV = ?", 0, iCodVendedor, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 191752

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191753

    'Para cada comissão encontrada
    Do While lErro = AD_SQL_SUCESSO

        'Exclui a comissão da tabela de Comissões de Pedido de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ComissoesPedidoSRV", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 191754

        'Busca a próxima
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191755

    Loop

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ComissoesPedidoSRV_Exclui = SUCESSO

    Exit Function

Erro_ComissoesPedidoSRV_Exclui:

    Select Case gErr

        Case 191750, 191751
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 191752, 191753, 191755
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)

        Case 191754
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMISSOESPEDIDOSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191756)

    End Select

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function ParcelasPedidoSRV_Altera(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Atualiza as pracelas do Pedido de Servico Passdados

Dim lErro As Long

On Error GoTo Erro_ParcelasPedidoSRV_Altera

    'Exclui as Parcelas ligadas ao Pedido de Venda
    lErro = CF("ParcelasPedidoSRV_Exclui", objPedidoVenda)
    If lErro <> AD_SQL_SUCESSO Then gError 191757

    'Inclui as novas parcelas do pedido de Venda
    lErro = CF("ParcelasPedidoSRV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 191758

    ParcelasPedidoSRV_Altera = SUCESSO

    Exit Function

Erro_ParcelasPedidoSRV_Altera:

    ParcelasPedidoSRV_Altera = gErr

    Select Case gErr

        Case 191757, 191758

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191759)

    End Select

    Exit Function

End Function

Function ParcelasPedidoSRV_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui do BD as Parcelas do Pedido

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iNumParcela As Integer

On Error GoTo Erro_ParcelasPedidoSRV_Exclui

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 191760

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 191761

    'Busca as Parcelas do Pedido de Venda
    lErro = Comando_ExecutarPos(lComando, "SELECT NumParcela FROM ParcelasPedidoSRV WHERE FilialEmpresa = ? AND PedidoSRV = ?", 0, iNumParcela, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 191762

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191763

    'Para cada parcela encontrada
    Do While lErro <> AD_SQL_SEM_DADOS

        'Exclui a parcela da tabela de parcelas de Pedido de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ParcelasPedidoSRV", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 191764

        'Busca a Próxima
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191765

    Loop

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ParcelasPedidoSRV_Exclui = SUCESSO

    Exit Function

Erro_ParcelasPedidoSRV_Exclui:

    ParcelasPedidoSRV_Exclui = gErr

    Select Case gErr

        Case 191760, 191761
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 191762, 191763, 191765
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)

        Case 191764
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PARCELASPEDIDOSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191766)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Function PedSRV_QuantNaoCancelada(ByVal objPedidoVenda As ClassPedidoDeVenda, dQuantNaoCancelada As Double) As Long
'Retorna a quantidade não cancelada em todos o Pedido

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim iStatus As Integer

On Error GoTo Erro_PedSRV_QuantNaoCancelada

    'Zera o acumulador
    dQuantNaoCancelada = 0

    'Para cada item acumula a Quantidade não cancelada
    For Each objItemPV In objPedidoVenda.colItensPedido
        objItemPV.dQuantAFaturar = objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada
        dQuantNaoCancelada = dQuantNaoCancelada + objItemPV.dQuantAFaturar
    Next

    'Para cada item acumula a Quantidade não cancelada
    For Each objItemPV In objPedidoVenda.colItensPedidoSRV
        objItemPV.dQuantAFaturar = objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada
        dQuantNaoCancelada = dQuantNaoCancelada + objItemPV.dQuantAFaturar
    Next

    PedSRV_QuantNaoCancelada = SUCESSO

    Exit Function

Erro_PedSRV_QuantNaoCancelada:

    PedSRV_QuantNaoCancelada = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191910)

    End Select

    Exit Function

End Function

Function PedidoServico_Baixar(ByVal objPedidoVenda As ClassPedidoDeVenda, Optional ByVal iStatus As Integer = STATUS_BAIXADO) As Long
'transfere os dados do pedido para as tabelas de pedidos baixados
'IMPORTANTE: Tem que ser chamado dentro de TRANSACAO

Dim lErro As Long
Dim lNumIntDoc As Long
Dim tPedidoVenda As typePedidoVenda
Dim tItemPedido As typeItemPedido
Dim objBloqueioPV As New ClassBloqueioPV
Dim objMovEstoque As New ClassMovEstoque
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iStatusBD As Integer

On Error GoTo Erro_PedidoServico_Baixar
        
    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191927
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Status FROM PedidoServico WHERE FilialEmpresa = ? AND Codigo = ?", 0, iStatusBD, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 191914

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191915
    
    If lErro = AD_SQL_SEM_DADOS Then gError 191916

    'Faz o "lock" no Pedido de Venda
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 191917

    lErro = Comando_ExecutarPos(alComando(2), "UPDATE PedidoServico SET Status = ?", alComando(1), iStatus)
    If lErro <> AD_SQL_SUCESSO Then gError 191918
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PedidoServico_Baixar = SUCESSO

    Exit Function

Erro_PedidoServico_Baixar:

    PedidoServico_Baixar = gErr

    Select Case gErr
            
        Case 191911 To 191913, 191926
            
        Case 191914, 191915
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSERVICO", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)

        Case 191916
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SERVICO_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)

        Case 191917
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOSRV", gErr, objPedidoVenda.lCodigo)
            
        Case 191918
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOSERVICO", gErr, objPedidoVenda.lCodigo)
            
        Case 191927
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191919)
        
    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'Function ReservasPSRV_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Exclui as Reservas do pedido de Servico passado
''CHAMADA EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim iFilialEmpresa As Integer
'Dim alComando(1 To 2) As Long
'Dim iIndice As Integer
'
'On Error GoTo Erro_ReservasPSRV_Exclui
'
'    'Abre o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 191920
'    Next
'
'    'Lê as reservas do Pedido
'    lErro = Comando_ExecutarPos(alComando(1), "SELECT FilialEmpresa FROM Reserva WHERE FilialEmpresa = ? AND TipoDoc = ? AND DocOrigem = ?", 0, iFilialEmpresa, objPedidoVenda.iFilialEmpresa, TIPO_PEDIDO_SRV_COD, objPedidoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 191921
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 191922
'
'    'Para cada reserva lida
'    Do While lErro = AD_SQL_SUCESSO
'
'        'Exclui a Reserva do Pedido da tabela de Reservas
'        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM Reserva", alComando(1))
'        If lErro <> AD_SQL_SUCESSO Then gError 191923
'
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191924
'
'    Loop
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    ReservasPSRV_Exclui = SUCESSO
'
'    Exit Function
'
'Erro_ReservasPSRV_Exclui:
'
'    ReservasPSRV_Exclui = gErr
'
'    Select Case gErr
'
'        Case 191920
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 191921, 191922, 191924
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA1", gErr)
'
'        Case 191923
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RESERVASPSRV", gErr, objPedidoVenda.lCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191925)
'
'    End Select
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Private Function PedidoServico_AlteraVinculo_Orcamento(objPedidoVenda As ClassPedidoDeVenda)
'Vincula ou desvincula um Orçamento de Venda a um Pedido de Venda.

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim lComando As Long
Dim lComando2 As Long
Dim lNumIntPedVenda As Long
Dim lNumIntNFiscal As Long
Dim lCodigo As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_PedidoServico_AlteraVinculo_Orcamento

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 191928
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 191929
        
    'Se o usuário selecionou um Orcamento
    If objPedidoVenda.lNumIntOrcVenda <> 0 Then
    
        objOrcamentoVenda.lNumIntDoc = objPedidoVenda.lNumIntOrcVenda
    
        'Seleciona o Orçamento na Tela de OrcamentoVenda
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntPedSRV, NumIntNFiscal, FilialEmpresa, Codigo FROM OrcamentoSRV WHERE NumIntDoc = ?", _
        0, lNumIntPedVenda, lNumIntNFiscal, iFilialEmpresa, lCodigo, objOrcamentoVenda.lNumIntDoc)
        If lErro <> SUCESSO Then gError 191930
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191931
        
        'Se o Orçamento não existe ---> Erro.
        If lErro <> AD_SQL_SUCESSO Then gError 191932
                
        objOrcamentoVenda.lNumIntPedVenda = lNumIntPedVenda
        objOrcamentoVenda.lNumIntNFiscal = lNumIntNFiscal
        
        'Se o Orcamento estiver vinculado a uma Nota Fiscal ou a um pedido diferente ---> Erro.
        If objOrcamentoVenda.lNumIntPedVenda <> 0 And objOrcamentoVenda.lNumIntPedVenda <> objPedidoVenda.lNumIntDoc Then gError 191933
        If objOrcamentoVenda.lNumIntNFiscal <> 0 Then gError 191934
    
        'Vincula o Orçamento ao Pedido
        lErro = Comando_ExecutarPos(lComando2, "UPDATE OrcamentoSRV SET NumIntPedSRV = ?", lComando, objPedidoVenda.lNumIntDoc)
        If lErro <> SUCESSO Then gError 191935
        
    'Se o não foi seleciona Orçamento
    Else
        
        objOrcamentoVenda.lNumIntPedVenda = objPedidoVenda.lNumIntDoc
        
        'Tenta selecionar algum Orçamento Vinculado ao Pedido.
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntNFiscal, Codigo, FilialEmpresa FROM OrcamentoSRV WHERE NumIntPedSRV = ?", _
        0, lNumIntNFiscal, lCodigo, iFilialEmpresa, objOrcamentoVenda.lNumIntPedVenda)
        If lErro <> SUCESSO Then gError 191936
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191937
        
        'Se encontrou um Orçamento Vinculado ao Pedido
        If lErro = AD_SQL_SUCESSO Then
            
            'Desvincula o Orçamento do Pedido, pois o usuário está gravando um Pedido sem Orçamento.
            lErro = Comando_ExecutarPos(lComando2, "UPDATE OrcamentoSRV SET NumIntPedSRV = ?", lComando, objPedidoVenda.lNumIntOrcVenda)
            If lErro <> SUCESSO Then gError 191938
            
        End If
        
    End If
            
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    PedidoServico_AlteraVinculo_Orcamento = SUCESSO
    
    Exit Function
    
Erro_PedidoServico_AlteraVinculo_Orcamento:
    
    PedidoServico_AlteraVinculo_Orcamento = gErr
    
    Select Case gErr
                              
        Case 191928, 191929
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                                          
        Case 191930, 191931
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOSRV2", gErr, objOrcamentoVenda.lNumIntDoc)
            
        Case 191932
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_NAO_CADASTRADO", gErr, objOrcamentoVenda.lNumIntDoc)
            
        Case 191933
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_VINCULADO_PEDVENDA", gErr, lCodigo, iFilialEmpresa)
            
        Case 191934
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_VINCULADO_NFISCAL", gErr, lCodigo, iFilialEmpresa)
                       
        Case 191935, 191938
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ORCAMENTOSRV", gErr, lCodigo, iFilialEmpresa)
                       
        Case 191936, 191937
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOSRV1", gErr, objOrcamentoVenda.lNumIntPedVenda)
                       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191939)
     
    End Select
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
        
    Exit Function

End Function

Private Function ItensPSRV_QuantFaturadaAlterada(ByVal objPedVenda As ClassPedidoDeVenda, ByVal objPedidoVenda_BD As ClassPedidoDeVenda) As Long
'Para cada item verifica se a quantidade Faturada é a mesma que a do Banco de Dados

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim objItemPV_BD As ClassItemPedido
Dim iIndice As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_ItensPSRV_QuantFaturadaAlterada

    iIndice = 0

    'Para cada item
    For Each objItemPV In objPedVenda.colItensPedido

        iIndice = iIndice + 1
            
        'Se for um item já existente
        If objItemPV.lNumIntDoc <> 0 Then
            'Pesquisa na colecao de itens do Pedido Venda lido no BD
            For Each objItemPV_BD In objPedidoVenda_BD.colItensPedido
            
                If objItemPV_BD.lNumIntDoc = objItemPV.lNumIntDoc Then Exit For
                
            Next

            'Se não achar erro
            If objItemPV_BD.lNumIntDoc <> objItemPV.lNumIntDoc Then gError 191958
            
            'Verifica se a quantidade faturada foi alterada
            If objItemPV.dQuantFaturada <> objItemPV_BD.dQuantFaturada Then gError 191959
        
        End If
        
    Next

    iIndice = 0

    'Para cada item
    For Each objItemPV In objPedVenda.colItensPedidoSRV

        iIndice = iIndice + 1
            
        'Se for um item já existente
        If objItemPV.lNumIntDoc <> 0 Then
            'Pesquisa na colecao de itens do Pedido Venda lido no BD
            For Each objItemPV_BD In objPedidoVenda_BD.colItensPedidoSRV
            
                If objItemPV_BD.lNumIntDoc = objItemPV.lNumIntDoc Then Exit For
                
            Next

            'Se não achar erro
            If objItemPV_BD.lNumIntDoc <> objItemPV.lNumIntDoc Then gError 191960
            
            'Verifica se a quantidade faturada foi alterada
            If objItemPV.dQuantFaturada <> objItemPV_BD.dQuantFaturada Then gError 191961
        
        End If
        
    Next

    ItensPSRV_QuantFaturadaAlterada = SUCESSO

    Exit Function

Erro_ItensPSRV_QuantFaturadaAlterada:

    ItensPSRV_QuantFaturadaAlterada = Err

    Select Case gErr

        Case 191958, 191960
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPSRV_NAO_ENCONTRADO", gErr, objItemPV.lNumIntDoc)

        Case 191959, 191961
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_FATURADA_ALTERADA_SRV", gErr, iIndice, objPedVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 191962)

    End Select

    Exit Function

End Function

Function ItensPSRV_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui os itens presentes na coleção de itens do Pedido de Serviço passado

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objItemPV As ClassItemPedido
Dim iIndice As Integer
Dim alComando(1 To 4) As Long

On Error GoTo Erro_ItensPSRV_Exclui

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191966
    Next

    'Para cada item na coleção
    For Each objItemPV In objPedidoVenda.colItensPedido

        'Busca o ítem no BD
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM ItensPedidoSRV WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 191967

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 191968
        
        If lErro = AD_SQL_SEM_DADOS Then gError 191969 'Não encontrou

        'Exclui o item da tabela de itens de Pedido de Venda
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ItensPedidoSRV", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 191970

        lErro = CF("ItemPSRV_ExcluiTributacao", objItemPV, alComando(3), alComando(4))
        If lErro <> SUCESSO Then gError 191971

    Next

    'Para cada item na coleção
    For Each objItemPV In objPedidoVenda.colItensPedidoSRV

        'Busca o ítem no BD
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM ItensPedidoSRV WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 193479

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193480
        
        If lErro = AD_SQL_SEM_DADOS Then gError 193481 'Não encontrou

        'Exclui o item da tabela de itens de Pedido de Venda
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ItensPedidoSRV", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 193482

        lErro = CF("ItemPSRV_ExcluiTributacao", objItemPV, alComando(3), alComando(4))
        If lErro <> SUCESSO Then gError 193483

    Next

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensPSRV_Exclui = SUCESSO

    Exit Function

Erro_ItensPSRV_Exclui:

    ItensPSRV_Exclui = gErr

    Select Case gErr

        Case 191966
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 191967, 191968, 193479, 193480
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDOSRV2", gErr)

        Case 191969, 193481
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPSRV_NAO_ENCONTRADO", gErr, objPedidoVenda.lNumIntDoc)

        Case 191970, 193482
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPEDIDOSRV", gErr, objPedidoVenda.lNumIntDoc)

        Case 191971, 193483

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191972)

    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoServico_Baixar_Pedido(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
    
Dim lErro As Long
Dim alComando(4 To 27) As Long
Dim lTransacao As Long
Dim iIndice As Integer

On Error GoTo Erro_PedidoServico_Baixar_Pedido
    
    
    'Atualiza Subtraindo a quantPedida do Produto em UM de Venda na tabela ProdutosFilial
    lErro = ProdutoFilial_Subtrai_QuantPedidaSRV(objPedidoVenda, objPedidoVenda.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 193069
    
    'baixa o pedido
    lErro = CF("PedidoServico_Baixar", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193070
    
    PedidoServico_Baixar_Pedido = SUCESSO
    
    Exit Function
    
Erro_PedidoServico_Baixar_Pedido:

    PedidoServico_Baixar_Pedido = gErr
    
    Select Case gErr
        
        Case 193068
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 193069, 193070
    
        Case 193071
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193072)
            
    End Select

    Exit Function
    
End Function

Function PedidoServico_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui o Pedido de Venda passado

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 14) As Long
Dim iIndice As Integer
Dim objItemPV As ClassItemPedido
Dim objItemOP As New ClassItemOP
Dim objMovEstoque As New ClassMovEstoque
Dim colobjRC As New Collection
Dim objRC As ClassRequisicaoCompras
Dim sRequisicoes As String, lCodigoPC As Long

On Error GoTo Erro_PedidoServico_Exclui

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 193139

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 193140
    Next

    'Faz lock no Pedido
    lErro = CF("PedidoServico_Lock", objPedidoVenda, alComando(0))
    If lErro <> SUCESSO And lErro <> 191683 Then gError 193141

    If lErro <> SUCESSO Then gError 193142

    'Lê e faz lock nos itens do Pedido
    lErro = CF("ItensPSRV_Le_Lock", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193150
        
    lErro = ProdutoFilial_Subtrai_QuantPedidaSRV(objPedidoVenda, objPedidoVenda.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 193151
    
    'Para todos os itens de PedidoVenda
    For Each objItemPV In objPedidoVenda.colItensPedido

        'Verifica se item tem Quantidade Faturada positiva
        If objItemPV.dQuantFaturada > 0 Then gError 193152

'        'Verifica se item está vinculado a item de ordem de produção
'        lErro = CF("ItemOP_Le_ItemPSRV", objItemOP, objItemPV)
'        If lErro <> SUCESSO And lErro <> 191559 Then gError 193153
'
'        If lErro = SUCESSO Then gError 193154

    Next
    
    For Each objItemPV In objPedidoVenda.colItensPedidoSRV

        'Verifica se item tem Quantidade Faturada positiva
        If objItemPV.dQuantFaturada > 0 Then gError 193155

'        'Verifica se item está vinculado a item de ordem de servico
'        lErro = CF("ItemOP_Le_ItemPSRV", objItemOP, objItemPV)
'        If lErro <> SUCESSO And lErro <> 191559 Then gError 46146
'
'        If lErro = SUCESSO Then gError 46147

    Next
    
    
    'Verifica se existe RequisicaoCompra vinculada ao Pedido de Venda
    lErro = CF("PedidoServico_Busca_ReqCompra", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193156

    'Verifica se existe PedidoCompra vinculada ao Pedido de Venda
    lErro = CF("PedidoServico_Busca_PedidoCompra", objPedidoVenda, lCodigoPC)
    If lErro <> SUCESSO And lErro <> 193090 Then gError 193157

    If lErro = SUCESSO Then gError 193158

    'Verifica se existe requisicao associado ao pedido de venda
    lErro = CF("ItemPSRVItemRC_Le_CodigoPV", objPedidoVenda, colobjRC)
    If lErro <> SUCESSO Then gError 193159

    If colobjRC.Count > 0 Then
        For Each objRC In colobjRC
            sRequisicoes = sRequisicoes & "Filial=" & Str(objRC.iFilialEmpresa) & " Requisicao=" & Str(objRC.lCodigo) & Chr(10)
        Next
        gError 193160
    End If
'
'    'Lê as Reservas do pedido
'    lErro = CF("ReservasPedidoSRV_Le", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 193161
'
'    'Cria um movimento de estoque para a exclusão
'    lErro = CF("MovEstoqueExclusao_Cria", objMovEstoque, objPedidoVenda)
'    If lErro <> SUCESSO Then gError 193162
'
'    If objMovEstoque.colItens.Count > 0 Then
'        lErro = CF("MovEstoque_Grava", objMovEstoque)
'        If lErro <> SUCESSO Then gError 193163
'    End If
'
'    'Exclui as Reservas do Pedido
'    lErro = CF("ReservasPSRV_Exclui", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 193164

    'Exclui os bloqueios do Pedido
    lErro = CF("BloqueiosGen_Exclui", MAPBLOQGEN_TIPOTELA_PEDIDOSRV, objPedidoVenda)
    If lErro <> SUCESSO Then gError 193165

    'Exclui as comissões do Pedido
    lErro = CF("ComissoesPedidoSRV_Exclui", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193166

    'Exclui as Parcelas do Pedido
    lErro = CF("ParcelasPedidoSRV_Exclui", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193167

    'Exclui os itens do Pedido
    lErro = CF("ItensPSRV_Exclui", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193168

    'Caso exista vínculo, desvincula o Orçamento de Venda do Pedido a ser excluido.
    lErro = PedidoServico_Desvincula_Orcamento(objPedidoVenda) 'Por Leo em 29/04/02
    If lErro <> SUCESSO Then gError 193169
    
    'Exclui o Pedido de Venda do BD
    lErro = CF("PedidoServico_Exclui_BD", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193170
    
    'Diminui a Estatistica para os Produtos e Data de Emissao do Pedido de Venda Gerados em SldDiaFat e SldMesFat
    lErro = PedSRV_Exclui_Estatisticas(objPedidoVenda)
    If lErro <> SUCESSO Then gError 193171
    
    lErro = CF("InfoAdicionais_Exclui", TIPODOC_INFOADIC_PSRV, objPedidoVenda.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 193171
    
    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 193172

    lErro = CF("WorkFlow_Trata_Transacao", "SEV", objPedidoVenda.objTela, gsUsuario, objPedidoVenda)
    If lErro <> SUCESSO Then gError 178238
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoServico_Exclui = SUCESSO

    Exit Function

Erro_PedidoServico_Exclui:

    PedidoServico_Exclui = gErr

    Select Case gErr

        Case 193139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 193140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 193141, 193150, 193151, 193153, 193156, 193157, 193159, 193161 To 193171

        Case 193142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SERVICO_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)

        Case 193152, 193155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PSRV_FATURADO", gErr, objPedidoVenda.lCodigo, objItemPV.lNumIntDoc, objItemPV.dQuantFaturada)

        Case 193154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_OP_VINCULADO_ITEM_PSRV", gErr, objPedidoVenda.lCodigo, objItemPV.lNumIntDoc, objItemOP.lNumIntDoc)

        Case 193158
            Call Rotina_Erro(vbOKOnly, "ERRO_PSRV_VINCULADO_PEDIDOCOMPRA", gErr, lCodigoPC)
        
        Case 193160
            Call Rotina_Erro(vbOKOnly, "ERRO_PSRV_VINCULADO_REQCOMPRA", gErr, sRequisicoes)

        Case 193172
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193173)

    End Select

    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
'
'Function BloqueiosPedidoSRV_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Exclui os Bloqueios do Pedido de Servico passado
'
'Dim lErro As Long
'Dim iSequencial As Integer
'Dim alComando(1 To 2) As Long
'Dim iIndice As Integer
'
'On Error GoTo Erro_BloqueiosPedidoSRV_Exclui
'
'    'Abre os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 193098
'    Next
'
'    'Busca os Bloqueios do Pedido passado
'    lErro = Comando_ExecutarPos(alComando(1), "SELECT Sequencial FROM BloqueiosPedidoSRV WHERE FilialEmpresa = ? AND PedidoSRV = ?", 0, iSequencial, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 193099
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193100
'
'    'Para  cadabloqueio encontrado
'    Do While lErro = AD_SQL_SUCESSO
'        'Exclui o bloqueio do BD
'        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM BloqueiosPedidoSRV", alComando(1))
'        If lErro <> AD_SQL_SUCESSO Then gError 193101
'
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193102
'
'    Loop
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    BloqueiosPedidoSRV_Exclui = SUCESSO
'
'    Exit Function
'
'Erro_BloqueiosPedidoSRV_Exclui:
'
'    BloqueiosPedidoSRV_Exclui = gErr
'
'    Select Case gErr
'
'        Case 193098
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 193099, 193100, 193102
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'
'        Case 193101
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BLOQUEIOPEDIDOSRV", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193102)
'
'    End Select
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Function PedidoServico_Desvincula_Orcamento(ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Desvincula o Orçamento do Pedido de Servico a ser excluido.

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_PedidoServico_Desvincula_Orcamento
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193103
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 193104
    
    'Tenta buscar um Orçamento Vinculado ao Pedido de Venda
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM OrcamentoSRV WHERE NumIntPedSRV = ?", 0, lNumIntDoc, objPedVenda.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 193105

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193106
    
    If lErro = SUCESSO Then
    
        'Desvincula o Pedido da Tabela de Orçamentos.
        lErro = Comando_ExecutarPos(lComando1, "UPDATE OrcamentoSRV SET NumIntPedSRV = ?", lComando, 0)
        If lErro <> AD_SQL_SUCESSO Then gError 193107

    End If

    PedidoServico_Desvincula_Orcamento = SUCESSO

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

Erro_PedidoServico_Desvincula_Orcamento:

    PedidoServico_Desvincula_Orcamento = gErr

    Select Case gErr

        Case 193103, 193104
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 193105, 193106
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOSRV1", gErr, objPedVenda.lNumIntDoc)

        Case 193107
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ORCAMENTOSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193108)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function PedidoServico_Exclui_BD(ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Exclui o Pedido de Venda da tabela de Pedido de Vendas

Dim lErro As Long
Dim lNumIntDoc As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_PedidoServico_Exclui_BD

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 193109
    Next

    'Busca o Pedido no BD
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM PedidoServico WHERE FilialEmpresa = ? AND Codigo = ? ", 0, lNumIntDoc, objPedVenda.iFilialEmpresa, objPedVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 193110

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193111
    
    If lErro = AD_SQL_SEM_DADOS Then gError 193112

    'Exclui o Pedido da Tabela de Pedidos
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM PedidoServico ", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 193113

    lErro = CF("PedidoServico_ExcluiTributacao", objPedVenda)
    If lErro <> AD_SQL_SUCESSO Then gError 193114

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoServico_Exclui_BD = SUCESSO

    Exit Function

Erro_PedidoServico_Exclui_BD:

    PedidoServico_Exclui_BD = gErr

    Select Case gErr

        Case 193109
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 193110, 193111
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSERVICO", gErr, objPedVenda.iFilialEmpresa, objPedVenda.lCodigo)

        Case 193112
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SERVICO_NAO_CADASTRADO", gErr, objPedVenda.lCodigo)

        Case 193113
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PEDIDOSERVICO", gErr, objPedVenda.iFilialEmpresa, objPedVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193114)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PedSRV_Exclui_Estatisticas(ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Exclui a estatística dos Produtos Pedidos

Dim lErro As Long

On Error GoTo Erro_PedSRV_Exclui_Estatisticas
           
    'Exclui a estatistica do Dia do produto
    lErro = SldDiaFat_Exclui_EstPSRV(objPedVenda)
    If lErro <> SUCESSO Then gError 193136
    
    'Exclui a estatística do Mês do produto
    lErro = SldMesFat_Exclui_EstPSRV(objPedVenda)
    If lErro <> SUCESSO Then gError 193137

    PedSRV_Exclui_Estatisticas = SUCESSO
    
    Exit Function
    
Erro_PedSRV_Exclui_Estatisticas:
    
    PedSRV_Exclui_Estatisticas = gErr

    Select Case gErr
        
        Case 193136, 193137
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191138)

    End Select
    
    Exit Function

End Function

Private Function SldDiaFat_Exclui_EstPSRV(ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Grava a estatística do dia do Produto Pedido

Dim lErro As Long
Dim objItemPV As New ClassItemPedido
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_SldDiaFat_Exclui_EstPSRV
        
    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 193121
    Next
          
    'Para cada item do PedidoServico
    For Each objItemPV In objPedVenda.colItensPedido
        
        lErro = SldDiaFat_Exclui_EstPSRV1(alComando, objPedVenda, objItemPV)
        If lErro <> SUCESSO Then gError 193122
        
    Next
        
    'Para cada item do PedidoServico
    For Each objItemPV In objPedVenda.colItensPedidoSRV
        
        lErro = SldDiaFat_Exclui_EstPSRV1(alComando, objPedVenda, objItemPV)
        If lErro <> SUCESSO Then gError 193123
        
    Next
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    SldDiaFat_Exclui_EstPSRV = SUCESSO
    
    Exit Function
    
Erro_SldDiaFat_Exclui_EstPSRV:
    
    SldDiaFat_Exclui_EstPSRV = gErr

    Select Case gErr
    
        Case 193121
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 193122, 193123
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193124)

    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function SldDiaFat_Exclui_EstPSRV1(alComando() As Long, ByVal objPedVenda As ClassPedidoDeVenda, objItemPV As ClassItemPedido) As Long
'Grava a estatística do dia do Produto Pedido

Dim lErro As Long
Dim objSldDiaFat As New ClassSldDiaFat
Dim dQuantUMVenda As Double
Dim dFator As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_SldDiaFat_Exclui_EstPSRV1
        
        
    objProduto.sCodigo = objItemPV.sProduto
    
    'Lê o produto para pegar a UM_Venda
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 193115

    If lErro = 28030 Then gError 193116
        
    'Lê o fator de conversão da UM_Venda
    lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 193117

    'Converte a UM da quantidade
    dQuantUMVenda = ((objItemPV.dQuantidade - objItemPV.dQuantCancelada) * dFator)
    
    objSldDiaFat.iFilialEmpresa = giFilialEmpresa
    objSldDiaFat.dtData = objPedVenda.dtDataEmissao
    objSldDiaFat.sProduto = objItemPV.sProduto
    
    'Lê todos os atributos na tabela SldDiaFat e faz lock exclusive
    lErro = CF("SldDiaFat_Le_Lock", alComando(1), objSldDiaFat)
    If lErro <> SUCESSO And lErro <> 52729 Then gError 193118
                
    'Se encontrou
    If lErro = SUCESSO Then
        
        objSldDiaFat.dQuantPedida = objSldDiaFat.dQuantPedida - dQuantUMVenda
        
        'Atualiza a tabela de SldDiaFat
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE SldDiaFat SET QuantPedidaSRV =?", alComando(1), objSldDiaFat.dQuantPedida)
        If lErro <> AD_SQL_SUCESSO Then gError 193119
    
    End If
        
    SldDiaFat_Exclui_EstPSRV1 = SUCESSO
    
    Exit Function
    
Erro_SldDiaFat_Exclui_EstPSRV1:
    
    SldDiaFat_Exclui_EstPSRV1 = gErr

    Select Case gErr
    
        Case 193115, 193117, 193118
        
        Case 193116
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 193119
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDDIAFAT", gErr, objSldDiaFat.iFilialEmpresa, objSldDiaFat.sProduto, objSldDiaFat.dtData)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193120)

    End Select
    
    Exit Function

End Function

Private Function SldMesFat_Exclui_EstPSRV(ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Grava a estatística do Mês do produto

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim objSldMesFat As New ClassSldMesFat
Dim iMes As Integer
Dim dFator As Double
Dim dQuantUMVenda As Double
Dim sMes As String
Dim alComando(1 To 2) As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_SldMesFat_Exclui_EstPSRV

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 191131
    Next

    'Para cada item do Pedido de Venda
    For Each objItemPV In objPedVenda.colItensPedido
        
        lErro = SldMesFat_Exclui_EstPSRV1(alComando, objPedVenda, objItemPV)
        If lErro <> SUCESSO Then gError 191132
        
    Next
    
    'Para cada item do Pedido de Venda
    For Each objItemPV In objPedVenda.colItensPedidoSRV
        
        lErro = SldMesFat_Exclui_EstPSRV1(alComando, objPedVenda, objItemPV)
        If lErro <> SUCESSO Then gError 191133
        
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    SldMesFat_Exclui_EstPSRV = SUCESSO

    Exit Function

Erro_SldMesFat_Exclui_EstPSRV:

    SldMesFat_Exclui_EstPSRV = gErr

    Select Case gErr
        
        Case 191131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 191132, 191133
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191134)

    End Select

    Exit Function
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SldMesFat_Exclui_EstPSRV1(alComando() As Long, ByVal objPedVenda As ClassPedidoDeVenda, ByVal objItemPV As ClassItemPedido) As Long
'Grava a estatística do Mês do produto

Dim lErro As Long
Dim objSldMesFat As New ClassSldMesFat
Dim iMes As Integer
Dim dFator As Double
Dim dQuantUMVenda As Double
Dim sMes As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_SldMesFat_Exclui_EstPSRV1

    objSldMesFat.iFilialEmpresa = giFilialEmpresa
    objSldMesFat.iAno = Year(objPedVenda.dtDataEmissao)
    objSldMesFat.sProduto = objItemPV.sProduto
    
    'Lê todos os atributos na tabela SldMesFat e faz lock exclusive
    lErro = CF("SldMesFat_Le_Lock", alComando(1), objSldMesFat)
    If lErro <> SUCESSO And lErro <> 52733 Then gError 193125
            
    If lErro = SUCESSO Then
        
        iMes = Month(objPedVenda.dtDataEmissao)
        
        objProduto.sCodigo = objItemPV.sProduto
        
        'Lê o produto para pegar a UMVenda
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 193126
        
        If lErro = 28030 Then gError 193127
                
        'Calcula o fator para a UM_Venda
        lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError 193128
    
        'converte a quantidade
        dQuantUMVenda = ((objItemPV.dQuantidade - objItemPV.dQuantCancelada) * dFator)
            
        objSldMesFat.dQuantPedida(iMes) = objSldMesFat.dQuantPedida(iMes) - dQuantUMVenda
        
        sMes = CStr(iMes)

        'Atualiza a tabela de SldMesFat para inserção
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE SldMesFat SET QuantPedida" + sMes + "=?", alComando(1), objSldMesFat.dQuantPedida(iMes))
        If lErro <> AD_SQL_SUCESSO Then gError 193129
    
    End If

    SldMesFat_Exclui_EstPSRV1 = SUCESSO

    Exit Function

Erro_SldMesFat_Exclui_EstPSRV1:

    SldMesFat_Exclui_EstPSRV1 = gErr

    Select Case gErr
        
        Case 193125, 193126, 193128
        
        Case 193127
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 193129
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESFAT", gErr, objSldMesFat.iAno, objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193130)

    End Select

    Exit Function
    
End Function

Public Function PSRV_PedidoRepr_Valida(ByVal lPedido As Long, ByVal iFilialEmpresa As Integer, ByVal sPedRepr As String) As Long

Dim lErro As Long

On Error GoTo Erro_PSRV_PedidoRepr_Valida
    
    PSRV_PedidoRepr_Valida = SUCESSO
    
    Exit Function
    
Erro_PSRV_PedidoRepr_Valida:

    PSRV_PedidoRepr_Valida = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193462)

    End Select

    Exit Function
    
End Function
'
'Function ReservasPSRV_Altera(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long) As Long
''Atualiza as Reservas do Pedido de Venda passado
'
'Dim lErro As Long
'
'    'Exclui as reservas do Pedido de Venda
'    lErro = CF("ReservasPSRV_Exclui", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 193484
'
'    'Grava as Reservas do Pedido de Venda
'    lErro = CF("ReservasPSRV_Grava", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 193485
'
'    ReservasPSRV_Altera = SUCESSO
'
'    Exit Function
'
'Erro_ReservasPSRV_Altera:
'
'    ReservasPSRV_Altera = gErr
'
'    Select Case gErr
'
'        Case 193484, 193485
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152549)
'
'    End Select
'
'    Exit Function
'
'End Function

Function MO_Automatico(lCodigo As Long) As Long
'retorna o número da proximo mao de obra disponivel

    MO_Automatico = CF("Config_ObterAutomatico", "SRVConfig", "NUM_PROX_MAODEOBRA", "MaoDeObra", "Codigo", lCodigo)

End Function

Function MO_Exclui(ByVal objMO As ClassMO) As Long
'Exclui um tipo de garantia

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_MO_Exclui
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 193819

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 193820
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM MaoDeObra WHERE Codigo = ?", 0, lCodigo, objMO.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 193821
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193822
        
    If lErro = AD_SQL_SEM_DADOS Then gError 193823

    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM MaoDeObra", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 193824

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 193825

    MO_Exclui = SUCESSO
    
    Exit Function
    
Erro_MO_Exclui:
    
    MO_Exclui = gErr
    
    Select Case gErr
    
        Case 193819
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 193820
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 193821, 193822
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAODEOBRA", gErr, objMO.lCodigo)
        
        Case 193823
            Call Rotina_Erro(vbOKOnly, "ERRO_MAODEOBRA_NAO_CADASTRADA", gErr, objMO.lCodigo)
        
        Case 193824
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MAODEOBRA", gErr, objMO.lCodigo)

        Case 193825
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 193828
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MAODEOBRATIPO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193826)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function MO_Grava(ByVal objMO As ClassMO) As Long
'Grava a Mao de obra passada como parâmetro

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(2) As Long
Dim lCodigo As Long
Dim tMO As typeMO
Dim iAtivo As Integer
Dim lTransacao As Long

On Error GoTo Erro_MO_Grava

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 193827

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 193828
    Next
    
    tMO.sNome = String(STRING_MO_NOME, 0)
    tMO.sNomeReduzido = String(STRING_MO_NOMERED, 0)
    tMO.sObservacao = String(STRING_MO_OBS, 0)

    'Procura por uma outra mao de obra com o mesmo Nome Reduzido
    lErro = Comando_Executar(alComando(0), "SELECT Codigo FROM MaoDeObra WHERE Codigo <> ? AND NomeReduzido = ? ", lCodigo, objMO.lCodigo, objMO.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 193829
    
    'Tenta ler o código desta MO
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193830

    'Se existir outra MO com o mesmo Nome Reduzido, erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 193831
        
    'Seleciona a MO no BD por codigo
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo, Nome, NomeReduzido, Observacao, Ativo  FROM MaoDeObra WHERE Codigo = ?", 0, _
        tMO.lCodigo, tMO.sNome, tMO.sNomeReduzido, tMO.sObservacao, tMO.iAtivo, objMO.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 193832
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193833
    
    'MaoDeObra não está cadastrado - GRAVACAO
    If lErro = AD_SQL_SEM_DADOS Then
        
        lErro = Comando_Executar(alComando(2), "INSERT INTO MaoDeObra (Codigo,Nome,NomeReduzido,Observacao, Ativo, Tipo) VALUES (?,?,?,?,?,?)", _
            objMO.lCodigo, objMO.sNome, objMO.sNomeReduzido, objMO.sObservacao, objMO.iAtivo, objMO.iTipo)
        If lErro <> AD_SQL_SUCESSO Then gError 193834
    
    Else 'Mao de Obra já existe - MODIFICACAO
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE MaoDeObra SET Nome = ?, NomeReduzido = ?, Observacao = ?, Ativo = ?, Tipo = ?", alComando(1), _
            objMO.sNome, objMO.sNomeReduzido, objMO.sObservacao, objMO.iAtivo, objMO.iTipo)
        If lErro <> AD_SQL_SUCESSO Then gError 193835
    
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 193836

    MO_Grava = SUCESSO

    Exit Function

Erro_MO_Grava:

    MO_Grava = gErr

    Select Case gErr

        Case 193827
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 193828
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 193829, 193830
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAODEOBRA", gErr)

        Case 193831
            Call Rotina_Erro(vbOKOnly, "ERRO_MO_NOME_RED_DUPLICADO", gErr)

        Case 193836
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148981)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function


Function OS_Automatico(sCodigoOS As String, ByVal iFilialEmpresa As Integer) As Long
'retorna o número da proxima OS

Dim lTransacao As Long
Dim lErro As Long
Dim lNumProx As Long

On Error GoTo Erro_OS_Automatico

    'Gera número automático.
    lErro = CF("Config_ObterAutomatico", "SRVConfig", "NUM_PROX_OS", "OS", "CodigoNumerico", lNumProx)
    If lErro <> SUCESSO Then gError 195107
    
    sCodigoOS = CStr(lNumProx)
    
    OS_Automatico = SUCESSO

    Exit Function

Erro_OS_Automatico:

    OS_Automatico = gErr

    Select Case gErr

        Case 195107

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195108)

    End Select
    
    Exit Function

End Function

Function OS_Grava(ByVal objOS As ClassOS) As Long
'cria ou atualiza uma ordem de servico

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_OS_Grava

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 195213

    lErro = CF("OS_Grava_EmTrans", objOS)
    If lErro <> AD_SQL_SUCESSO Then gError 195214

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 195215

    OS_Grava = SUCESSO

    Exit Function

Erro_OS_Grava:

    OS_Grava = gErr

    Select Case gErr

        Case 195213
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 195214

        Case 195215
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195216)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function

Function OS_Grava_EmTrans(ByVal objOS As ClassOS) As Long
'cria ou atualiza uma ordem de servico

Dim lErro As Long
Dim iIndice As Integer
Dim sDescricao As String
Dim dtDataEmissao As Date
Dim iNumItens As Integer
Dim iNumItensBaixados As Integer
Dim objItemOS As New ClassItemOS
Dim alComando(16) As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lCodigoNumerico As Long
Dim iStatus As Integer
Dim lNumIntItemOS As Long

On Error GoTo Erro_OS_Grava_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comandos
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195274
    Next
    
    If objOS.iNumItens = objOS.iNumItensBaixados Then
        objOS.iStatus = STATUSOS_BAIXADA
    Else
        objOS.iStatus = STATUSOS_ABERTA
    End If

    'Leitura da tabela de Ordens de Servico
    lErro = Comando_ExecutarPos(alComando(1), "SELECT DataEmissao, NumItens, NumItensBaixados, Status FROM OS WHERE FilialEmpresa = ? AND Codigo = ? ", 0, dtDataEmissao, iNumItens, iNumItensBaixados, iStatus, objOS.iFilialEmpresa, objOS.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195275

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195276

    'Existe a OS no BD
    If lErro = AD_SQL_SUCESSO Then

        'Lock na OS
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 195277

        'atualiza os itens pré-existentes da OS
        lErro = OS_Grava4(objOS)
        If lErro <> SUCESSO Then gError 195278

        lErro = Comando_ExecutarPos(alComando(2), "UPDATE OS SET Tipo = ?, DataEmissao = ?, NumItens = ?, NumItensBaixados = ?, Status = ?, CodSolSRV = ?, Produto = ?, Lote = ?, FilialOP = ?, CodPedSRV = ?, FilialPedSRV = ?", alComando(1), objOS.lTipo, objOS.dtDataEmissao, objOS.iNumItens, objOS.iNumItensBaixados, objOS.iStatus, objOS.lCodSolSRV, objOS.sProduto, objOS.sLote, objOS.iFilialOP, objOS.lCodPedSRV, objOS.iFilialPedSRV)
        If lErro <> AD_SQL_SUCESSO Then gError 195279

    Else 'Nao existe a OS no BD

        
        If IsNumeric(objOS.sCodigo) Then
            lCodigoNumerico = StrParaLong(objOS.sCodigo)
        End If
        
        'Insere nova OP
        lErro = Comando_Executar(alComando(2), "INSERT INTO OS (Tipo, Codigo, FilialEmpresa, DataEmissao, NumItens, NumItensBaixados, CodigoNumerico, Status, CodSolSRV, Produto, Lote, FilialOP, CodPedSRV, FilialPedSRV) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", objOS.lTipo, objOS.sCodigo, objOS.iFilialEmpresa, objOS.dtDataEmissao, objOS.iNumItens, objOS.iNumItensBaixados, lCodigoNumerico, objOS.iStatus, objOS.lCodSolSRV, objOS.sProduto, objOS.sLote, objOS.iFilialOP, objOS.lCodPedSRV, objOS.iFilialPedSRV)
        If lErro <> AD_SQL_SUCESSO Then gError 195280

    End If
        

    'obtem o proximo codigo para NumIntItemOS
    lErro = OS_Grava3(lNumIntItemOS)
    If lErro <> AD_SQL_SUCESSO Then gError 195281

    'inclui os novos itens da OS
    For Each objItemOS In objOS.colItens
        
        If objItemOS.lNumIntDoc = 0 Then

            objItemOS.lNumIntDoc = lNumIntItemOS
            
            lErro = OS_Grava1(objItemOS)
            If lErro <> SUCESSO Then gError 195282

            lNumIntItemOS = lNumIntItemOS + 1

        End If
        

    Next

    'atualiza o proximo codigo para lNumIntItemOS
    lErro = OS_Grava5(lNumIntItemOS)
    If lErro <> AD_SQL_SUCESSO Then gError 195283

    lErro = OS_Grava_Operacoes_EmTrans(objOS)
    If lErro <> SUCESSO Then gError 195284
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OS_Grava_EmTrans = SUCESSO

    Exit Function

Erro_OS_Grava_EmTrans:

    OS_Grava_EmTrans = gErr

    Select Case gErr

        Case 195274
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195275, 195276
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OS", gErr)

        Case 195277
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_OS", gErr)

        Case 195278, 195281 To 195284

        Case 195279
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_OS", gErr, objOS.iFilialEmpresa, objOS.sCodigo)

        Case 195280
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OS", gErr, objOS.iFilialEmpresa, objOS.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195285)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function OS_Grava4(ByVal objOS As ClassOS) As Long
'atualiza os itens pré-existentes da OS

Dim alComando(1 To 5) As Long, iIndice As Integer
Dim tItemOS As typeItemOS, objItemOS As ClassItemOS
Dim iAchou As Integer, sCcl As String, lErro As Long
Dim sParam1 As String, sParam2 As String

On Error GoTo Erro_OS_Grava4

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195220
    Next

    tItemOS.sCcl = String(STRING_CCL, 0)
    tItemOS.sServico = String(STRING_PRODUTO, 0)
    tItemOS.sVersao = String(STRING_KIT_VERSAO, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Servico, NumIntDoc, Item, Ccl, Quantidade, DataInicio, DataFim, Prioridade, StatusItem, Versao FROM ItemOS WHERE FilialEmpresa = ? AND Codigo = ? ORDER BY Item", 0, _
    tItemOS.sServico, tItemOS.lNumIntDoc, tItemOS.iItem, tItemOS.sCcl, tItemOS.dQuantidade, tItemOS.dtDataInicio, tItemOS.dtDataFim, tItemOS.iPrioridade, tItemOS.iStatusItem, tItemOS.sVersao, objOS.iFilialEmpresa, objOS.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195221

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195222

    Do While lErro <> AD_SQL_SEM_DADOS

        iAchou = 0

        'localiza item na coleção
        For Each objItemOS In objOS.colItens
            If objItemOS.lNumIntDoc = tItemOS.lNumIntDoc Then
                iAchou = 1
                Exit For
            End If
        Next

        'se nao achou o item na tela ==> exclui o item
        If iAchou = 0 Then
        
            lErro = ItemOS_Exclui(tItemOS.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 195223
        
        Else

            If tItemOS.sCcl <> objItemOS.sCcl Then
    
                'bloquear ccl
                sCcl = String(STRING_CCL, 0)
                lErro = Comando_ExecutarLockado(alComando(3), "SELECT Ccl FROM Ccl WHERE Ccl = ?", sCcl, objItemOS.sCcl)
                If lErro <> AD_SQL_SUCESSO Then gError 195224
    
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195225
    
                If lErro = AD_SQL_SEM_DADOS Then gError 195226
    
                lErro = Comando_LockShared(alComando(3))
                If lErro <> AD_SQL_SUCESSO Then gError 195227
    
            End If
    
            With objItemOS
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItemOS SET Item = ?,  Ccl = ?, Quantidade = ?, DataInicio = ?, DataFim = ?, Prioridade = ?, StatusItem = ?, Observacao = ? ", alComando(1), _
                   .iItem, .sCcl, .dQuantidade, .dtDataInicio, .dtDataFim, .iPrioridade, .iStatusItem, .sObservacao)
            End With
            If lErro <> AD_SQL_SUCESSO Then gError 195228
    
        End If

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195229

    Loop

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OS_Grava4 = SUCESSO

    Exit Function

Erro_OS_Grava4:

    OS_Grava4 = gErr

    Select Case gErr

        Case 195220
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195221, 195222, 195229
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOS", gErr)

        Case 195223

        Case 195224, 195225
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", gErr)

        Case 195226
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, objItemOS.sCcl)

        Case 195227
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", gErr)

        Case 195228
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITEMOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195230)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItemOS_Exclui(lNumIntItemOS As Long) As Long

Dim lErro As Long
Dim lNumIntDocOper As Long
Dim alComando(1 To 6) As Long
Dim iIndice As Integer

On Error GoTo Erro_ItemOS_Exclui

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195231
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM ItemOSOperacoes WHERE NumIntDocItemOS = ? ", 0, lNumIntDocOper, lNumIntItemOS)
    If lErro <> AD_SQL_SUCESSO Then gError 195232
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195233

    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_Executar(alComando(3), "DELETE FROM ItemOSOperacoesMaquinas WHERE NumIntDocOper = ?", lNumIntDocOper)
        If lErro <> AD_SQL_SUCESSO Then gError 195234

        lErro = Comando_Executar(alComando(4), "DELETE FROM ItemOSOperacoesMO WHERE NumIntDocOper = ?", lNumIntDocOper)
        If lErro <> AD_SQL_SUCESSO Then gError 195235

        lErro = Comando_Executar(alComando(5), "DELETE FROM ItemOSOperacoesPecas WHERE NumIntDocOper = ?", lNumIntDocOper)
        If lErro <> AD_SQL_SUCESSO Then gError 195236

        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ItemOSOperacoes", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 195237

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195238

    Loop

    lErro = Comando_Executar(alComando(6), "DELETE FROM ItemOS WHERE NumIntDoc = ?", lNumIntItemOS)
    If lErro <> AD_SQL_SUCESSO Then gError 195239

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    ItemOS_Exclui = SUCESSO

    Exit Function

Erro_ItemOS_Exclui:

    ItemOS_Exclui = gErr

    Select Case gErr

        Case 195231
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195232, 195233, 195238
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOSOPERACOES", gErr)

        Case 195234
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOSOPERACOESMAQUINAS", gErr)

        Case 195235
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOSOPERACOESMO", gErr)

        Case 195236
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOSOPERACOESPECAS", gErr)

        Case 195237
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOSOPERACOES", gErr)

        Case 195239
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195240)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function OS_Grava3(lNumIntItemOS As Long) As Long
'obtem o proximo codigo para lNumIntDocItem e lEmpenho

Dim lErro As Long, sConteudo As String
Dim iIndice As Integer
Dim alComando(1 To 1) As Long

On Error GoTo Erro_OS_Grava3

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195241
    Next

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Conteudo FROM SRVConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, "NUM_PROX_ITEMOS", EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then gError 195242

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 195243

    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 195244

    lNumIntItemOS = CLng(sConteudo)

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    OS_Grava3 = SUCESSO

    Exit Function

Erro_OS_Grava3:

    OS_Grava3 = gErr

    Select Case gErr

        Case 195241
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195242, 195243
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SRVCONFIG", gErr, "NUM_PROX_ITEMOS", EMPRESA_TODA)

        Case 195244
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_SRVCONFIG", gErr, "NUM_PROX_ITEMOS", EMPRESA_TODA)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195245)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function OS_Grava1(ByVal objItemOS As ClassItemOS) As Long
'Inclui um item numa OP
'Tem que ser chamada dentro de transação

Dim lErro As Long
Dim iPCP As Integer
Dim colProdutos As New Collection
Dim sCcl As String
Dim alComando(1 To 5) As Long
Dim iCompras As Integer
Dim iIndice As Integer

On Error GoTo Erro_OS_Grava1

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195255
    Next

    'Leitura da tabela Produtos
    lErro = Comando_ExecutarLockado(alComando(1), "SELECT PCP, Compras FROM Produtos WHERE Codigo = ? ", iPCP, iCompras, objItemOS.sServico)
    If lErro <> AD_SQL_SUCESSO Then gError 195246

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195247

    If lErro = AD_SQL_SEM_DADOS Then gError 195248

    lErro = Comando_LockShared(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 195249

    'garantir a existencia da Ccl
    If objItemOS.sCcl <> "" Then

        sCcl = String(STRING_CCL, 0)

        lErro = Comando_ExecutarLockado(alComando(2), "SELECT Ccl FROM Ccl WHERE Ccl = ? ", sCcl, objItemOS.sCcl)
        If lErro <> AD_SQL_SUCESSO Then gError 195250

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195251

        If lErro = AD_SQL_SEM_DADOS Then gError 195252

        lErro = Comando_LockShared(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 195253

    End If
    
    'Insere item da OS
    lErro = Comando_Executar(alComando(3), "INSERT INTO ItemOS (NumIntDoc, FilialEmpresa, Codigo, Item, Servico, " & _
    "Ccl, SiglaUM, Quantidade, DataInicio, DataFim, Prioridade, StatusItem, Versao, Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
    objItemOS.lNumIntDoc, objItemOS.iFilialEmpresa, objItemOS.sCodigo, objItemOS.iItem, objItemOS.sServico, _
    objItemOS.sCcl, objItemOS.sSiglaUM, objItemOS.dQuantidade, objItemOS.dtDataInicio, objItemOS.dtDataFim, objItemOS.iPrioridade, _
    objItemOS.iStatusItem, objItemOS.sVersao, objItemOS.sObservacao)
    If lErro <> SUCESSO Then gError 195254

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    OS_Grava1 = SUCESSO

    Exit Function

Erro_OS_Grava1:

    OS_Grava1 = gErr

    Select Case gErr

        Case 195246, 195247
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case 195248
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objItemOS.sServico)

        Case 195249
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTO", gErr)

        Case 195250, 195251
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", gErr)

        Case 195252
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, objItemOS.sCcl)

        Case 195253
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", gErr)

        Case 195254
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITEMOS", gErr)

        Case 195255
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195256)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function OS_Grava5(ByVal lNumIntItemOS As Long) As Long
'obtem o proximo codigo para lNumIntDocItem e lEmpenho

Dim lErro As Long, sConteudo As String
Dim iIndice As Integer
Dim alComando(1 To 2) As Long

On Error GoTo Erro_OS_Grava5

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195257
    Next

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Conteudo FROM SRVConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, "NUM_PROX_ITEMOS", EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then gError 195258

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 195259

    lErro = Comando_ExecutarPos(alComando(2), "UPDATE SRVConfig SET Conteudo = ?", alComando(1), lNumIntItemOS)
    If lErro <> AD_SQL_SUCESSO Then gError 195260

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    OS_Grava5 = SUCESSO

    Exit Function

Erro_OS_Grava5:

    OS_Grava5 = gErr

    Select Case gErr

        Case 195257
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195258, 195259
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SRVCONFIG", gErr, "NUM_PROX_ITEMOS", EMPRESA_TODA)

        Case 195260
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SRVCONFIG", gErr, "NUM_PROX_ITEMOS", EMPRESA_TODA)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195261)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function OS_Grava_Operacoes_EmTrans(ByVal objOS As ClassOS) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim objItemOS As New ClassItemOS
Dim objItemOSOperacoes As New ClassItemOSOperacoes
Dim lNumIntDoc As Long
Dim bApagarOperacao As Boolean

On Error GoTo Erro_OS_Grava_Operacoes_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comandos
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195347
    Next
    
    'Inclui as Operações dos itens da OS
    For Each objItemOS In objOS.colItens
    
        For Each objItemOSOperacoes In objItemOS.colItemOSOperacoes
        
            'Le a tabela ItemOSOperacoes
            lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ItemOSOperacoes WHERE NumIntDocItemOS= ? AND Seq= ? ", _
                        0, lNumIntDoc, objItemOS.lNumIntDoc, objItemOSOperacoes.iSeq)
            If lErro <> AD_SQL_SUCESSO Then gError 195348
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195349
        
            'Se existir => UPDATE, senão => INSERT
            If lErro = AD_SQL_SUCESSO Then
            
                objItemOSOperacoes.lNumIntDoc = lNumIntDoc
                objItemOSOperacoes.lNumIntDocItemOS = objItemOS.lNumIntDoc
            
                'Trava registro contra alterações/Leituras
                lErro = Comando_LockExclusive(alComando(0))
                If lErro <> SUCESSO Then gError 195350
                
                With objItemOSOperacoes
                
                    lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItemOSOperacoes SET Seq= ? , NumIntDocCompet= ? , NumIntDocCT= ? , Observacao= ? , " & _
                            "SeqPai= ? , SeqArvore= ? , Nivel= ? , SeqRoteiro = ?, SeqRoteiroPai = ? , NivelRoteiro = ? , Servico = ? , Versao = ? , IgnoraTaxaProducao = ? , " & _
                            "ConsideraCarga = ? , Origem = ? , NumIntDocOperOrigem = ?, NumMaxMaqPorOper = ?, NumRepeticoes = ? ", alComando(0), _
                            .iSeq, .lNumIntDocCompet, .lNumIntDocCT, .sObservacao, _
                            .iSeqPai, .iSeqArvore, .iNivel, .iSeqRoteiro, .iSeqRoteiroPai, .iNivelRoteiro, _
                            .sServico, .sVersao, .iIgnoraTaxaProducao, .iConsideraCarga, _
                            .iOrigem, .lNumIntDocOPerOrigem, .iNumMaxMaqPorOper, .iNumRepeticoes)
                    If lErro <> AD_SQL_SUCESSO Then gError 135351
                
                End With
                
                lErro = ItemOSOperacoesPecas_Exclui(objItemOSOperacoes)
                If lErro <> SUCESSO Then gError 195352
                            
                lErro = ItemOSOperacoesMaquinas_Exclui(objItemOSOperacoes)
                If lErro <> SUCESSO Then gError 195353
                
                lErro = ItemOSOperacoesMO_Exclui(objItemOSOperacoes)
                If lErro <> SUCESSO Then gError 195354
                
            Else
            
                'Obter NumIntDoc
                lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_ITEMOSOPERACOES", lNumIntDoc)
                If lErro <> SUCESSO Then gError 195355
        
                objItemOSOperacoes.lNumIntDoc = lNumIntDoc
                objItemOSOperacoes.lNumIntDocItemOS = objItemOS.lNumIntDoc
                
                With objItemOSOperacoes
                
                    lErro = Comando_Executar(alComando(1), "INSERT INTO ItemOSOperacoes( NumIntDoc, NumIntDocItemOS, Seq, NumIntDocCompet, NumIntDocCT, Observacao, " & _
                            "SeqPai, SeqArvore, Nivel, SeqRoteiro, SeqRoteiroPai, NivelRoteiro, Servico, Versao, IgnoraTaxaProducao, ConsideraCarga, Origem, NumIntDocOperOrigem, NumMaxMaqPorOper, NumRepeticoes) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                            .lNumIntDoc, .lNumIntDocItemOS, .iSeq, .lNumIntDocCompet, .lNumIntDocCT, .sObservacao, _
                            .iSeqPai, .iSeqArvore, .iNivel, .iSeqRoteiro, .iSeqRoteiroPai, .iNivelRoteiro, .sServico, .sVersao, .iIgnoraTaxaProducao, .iConsideraCarga, .iOrigem, .lNumIntDocOPerOrigem, .iNumMaxMaqPorOper, .iNumRepeticoes)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195356
                
                End With
                
            End If
            
            lErro = ItemOSOperacoesPecas_Grava(objItemOSOperacoes)
            If lErro <> SUCESSO Then gError 195357
                        
            lErro = ItemOSOperacoesMaquinas_Grava(objItemOSOperacoes)
            If lErro <> SUCESSO Then gError 195358
            
            lErro = ItemOSOperacoesMO_Grava(objItemOSOperacoes)
            If lErro <> SUCESSO Then gError 195359
            
            
        Next
        
        'Le a tabela OrdemProducaoOperacoes novamente para verificar se haverá deleções
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM ItemOSOperacoes WHERE NumIntDocItemOS= ? ", _
                    0, lNumIntDoc, objItemOS.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 195360
        
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195361
    
        'Faça enquanto existir
        Do While lErro = AD_SQL_SUCESSO
        
            bApagarOperacao = True
            
            'Para cada OrdemProducaoOperacoes na coleção
            For Each objItemOSOperacoes In objItemOS.colItemOSOperacoes
            
                If objItemOSOperacoes.lNumIntDoc = lNumIntDoc Then
                
                     bApagarOperacao = False
                     Exit For
                
                End If
            
            Next
            
            If bApagarOperacao Then
                
                lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM ItemOSOperacoes", alComando(2))
                If lErro <> AD_SQL_SUCESSO Then gError 195362
                
                lErro = ItemOSOperacoesPecas_Exclui(objItemOSOperacoes)
                If lErro <> SUCESSO Then gError 195363
                            
                lErro = ItemOSOperacoesMaquinas_Exclui(objItemOSOperacoes)
                If lErro <> SUCESSO Then gError 195364
                
                lErro = ItemOSOperacoesMO_Exclui(objItemOSOperacoes)
                If lErro <> SUCESSO Then gError 195365
                    
            End If
    
            'vai para o próximo
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195366
    
        Loop
        
    Next
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    OS_Grava_Operacoes_EmTrans = SUCESSO
    
    Exit Function

Erro_OS_Grava_Operacoes_EmTrans:

    OS_Grava_Operacoes_EmTrans = gErr
    
    Select Case gErr
        
        Case 195347
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195348, 195349, 195360, 195361, 195366
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOSOPERACOES", gErr)

        Case 195350
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMOSOPERACOES", gErr)

        Case 195351
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITEMOSOPERACOES", gErr)
            
        Case 195352 To 195355, 195357 To 195359, 195363 To 195365

        Case 195356
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMOSOPERACOES", gErr)
            
        Case 195362
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOSOPERACOES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195367)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Function ItemOSOperacoesPecas_Exclui(ByVal objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_ItemOSOperacoesPecas_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195262
    Next
    
    lErro = Comando_Executar(alComando(1), "DELETE FROM ItemOSOperacoesPecas WHERE NumIntDocOper = ?", objItemOSOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 195263

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOSOperacoesPecas_Exclui = SUCESSO

    Exit Function

Erro_ItemOSOperacoesPecas_Exclui:

    ItemOSOperacoesPecas_Exclui = gErr

    Select Case gErr

        Case 195262
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195263
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOSOPERACOESPECAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195264)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItemOSOperacoesMaquinas_Exclui(ByVal objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_ItemOSOperacoesMaquinas_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195265
    Next
    
    lErro = Comando_Executar(alComando(1), "DELETE FROM ItemOSOperacoesMaquinas WHERE NumIntDocOper = ?", objItemOSOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 195266

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOSOperacoesMaquinas_Exclui = SUCESSO

    Exit Function

Erro_ItemOSOperacoesMaquinas_Exclui:

    ItemOSOperacoesMaquinas_Exclui = gErr

    Select Case gErr

        Case 195265
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195266
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOSOPERACOESMAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195267)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItemOSOperacoesMO_Exclui(ByVal objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_ItemOSOperacoesMO_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195268
    Next
    
    lErro = Comando_Executar(alComando(1), "DELETE FROM ItemOSOperacoesMO WHERE NumIntDocOper = ?", objItemOSOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 195269

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOSOperacoesMO_Exclui = SUCESSO

    Exit Function

Erro_ItemOSOperacoesMO_Exclui:

    ItemOSOperacoesMO_Exclui = gErr

    Select Case gErr

        Case 195268
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195269
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOSOPERACOESMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195270)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ItemOSOperacoesPecas_Grava(ByVal objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim objItemOSOperacoesPecas As New ClassItemOSOperacoesPecas

On Error GoTo Erro_ItemOSOperacoesPecas_Grava

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 195271
        
    'Para cada objOrdemProducaoInsumos na coleção
    For Each objItemOSOperacoesPecas In objItemOSOperacoes.colPecas
        
        objItemOSOperacoesPecas.lNumIntDocOper = objItemOSOperacoes.lNumIntDoc

        With objItemOSOperacoesPecas
            
            lErro = Comando_Executar(lComando, "INSERT INTO ItemOSOperacoesPecas (NumIntDocOper, Produto, Quantidade, UMProduto, VersaoKitComp, Composicao, PercentualPerda, CustoStandard) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", _
                                                .lNumIntDocOper, .sProduto, .dQuantidade, .sUMProduto, .sVersaoKitComp, .iComposicao, .dPercentualPerda, .dCustoStandard)
            If lErro <> AD_SQL_SUCESSO Then gError 195272
            
        End With
                           
    Next
        
    'fecha o comando
    Call Comando_Fechar(lComando)
        
    ItemOSOperacoesPecas_Grava = SUCESSO

    Exit Function

Erro_ItemOSOperacoesPecas_Grava:

    ItemOSOperacoesPecas_Grava = gErr

    Select Case gErr

        Case 195271
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195272
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITEMOSOPERACOESPECAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195273)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItemOSOperacoesMaquinas_Grava(ByVal objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim objItemOSOperacoesMaquinas As New ClassItemOSOperacoesMaquinas

On Error GoTo Erro_ItemOSOperacoesMaquinas_Grava

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 195344
    
        
    'Para cada objOrdemProducaoInsumos na coleção
    For Each objItemOSOperacoesMaquinas In objItemOSOperacoes.colMaquinas
        
        objItemOSOperacoesMaquinas.lNumIntDocOper = objItemOSOperacoes.lNumIntDoc

        With objItemOSOperacoesMaquinas
            
            lErro = Comando_Executar(lComando, "INSERT INTO ItemOSOperacoesMaquinas (NumIntDocOper, NumIntDocMaq, Horas, Quantidade) VALUES (?, ?, ?, ?)", _
                                               .lNumIntDocOper, .lNumIntDocMaq, .dHoras, .iQuantidade)
            If lErro <> AD_SQL_SUCESSO Then gError 195345
            
        End With
                           
    Next
        
    'fecha o comando
    Call Comando_Fechar(lComando)
        
    ItemOSOperacoesMaquinas_Grava = SUCESSO

    Exit Function

Erro_ItemOSOperacoesMaquinas_Grava:

    ItemOSOperacoesMaquinas_Grava = gErr

    Select Case gErr

        Case 195344
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195345
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITEMOSOPERACOESMAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195346)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function OS_Exclui(ByVal objOS As ClassOS) As Long
'Função que exclui uma ordem de servico a partir do código passado em objOS

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_OS_Exclui

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 195308

    lErro = OS_Exclui_EmTrans(objOS)
    If lErro <> SUCESSO And lErro <> 195301 Then gError 195309

    If lErro <> SUCESSO Then gError 195312

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 195310

    OS_Exclui = SUCESSO

    Exit Function

Erro_OS_Exclui:

    OS_Exclui = gErr

    Select Case gErr

        Case 195308
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 195309, 195312
        
        Case 195310
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 195311)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function

Function OS_Exclui_EmTrans(ByVal objOS As ClassOS) As Long
'Função que exclui uma ordem de serviço a partir do código passado em objOS

Dim lErro As Long
Dim iItem As Integer
Dim iIndice As Integer
Dim iStatusItem As Integer, sServico As String, dQtde As Double
Dim iNumItens As Integer
Dim lNumIntDoc As Long
Dim alComando(1 To 10) As Long
Dim objItemOS As New ClassItemOS
Dim sCodigo As String
Dim lCodigo As Long
Dim lCodigoMovEst As Long

On Error GoTo Erro_OS_Exclui_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195298
    Next

    'tenta ler se existe a OS
    lErro = Comando_ExecutarPos(alComando(3), "SELECT NumItens FROM OS WHERE FilialEmpresa = ? AND Codigo = ? ", 0, iNumItens, objOS.iFilialEmpresa, objOS.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195299

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195300

    If lErro = AD_SQL_SEM_DADOS Then gError 195301

    'seleciona os itens da OS
    sServico = String(STRING_PRODUTO, 0)
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, Item, StatusItem, Servico, Quantidade FROM ItemOS WHERE FilialEmpresa = ? AND Codigo = ? ", 0, lNumIntDoc, iItem, iStatusItem, sServico, dQtde, objOS.iFilialEmpresa, objOS.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195302

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195303

    'percorre os itens
    Do While lErro <> AD_SQL_SEM_DADOS

        objItemOS.lNumIntDoc = lNumIntDoc
        objItemOS.iFilialEmpresa = objOS.iFilialEmpresa
        objItemOS.sServico = sServico
        objItemOS.dQuantidade = dQtde
        objItemOS.iItem = iItem
        
        '***********************************************
        'Testa para ver se tem movimentações de materiais em cima do item da OS\apontamentos
        lErro = Comando_Executar(alComando(8), "SELECT Codigo FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? ", lCodigoMovEst, MOVEST_TIPONUMINTDOCORIGEM_ITEMOS, objItemOS.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 206428
    
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206429
        
        'Se tem, não pode excluir
        If lErro = AD_SQL_SUCESSO Then gError 206430
        
        'Testa para ver se tem movimentações de materiais em cima do item da OS\apontamentos
        lErro = Comando_Executar(alComando(9), "SELECT CodigoAp FROM OSApMO WHERE OS = ? AND ProdutoOS = ? ", lCodigoMovEst, objOS.sCodigo, objItemOS.sServico)
        If lErro <> AD_SQL_SUCESSO Then gError 206431
    
        lErro = Comando_BuscarPrimeiro(alComando(9))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206432
        
        'Se tem, não pode excluir
        If lErro = AD_SQL_SUCESSO Then gError 206433
        
        'Testa para ver se tem movimentações de materiais em cima do item da OS\apontamentos
        lErro = Comando_Executar(alComando(10), "SELECT CodigoAp FROM OSApMaq WHERE OS = ? AND ProdutoOS = ? ", lCodigoMovEst, objOS.sCodigo, objItemOS.sServico)
        If lErro <> AD_SQL_SUCESSO Then gError 206434
    
        lErro = Comando_BuscarPrimeiro(alComando(10))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206435
        
        'Se tem, não pode excluir
        If lErro = AD_SQL_SUCESSO Then gError 206436
        '***********************************************

        sServico = String(STRING_PRODUTO, 0)

        lErro = OS_Exclui_Operacoes_EmTrans(objItemOS)
        If lErro <> SUCESSO Then gError 195304

        lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM ItemOS", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 195305
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195306

    Loop

    lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM OS", alComando(3))
    If lErro <> AD_SQL_SUCESSO Then gError 195307

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OS_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_OS_Exclui_EmTrans:

    OS_Exclui_EmTrans = gErr

    Select Case gErr

        Case 195298
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195299, 195300
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OS", gErr)

        Case 195301, 195304

        Case 195302, 195303, 195306
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOS", gErr)

        Case 195305
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOS", gErr)

        Case 195307
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OS", gErr)
            
        Case 206428, 206429, 206431, 206432, 206434, 206435
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
        
        Case 206430, 206433, 206436
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OS_COM_MOV", gErr, lCodigoMovEst)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195308)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function OS_Exclui_Operacoes_EmTrans(ByVal objItemOS As ClassItemOS) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objItemOSOperacoes As ClassItemOSOperacoes
Dim lNumIntDocOper As Long

On Error GoTo Erro_OS_Exclui_Operacoes_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comandos
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195286
    Next
    
    'Le a tabela ItemOSOperacoes
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ItemOSOperacoes WHERE NumIntDocItemOS= ? ", _
                0, lNumIntDocOper, objItemOS.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 195287

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195288

    'Faz enquanto existir
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItemOSOperacoes = New ClassItemOSOperacoes
        
        objItemOSOperacoes.lNumIntDoc = lNumIntDocOper

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 195289
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItemOSOperacoes", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 195290
        
        lErro = ItemOSOperacoesPecas_Exclui(objItemOSOperacoes)
        If lErro <> SUCESSO Then gError 195291
    
        lErro = ItemOSOperacoesMaquinas_Exclui(objItemOSOperacoes)
        If lErro <> SUCESSO Then gError 195292
        
        lErro = ItemOSOperacoesMO_Exclui(objItemOSOperacoes)
        If lErro <> SUCESSO Then gError 195293
        
        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195294

    Loop

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    OS_Exclui_Operacoes_EmTrans = SUCESSO
    
    Exit Function

Erro_OS_Exclui_Operacoes_EmTrans:

    OS_Exclui_Operacoes_EmTrans = gErr
    
    Select Case gErr
        
        Case 195286
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195287, 195288, 195294
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOSOPERACOES", gErr)

        Case 195289
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMOSOPERACOES", gErr)

        Case 195290
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOSOPERACOES", gErr)
            
        Case 195291 To 195293
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152297)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Public Function ItemOSOperacoesMO_Grava(ByVal objItemOSOperacoes As ClassItemOSOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim objItemOSOperacoesMO As New ClassItemOSOperacoesMO

On Error GoTo Erro_ItemOSOperacoesMO_Grava

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 195341
    
        
    'Para cada objOrdemProducaoInsumos na coleção
    For Each objItemOSOperacoesMO In objItemOSOperacoes.colMO
        
        objItemOSOperacoesMO.lNumIntDocOper = objItemOSOperacoes.lNumIntDoc

        With objItemOSOperacoesMO
            
            lErro = Comando_Executar(lComando, "INSERT INTO ItemOSOperacoesMO (NumIntDocOper, CodigoMO, Horas) VALUES (?, ?, ?)", _
                                                .lNumIntDocOper, .lCodigoMO, .dHoras)
            If lErro <> AD_SQL_SUCESSO Then gError 195342
            
        End With
                           
    Next
        
    'fecha o comando
    Call Comando_Fechar(lComando)
        
    ItemOSOperacoesMO_Grava = SUCESSO

    Exit Function

Erro_ItemOSOperacoesMO_Grava:

    ItemOSOperacoesMO_Grava = gErr

    Select Case gErr

        Case 195341
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195342
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITEMOSOPERACOESMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195343)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ItensDeContratoSrv_Exclui(objItensDeContratoSRV As ClassItensDeContratoSrv) As Long

Dim lErro As Long, lTransacao As Long
Dim tItensContSrv As typeItensDeContratoSrv
Dim objSolicSRV As New ClassSolicSRV
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim objItensSolicSrv As ClassItensSolicSRV, objProduto As New ClassProduto
Dim objItensContSrvNumSerie As ClassItensContSrvNumSerie
Dim bChamadoForaTransacao As Boolean

On Error GoTo Erro_ItensDeContratoSrv_Exclui

    bChamadoForaTransacao = False
    
    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comandos
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195587
    Next
        
    If GL_lTransacao = 0 Then
        bChamadoForaTransacao = True
        
        'Abrir transacao
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 183396
        
    End If
        
    'se encontrou alguma solicitacao envolvendo o contrato e produto em questao
    lErro = CF("ItensSolicSRV_Le_Contrato", objSolicSRV, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.sCodigoContrato, objItensDeContratoSRV.sProduto)
    If lErro <> SUCESSO And lErro <> 195596 Then gError 195601
        
    If lErro = SUCESSO Then
    
        objProduto.sCodigo = objItensDeContratoSRV.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195690
            
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_OP Then
        
            For Each objItensSolicSrv In objSolicSRV.colItens
            
                If objItensSolicSrv.sLote = objItensDeContratoSRV.sLote And objItensSolicSrv.iFilialOP = objItensDeContratoSRV.iFilialOP Then gError 195687
                
            Next
        
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then

            'pesquisa as solicitacoes de servico
            For Each objItensSolicSrv In objSolicSRV.colItens
            
                For Each objItensContSrvNumSerie In objItensDeContratoSRV.colNumSerie
                
                    If objItensContSrvNumSerie.sNumSerie = objItensSolicSrv.sLote Then gError 195688
                    
                Next
                
            Next

        Else
            gError 195689
        End If

    End If

    If objItensDeContratoSRV.lNumIntDoc <> 0 Then
    
        lErro = Comando_Executar(alComando(1), "DELETE FROM ItensDeContratoSRVProd WHERE NumIntItemContratoSRV=?", objItensDeContratoSRV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 195603
    
        lErro = Comando_Executar(alComando(2), "DELETE FROM ItensDeContratoSRVNumSerie WHERE NumIntItemContratoSRV=?", objItensDeContratoSRV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 195604
            
        lErro = Comando_Executar(alComando(3), "DELETE FROM ItensDeContratoSRV WHERE NumIntDoc = ?", objItensDeContratoSRV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 195588
        
    End If
    
    If bChamadoForaTransacao Then
    
        'Fechar transacao
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 183398
   
   End If
   
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItensDeContratoSrv_Exclui = SUCESSO
    
    Exit Function

Erro_ItensDeContratoSrv_Exclui:
    
    ItensDeContratoSrv_Exclui = gErr
        
    Select Case gErr
    
        Case 195587
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 195588
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMCONTRATOSRV", gErr, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.lCodigo)
        
        Case 195601
        
        Case 195687, 195688, 195689
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMCONTRATOSRV_EM_SOLICITACAOSRV", gErr)

        Case 195603
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMCONTRATOSRVPRODUTO", gErr, objItensDeContratoSRV.lNumIntDoc)
        
        Case 195604
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMCONTRATOSRVNUMSERIE", gErr, objItensDeContratoSRV.lNumIntDoc)
        
        Case 183396
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 183398
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195585)
            
    End Select
    
    If bChamadoForaTransacao Then Call Transacao_Rollback
   
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function ItensDeContratoSrv_Grava(ByVal objItensDeContratoSRV As ClassItensDeContratoSrv) As Long
'grava na tabela Garantia

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 1) As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer

On Error GoTo Erro_ItensDeContratoSrv_Grava
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 195655

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195656
    Next

    'Faz o lock nos objetos vinculados a ItensDeContratoSrv
    lErro = ItensDeContratoSrvObjVinc_Lock(objItensDeContratoSRV)
    If lErro <> SUCESSO Then gError 195657

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM ItensDeContratoSrv WHERE FilialEmpresa = ? AND Codigo = ?", 0, lNumIntDoc, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195658
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195659
        
    If lErro = AD_SQL_SEM_DADOS Then

        'Insere a garantia nova
        lErro = ItensDeContratoSrv_Insere_BD(objItensDeContratoSRV)
        If lErro <> SUCESSO Then gError 195660

    Else
    
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 195661
    
        objItensDeContratoSRV.lNumIntDoc = lNumIntDoc
    
        'Altera a ItensDeContratoSrv
        lErro = ItensDeContratoSrv_Altera_BD(objItensDeContratoSRV, alComando(1))
        If lErro <> SUCESSO Then gError 195662
    
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 195663

    lErro = CF("WorkFlow_Trata_Transacao", "SRV", objItensDeContratoSRV.objTela, gsUsuario, objItensDeContratoSRV)
    If lErro <> SUCESSO Then gError 195664
          
    ItensDeContratoSrv_Grava = SUCESSO
    
    Exit Function
    
Erro_ItensDeContratoSrv_Grava:
    
    ItensDeContratoSrv_Grava = gErr
    
    Select Case gErr
    
        Case 195655
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 195656
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195657, 195660, 195662, 195664

        Case 195658, 195659
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMCONTRATOSRV", gErr, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.lCodigo)
        
        Case 195661
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSDECONTRATOSRV", gErr, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.lCodigo)
        
        Case 195663
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183936)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function ItensDeContratoSrvObjVinc_Lock(ByVal objItensDeContratoSRV As ClassItensDeContratoSrv) As Long
'Faz "lock" nos objetos associados a item de contrato de serviço

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim objProduto As New ClassProduto
Dim objRastroLote As New ClassRastreamentoLote
Dim iIndice As Integer
Dim objItensContSrvNumSerie As ClassItensContSrvNumSerie
Dim objItensDeContratoSrvProd As ClassItensDeContratoSrvProd
Dim objItensContrato As New ClassItensDeContrato
Dim objItensDeContratoSRV1 As ClassItensDeContratoSrv
Dim dQuantidade As Double
Dim colItensDeContratoSRV As New Collection

On Error GoTo Erro_ItensDeContratoSrvObjVinc_Lock

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195673
    Next

    objProduto.sCodigo = objItensDeContratoSRV.sProduto
    
    'Locka  o produto
    lErro = CF("Produto_Lock_Shared", alComando(1), objProduto)
    If lErro <> SUCESSO Then gError 195674
    
    If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 195675

    If gobjSRV.iVerificaLote = VERIFICA_LOTE Then
    
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_OP Then

            objRastroLote.sProduto = objItensDeContratoSRV.sProduto
            objRastroLote.sCodigo = objItensDeContratoSRV.sLote
            objRastroLote.iFilialOP = objItensDeContratoSRV.iFilialOP
    
            'Lê rastreamento do lote a partir do produto, filialOP e código do lote passados
            lErro = CF("RastreamentoLote_Le1", alComando(2), objRastroLote)
            If lErro <> SUCESSO And lErro <> 71889 Then gError 195676
            
            If lErro <> SUCESSO Then gError 195677

        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then

            objRastroLote.sProduto = objItensDeContratoSRV.sProduto

            For Each objItensContSrvNumSerie In objItensDeContratoSRV.colNumSerie
            
                objRastroLote.sCodigo = objItensContSrvNumSerie.sNumSerie

                'Lê rastreamento do lote a partir do produto, filialOP e código do lote passados
                lErro = CF("RastreamentoLote_Le1", alComando(3), objRastroLote)
                If lErro <> SUCESSO And lErro <> 71889 Then gError 195678
                
                If lErro <> SUCESSO Then gError 195679

            Next

        End If

    End If
    
    objItensContrato.lNumIntDoc = objItensDeContratoSRV.lNumIntItemContrato
    
    lErro = CF("ItensDeContrato_Lock_Shared", objItensContrato)
    If lErro <> SUCESSO And lErro <> 195669 Then gError 195680
    
    If lErro <> SUCESSO Then gError 195681

    lErro = CF("ItensDeContratoSrv_Le_NumIntItemContrato", objItensDeContratoSRV, colItensDeContratoSRV)
    If lErro <> SUCESSO And lErro <> 195693 Then gError 195697

    For Each objItensDeContratoSRV1 In colItensDeContratoSRV
    
        If objItensDeContratoSRV1.lCodigo <> objItensDeContratoSRV.lCodigo Then
            dQuantidade = dQuantidade + objItensDeContratoSRV1.dQuantidade
        End If
    
    Next

    If objItensContrato.dQuantidade < objItensDeContratoSRV.dQuantidade + dQuantidade Then gError 195704

    'Para cada item do contrato
    For Each objItensDeContratoSrvProd In objItensDeContratoSRV.colProduto

        objProduto.sCodigo = objItensDeContratoSrvProd.sProduto
        
        'Locka  o produto
        lErro = CF("Produto_Lock_Shared", alComando(4), objProduto)
        If lErro <> SUCESSO Then gError 195682

        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 195683

    Next

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensDeContratoSrvObjVinc_Lock = SUCESSO

    Exit Function

Erro_ItensDeContratoSrvObjVinc_Lock:

    ItensDeContratoSrvObjVinc_Lock = gErr

    Select Case gErr

        Case 195673
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195674, 195676, 195678, 195680, 195682

        Case 195675, 195683
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, objProduto.sCodigo)

        Case 195677
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_CADASTRADO_RASTREAMENTO", gErr, objItensDeContratoSRV.sLote, objItensDeContratoSRV.sProduto, objItensDeContratoSRV.iFilialOP)

        Case 195679
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMSERIE_NAO_CADASTRADO", gErr, objItensDeContratoSRV.sProduto, objItensDeContratoSRV.sLote)

        Case 195681
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMCONTRATO_NAO_CADASTRADO", gErr)

        Case 195704
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_ITEMCONTRATO_MENOR_SRV", gErr, objItensContrato.dQuantidade, objItensDeContratoSRV.dQuantidade + dQuantidade)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195684)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function ItensDeContratoSrv_Insere_BD(ByVal objItensDeContratoSRV As ClassItensDeContratoSrv) As Long
'Insere nova Garantia

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim alComando(1 To 3) As Long
Dim objItensDeContratoSrvProd As ClassItensDeContratoSrvProd
Dim objItensContSrvNumSerie As ClassItensContSrvNumSerie

On Error GoTo Erro_ItensDeContratoSrv_Insere_BD

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195685
    Next

    'Gera o número interno automático da Solicitacao
    lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_ITENSDECONTRATOSRV_INT", lNumIntDoc)
    If lErro <> SUCESSO Then gError 195686

    lErro = Comando_Executar(alComando(1), "INSERT INTO ItensDeContratoSrv (NumIntDoc, Codigo, FilialEmpresa, NumIntItemContrato, Quantidade, Lote, FilialOP, TipoGarantia, GarantiaTotal) VALUES (?,?,?,?,?,?,?,?,?)", lNumIntDoc, objItensDeContratoSRV.lCodigo, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.lNumIntItemContrato, objItensDeContratoSRV.dQuantidade, objItensDeContratoSRV.sLote, objItensDeContratoSRV.iFilialOP, objItensDeContratoSRV.lTipoGarantia, objItensDeContratoSRV.iGarantiaTotal)
    If lErro <> AD_SQL_SUCESSO Then gError 196587

    For Each objItensDeContratoSrvProd In objItensDeContratoSRV.colProduto

        lErro = Comando_Executar(alComando(2), "INSERT INTO ItensDeContratoSrvProd (NumIntItemContratoSrv, Produto) VALUES (?,?)", lNumIntDoc, objItensDeContratoSrvProd.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 195688

    Next
    
    For Each objItensContSrvNumSerie In objItensDeContratoSRV.colNumSerie

        lErro = Comando_Executar(alComando(3), "INSERT INTO ItensDeContratoSrvNumSerie (NumIntItemContratoSrv, NumSerie) VALUES (?,?)", lNumIntDoc, objItensContSrvNumSerie.sNumSerie)
        If lErro <> AD_SQL_SUCESSO Then gError 195689

    Next
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensDeContratoSrv_Insere_BD = SUCESSO

    Exit Function

Erro_ItensDeContratoSrv_Insere_BD:

    ItensDeContratoSrv_Insere_BD = gErr

    Select Case gErr

        Case 195685
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195686

        Case 195687
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSDECONTRATOSRV", gErr, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.lCodigo)

        Case 195688
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSDECONTRATOSRVPROD", gErr, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.lCodigo, objItensDeContratoSrvProd.sProduto)
        
        Case 195689
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSCONTSRVNUMSERIE", gErr, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.lCodigo, objItensContSrvNumSerie.sNumSerie)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195690)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function ItensDeContratoSrv_Altera_BD(objItensDeContratoSRV As ClassItensDeContratoSrv, ByVal lComando As Long) As Long
'Altera uma Solicitacao de Servico

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lNumIntDocAux As Long
Dim iIndice As Integer
Dim objItensDeContratoSrvProd As ClassItensDeContratoSrvProd
Dim alComando(1 To 5) As Long
Dim objItensDeContratoSrv_BD As New ClassItensDeContratoSrv
Dim objSolicSRV As New ClassSolicSRV
Dim objProduto As New ClassProduto
Dim objItensSolicSrv As ClassItensSolicSRV
Dim objItensContSrvNumSerie_BD As ClassItensContSrvNumSerie
Dim objItensContSrvNumSerie As ClassItensContSrvNumSerie
Dim iAchou As Integer
Dim objItensDeContratoSrvProd_BD As ClassItensDeContratoSrvProd
Dim sNumSerie As String
Dim iPrazo As Integer
Dim lNumIntItemContratoSrv As Long

On Error GoTo Erro_ItensDeContratoSrv_Altera_BD

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195686
    Next

    objItensDeContratoSrv_BD.iFilialEmpresa = objItensDeContratoSRV.iFilialEmpresa
    objItensDeContratoSrv_BD.lCodigo = objItensDeContratoSRV.lCodigo

    lErro = CF("ItensDeContratoSrv_Le", objItensDeContratoSrv_BD)
    If lErro <> SUCESSO And lErro <> 195584 Then gError 195687

    'se encontrou alguma solicitacao envolvendo o contrato e produto em questao
    lErro = CF("ItensSolicSRV_Le_Contrato", objSolicSRV, objItensDeContratoSrv_BD.iFilialEmpresa, objItensDeContratoSrv_BD.sCodigoContrato, objItensDeContratoSrv_BD.sProduto)
    If lErro <> SUCESSO And lErro <> 195596 Then gError 195688
    
    If lErro = SUCESSO Then
    
        'se o produto foi alterado ==> erro (nao pode alterar)
        If objItensDeContratoSrv_BD.sProduto <> objItensDeContratoSRV.sProduto Then gError 195689
    
        objProduto.sCodigo = objItensDeContratoSRV.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195690
            
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_OP Then
        
            For Each objItensSolicSrv In objSolicSRV.colItens
            
                If objItensSolicSrv.sLote = objItensDeContratoSrv_BD.sLote And objItensSolicSrv.iFilialOP = objItensDeContratoSrv_BD.iFilialOP Then
                
                    If objItensDeContratoSRV.sLote <> objItensDeContratoSrv_BD.sLote Or objItensDeContratoSRV.iFilialOP <> objItensDeContratoSrv_BD.iFilialOP Then gError 195691
                    
                End If
                
            Next
        
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then

            'pesquisa as solicitacoes de servico
            For Each objItensSolicSrv In objSolicSRV.colItens
            
                'pesquisa os numeros de serie que estao no BD
                For Each objItensContSrvNumSerie_BD In objItensDeContratoSrv_BD.colNumSerie
            
                    'se algum num. serie participa de uma solicitacao ==> nao pode ser excluido
                    If objItensSolicSrv.sLote = objItensContSrvNumSerie_BD.sNumSerie Then
                    
                        iAchou = 0
                    
                        For Each objItensContSrvNumSerie In objItensDeContratoSRV.colNumSerie
                        
                            If objItensContSrvNumSerie.sNumSerie = objItensContSrvNumSerie_BD.sNumSerie Then
                                iAchou = 1
                                Exit For
                            End If
                            
                        Next
                        
                        'se nao achou o num. serie ==> que foi excluido mas nao podia pois é usado em SolicSRV ==> erro
                        If iAchou = 0 Then gError 195692
                            
                    End If
                
                Next
                
            Next

        End If

    End If

    lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensDeContratoSrv SET NumIntItemContrato = ?, Quantidade = ?, Lote = ?, FilialOP = ?, TipoGarantia = ?, GarantiaTotal = ?", lComando, _
    objItensDeContratoSRV.lNumIntItemContrato, objItensDeContratoSRV.dQuantidade, objItensDeContratoSRV.sLote, objItensDeContratoSRV.iFilialOP, objItensDeContratoSRV.lTipoGarantia, objItensDeContratoSRV.iGarantiaTotal)
    If lErro <> AD_SQL_SUCESSO Then gError 195693

    'Remove de colProduto do BD os itens presente na colProduto da tela
    For Each objItensDeContratoSrvProd In objItensDeContratoSRV.colProduto
        For iIndice = objItensDeContratoSrv_BD.colProduto.Count To 1 Step -1
            If objItensDeContratoSrvProd.sProduto = objItensDeContratoSrv_BD.colProduto(iIndice).sProduto Then
                objItensDeContratoSrv_BD.colProduto.Remove iIndice
                Exit For
            End If
        Next
    Next

    'Remove de colNumSerie do BD os itens presente na colNumSerie da tela
    For Each objItensContSrvNumSerie In objItensDeContratoSRV.colNumSerie
        For iIndice = objItensDeContratoSrv_BD.colNumSerie.Count To 1 Step -1
            If objItensContSrvNumSerie.sNumSerie = objItensDeContratoSrv_BD.colNumSerie(iIndice).sNumSerie Then
                objItensDeContratoSrv_BD.colNumSerie.Remove iIndice
                Exit For
            End If
        Next
    Next

    For Each objItensDeContratoSrvProd In objItensDeContratoSRV.colProduto

        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntItemContratoSrv FROM ItensDeContratoSrvProd WHERE NumIntItemContratoSrv = ? AND Produto = ?", 0, lNumIntItemContratoSrv, objItensDeContratoSRV.lNumIntDoc, objItensDeContratoSrvProd.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 195694

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195695

        If lErro = AD_SQL_SEM_DADOS Then
        
            lErro = Comando_Executar(alComando(4), "INSERT INTO ItensDeContratoSrvProd (NumIntItemContratoSrv, Produto) VALUES (?,?)", objItensDeContratoSRV.lNumIntDoc, objItensDeContratoSrvProd.sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 195696
            
        End If

    Next

    'Exclui os ítens de Solicitacao que foram excluídos na tela
    For Each objItensDeContratoSrvProd_BD In objItensDeContratoSrv_BD.colProduto

        lErro = Comando_Executar(alComando(5), "DELETE FROM ItensDeContratoSrvProd WHERE NumIntItemContratoSrv = ? AND Produto = ?", objItensDeContratoSrvProd_BD.lNumIntItemContratoSrv, objItensDeContratoSrvProd_BD.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 195697
    
    Next
    
    For Each objItensContSrvNumSerie In objItensDeContratoSRV.colNumSerie

        sNumSerie = String(STRING_LOTE_RASTREAMENTO, 0)

        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumSerie FROM ItensDeContratoSrvNumSerie WHERE NumIntItemContratoSrv = ? AND NumSerie = ?", 0, sNumSerie, objItensDeContratoSRV.lNumIntDoc, objItensContSrvNumSerie.sNumSerie)
        If lErro <> AD_SQL_SUCESSO Then gError 195698

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195699

        If lErro = AD_SQL_SEM_DADOS Then
        
            lErro = Comando_Executar(alComando(4), "INSERT INTO ItensDeContratoSrv (NumIntItemContratoSrv, NumSerie) VALUES (?,?)", objItensDeContratoSRV.lNumIntDoc, objItensContSrvNumSerie.sNumSerie)
            If lErro <> AD_SQL_SUCESSO Then gError 195700
            
        End If
            
    Next

    'Exclui os ítens de GarantiaNumSerie que foram excluídos na tela
    For Each objItensContSrvNumSerie_BD In objItensDeContratoSrv_BD.colNumSerie

        lErro = Comando_Executar(alComando(5), "DELETE FROM ItensDeContratoSrvNumSerie WHERE NumIntItemContratoSrv = ? AND NumSerie = ?", objItensContSrvNumSerie_BD.lNumIntItemContratoSrv, objItensContSrvNumSerie_BD.sNumSerie)
        If lErro <> AD_SQL_SUCESSO Then gError 195701
    
    Next
    
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensDeContratoSrv_Altera_BD = SUCESSO

    Exit Function

Erro_ItensDeContratoSrv_Altera_BD:

    ItensDeContratoSrv_Altera_BD = gErr

    Select Case gErr

        Case 195686
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195687, 195688, 195690
        
        Case 195689
            Call Rotina_Erro(vbOKOnly, "ERRO_PROD_EM_SOLICSRV_GARANTIA", gErr, objItensDeContratoSrv_BD.sProduto)

        Case 195691
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_EM_SOLICSRV_GARANTIA", gErr, objItensDeContratoSrv_BD.sProduto, objItensDeContratoSrv_BD.sLote, objItensDeContratoSrv_BD.iFilialOP)

        Case 195692
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMSERIE_EM_SOLICSRV_GARANTIA", gErr, objItensDeContratoSrv_BD.sProduto, objItensContSrvNumSerie_BD.sNumSerie)

        Case 195693
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITENSDECONTRATOSRV", gErr, objItensDeContratoSRV.iFilialEmpresa, objItensDeContratoSRV.lCodigo)
        
        Case 195694, 195695
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSDECONTRATOSRVPROD1", gErr, objItensDeContratoSRV.lNumIntDoc, objItensDeContratoSrvProd.sProduto)

        Case 195696
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSDECONTRATOSRVPROD1", gErr, objItensDeContratoSRV.lNumIntDoc, objItensDeContratoSrvProd.sProduto)

        Case 195697
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMCONTRATOSRVPROD", gErr, objItensDeContratoSRV.lNumIntDoc, objItensDeContratoSrvProd.sProduto)

        Case 195698, 195699
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSDECONTRATOSRVNUMSERIE1", gErr, objItensDeContratoSRV.lNumIntDoc, objItensContSrvNumSerie.sNumSerie)

        Case 195700
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSCONTSRVNUMSERIE1", gErr, objItensDeContratoSRV.lNumIntDoc, objItensContSrvNumSerie.sNumSerie)

        Case 195701
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMCONTRATOSRVNUMSERIE1", gErr, objItensContSrvNumSerie.lNumIntItemContratoSrv, objItensContSrvNumSerie.sNumSerie)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195702)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function RoteiroSRV_Grava(ByVal objRoteiroSRV As ClassRoteiroSRV) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RoteiroSRV_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 194875

    lErro = CF("RoteiroSRV_Grava_EmTrans", objRoteiroSRV)
    If lErro <> SUCESSO Then gError 194876

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 194877

    RoteiroSRV_Grava = SUCESSO

    Exit Function

Erro_RoteiroSRV_Grava:

    RoteiroSRV_Grava = gErr

    Select Case gErr

        Case 194875
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 194876

        Case 194877
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194878)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function RoteiroSRV_Grava_EmTrans(ByVal objRoteiroSRV As ClassRoteiroSRV) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_RoteiroSRV_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194880
    Next

    'Le a tabelaRoteiroSRV
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteiroSRV WHERE Servico= ?  AND Versao= ? ", _
                0, lNumIntDoc, objRoteiroSRV.sServico, objRoteiroSRV.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 194881

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194882

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 194883

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RoteiroSRV SET Descricao= ? , DataCriacao= ? , DataUltModificacao= ? , Quantidade= ? , " & _
                "UM= ? , Autor= ? , Duracao= ? ", alComando(0), _
                objRoteiroSRV.sDescricao, objRoteiroSRV.dtDataCriacao, objRoteiroSRV.dtDataUltModificacao, objRoteiroSRV.dQuantidade, _
                objRoteiroSRV.sUM, objRoteiroSRV.sAutor, objRoteiroSRV.iDuracao)
        If lErro <> AD_SQL_SUCESSO Then gError 194884

        objRoteiroSRV.lNumIntDoc = lNumIntDoc
        
        lErro = RoteiroSRVOper_Exclui(objRoteiroSRV)
        If lErro <> SUCESSO Then gError 194906

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_ROTEIROSRV", lNumIntDoc)
        If lErro <> SUCESSO Then gError 194879

        objRoteiroSRV.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO RoteiroSRV( NumIntDoc, Servico, Versao, Descricao, " & _
                "DataCriacao, DataUltModificacao, Quantidade, UM, Autor, " & _
                "Duracao) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                objRoteiroSRV.lNumIntDoc, objRoteiroSRV.sServico, objRoteiroSRV.sVersao, objRoteiroSRV.sDescricao, _
                objRoteiroSRV.dtDataCriacao, objRoteiroSRV.dtDataUltModificacao, objRoteiroSRV.dQuantidade, objRoteiroSRV.sUM, objRoteiroSRV.sAutor, _
                objRoteiroSRV.iDuracao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194885

    End If
    
    lErro = RoteiroSRVOper_Insere(objRoteiroSRV)
    If lErro <> SUCESSO Then gError 194907

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteiroSRV_Grava_EmTrans = SUCESSO

    Exit Function

Erro_RoteiroSRV_Grava_EmTrans:

    RoteiroSRV_Grava_EmTrans = gErr

    Select Case gErr

        Case 194879, 194906, 194907

        Case 194880
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194881, 194882
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRV", gErr)

        Case 194883
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ROTEIROSRV", gErr)

        Case 194884
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ROTEIROSRV", gErr)

        Case 194885
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ROTEIROSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194886)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function RoteiroSRV_Exclui(ByVal objRoteiroSRV As ClassRoteiroSRV) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RoteiroSRV_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 194864

    lErro = CF("RoteiroSRV_Exclui_EmTrans", objRoteiroSRV)
    If lErro <> SUCESSO Then gError 194865

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 194866

    RoteiroSRV_Exclui = SUCESSO

    Exit Function

Erro_RoteiroSRV_Exclui:

    RoteiroSRV_Exclui = gErr

    Select Case gErr

        Case 194864
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 194865

        Case 194866
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194867)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function RoteiroSRV_Exclui_EmTrans(ByVal objRoteiroSRV As ClassRoteiroSRV) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_RoteiroSRV_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194868
    Next

    'Le a tabelaRoteiroSRV
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteiroSRV WHERE Servico= ? AND Versao= ? ", _
                0, lNumIntDoc, objRoteiroSRV.sServico, objRoteiroSRV.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 194869

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194870

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 194871
    
    objRoteiroSRV.lNumIntDoc = lNumIntDoc

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 194872

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RoteiroSRV", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 194873
    
    lErro = RoteiroSRVOper_Exclui(objRoteiroSRV)
    If lErro <> SUCESSO Then gError 194892

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteiroSRV_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_RoteiroSRV_Exclui_EmTrans:

    RoteiroSRV_Exclui_EmTrans = gErr

    Select Case gErr

        Case 194868
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194869, 194870
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRV", gErr)

        Case 194871
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTEIROSRV_NAO_CADASTRADO", gErr, objRoteiroSRV.sServico, objRoteiroSRV.sVersao)

        Case 194872
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ROTEIROSRV", gErr)

        Case 194873
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ROTEIROSRV", gErr)
            
        Case 194892

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194874)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Public Function RoteiroSRVOper_Insere(ByVal objRoteiroSRV As ClassRoteiroSRV) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objRoteiroSRVOper As ClassRoteiroSRVOper

On Error GoTo Erro_RoteiroSRVOper_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194908
    Next
    
    For Each objRoteiroSRVOper In objRoteiroSRV.colOperacoes
    
        objRoteiroSRVOper.lNumIntDocRotSRV = objRoteiroSRV.lNumIntDoc
    
        'Le a tabelaRoteiroSRVOper
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteiroSRVOper WHERE NumIntDoc= ? ", _
                    0, lNumIntDoc, objRoteiroSRVOper.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 194909
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194910
    
        'Se existir => Erro (Já deveria ter sido excluído)
        If lErro = AD_SQL_SUCESSO Then gError 194911

        'Se é uma operação nova, senão mantem a que vem da tela (Tem que ter sido excluído anteriormente)
        If objRoteiroSRVOper.lNumIntDoc = 0 Then

            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_ROTEIROSRVOPER", lNumIntDoc)
            If lErro <> SUCESSO Then gError 194912
    
            objRoteiroSRVOper.lNumIntDoc = lNumIntDoc
            
        End If

        lErro = Comando_Executar(alComando(1), "INSERT INTO RoteiroSRVOper( NumIntDoc, NumIntDocRotSrv, Seq, NumIntDocCompet, " & _
                "NumIntDocCT, Observacao, SeqPai, SeqArvore, Nivel, " & _
                "PosicaoArvore) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                objRoteiroSRVOper.lNumIntDoc, objRoteiroSRVOper.lNumIntDocRotSRV, objRoteiroSRVOper.iSeq, objRoteiroSRVOper.lNumIntDocCompet, _
                objRoteiroSRVOper.lNumIntDocCT, objRoteiroSRVOper.sObservacao, objRoteiroSRVOper.iSeqPai, objRoteiroSRVOper.iSeqArvore, objRoteiroSRVOper.iNivel, _
                objRoteiroSRVOper.iPosicaoArvore)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194913
        
        lErro = RoteiroSRVOperMaq_Insere(objRoteiroSRVOper)
        If lErro <> SUCESSO Then gError 194893

        lErro = RoteiroSRVOperMO_Insere(objRoteiroSRVOper)
        If lErro <> SUCESSO Then gError 194894

        lErro = RoteiroSRVOperMP_Insere(objRoteiroSRVOper)
        If lErro <> SUCESSO Then gError 194895

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteiroSRVOper_Insere = SUCESSO

    Exit Function

Erro_RoteiroSRVOper_Insere:

    RoteiroSRVOper_Insere = gErr

    Select Case gErr

        Case 194908
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194909, 194910
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPER", gErr)

        Case 194911
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTEIROSRVOPER_JA_CADASTRADO", gErr)

        Case 194912, 194893 To 194895

        Case 194913
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ROTEIROSRVOPER", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194914)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RoteiroSRVOper_Exclui(ByVal objRoteiroSRV As ClassRoteiroSRV) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objRoteiroSRVOper As ClassRoteiroSRVOper

On Error GoTo Erro_RoteiroSRVOper_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194896
    Next

    'Le a tabelaRoteiroSRVOper
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteiroSRVOper WHERE NumIntDocRotSrv= ? ", _
                0, lNumIntDoc, objRoteiroSRV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194897

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194898

     Do While lErro <> AD_SQL_SEM_DADOS
     
        Set objRoteiroSRVOper = New ClassRoteiroSRVOper
        
        objRoteiroSRVOper.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 194900
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RoteiroSRVOper", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 194901
        
        lErro = RoteiroSRVOperMaq_Exclui(objRoteiroSRVOper)
        If lErro <> SUCESSO Then gError 194931

        lErro = RoteiroSRVOperMO_Exclui(objRoteiroSRVOper)
        If lErro <> SUCESSO Then gError 194932

        lErro = RoteiroSRVOperMP_Exclui(objRoteiroSRVOper)
        If lErro <> SUCESSO Then gError 194933
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194899
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteiroSRVOper_Exclui = SUCESSO

    Exit Function

Erro_RoteiroSRVOper_Exclui:

    RoteiroSRVOper_Exclui = gErr

    Select Case gErr

        Case 194896
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194897, 194898, 194899
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPER", gErr)
        
        Case 194900
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ROTEIROSRVOPER", gErr)

        Case 194901
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ROTEIROSRVOPER", gErr)
            
        Case 194931 To 194933

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194902)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RoteiroSRVOperMaq_Insere(ByVal objRoteiroSRVOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objRoteiroSRVOperMaq As ClassRoteiroSRVOperMaq

On Error GoTo Erro_RoteiroSRVOperMaq_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194936
    Next
    
    For Each objRoteiroSRVOperMaq In objRoteiroSRVOper.colMaq
    
        objRoteiroSRVOperMaq.lNumIntDocOper = objRoteiroSRVOper.lNumIntDoc
    
        'Le a tabelaRoteiroSRVOperMaq
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteiroSRVOperMaq WHERE NumIntDoc= ? ", _
                    0, lNumIntDoc, objRoteiroSRVOperMaq.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 194937
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194938
        
        'Se existir => Erro (Já deveria ter sido excluído)
        If lErro = AD_SQL_SUCESSO Then gError 194939

        If objRoteiroSRVOperMaq.lNumIntDoc = 0 Then

            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_ROTEIROSRVOPERMAQ", lNumIntDoc)
            If lErro <> SUCESSO Then gError 194940
    
            objRoteiroSRVOperMaq.lNumIntDoc = lNumIntDoc
            
        End If

        lErro = Comando_Executar(alComando(1), "INSERT INTO RoteiroSRVOperMaq( NumIntDoc, NumIntDocOper, Seq, CodMaq, " & _
                "Horas, Qtd, Obs,FilialEmpMaq) VALUES (?,?,?,?,?,?,?,?)", _
                objRoteiroSRVOperMaq.lNumIntDoc, objRoteiroSRVOperMaq.lNumIntDocOper, objRoteiroSRVOperMaq.iSeq, objRoteiroSRVOperMaq.iCodMaq, _
                objRoteiroSRVOperMaq.dHoras, objRoteiroSRVOperMaq.iQtd, objRoteiroSRVOperMaq.sObs, objRoteiroSRVOperMaq.iFilialEmpMaq)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194941

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteiroSRVOperMaq_Insere = SUCESSO

    Exit Function

Erro_RoteiroSRVOperMaq_Insere:

    RoteiroSRVOperMaq_Insere = gErr

    Select Case gErr

        Case 194935

        Case 194936
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194937, 194938
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPERMAQ", gErr)

        Case 194939
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTEIROSRVOPERMAQ_JA_CADASTRADO", gErr)

        Case 194940

        Case 194941
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ROTEIROSRVOPERMAQ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194942)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RoteiroSRVOperMO_Insere(ByVal objRoteiroSRVOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objRoteiroSRVOperMO As ClassRoteiroSRVOperMO

On Error GoTo Erro_RoteiroSRVOperMO_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194964
    Next
    
    For Each objRoteiroSRVOperMO In objRoteiroSRVOper.colMO

        objRoteiroSRVOperMO.lNumIntDocOper = objRoteiroSRVOper.lNumIntDoc

        'Le a tabelaRoteiroSRVOperMO
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteiroSRVOperMO WHERE NumIntDocOper= ? ", _
                    0, lNumIntDoc, objRoteiroSRVOperMO.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 194965
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194966
    
        'Se existir => UPDATE, senão => INSERT
        If lErro = AD_SQL_SUCESSO Then gError 194967
        
        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_ROTEIROSRVOPERMO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 194968

        objRoteiroSRVOperMO.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO RoteiroSRVOperMO( NumIntDoc, NumIntDocOper, Seq, CodMO, " & _
                "Horas, Qtd, Obs) VALUES (?,?,?,?,?,?,?)", _
                objRoteiroSRVOperMO.lNumIntDoc, objRoteiroSRVOperMO.lNumIntDocOper, objRoteiroSRVOperMO.iSeq, objRoteiroSRVOperMO.iCodMO, _
                objRoteiroSRVOperMO.dHoras, objRoteiroSRVOperMO.iQtd, objRoteiroSRVOperMO.sObs)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194969

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteiroSRVOperMO_Insere = SUCESSO

    Exit Function

Erro_RoteiroSRVOperMO_Insere:

    RoteiroSRVOperMO_Insere = gErr

    Select Case gErr

        Case 194963

        Case 194964
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194965, 194966
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPERMO", gErr)

        Case 194967
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTEIROSRVOPERMO_JA_CADASTRADO", gErr)

        Case 194968

        Case 194969
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ROTEIROSRVOPERMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194970)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RoteiroSRVOperMP_Insere(ByVal objRoteiroSRVOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objRoteiroSRVOperMP As ClassRoteiroSRVOperMP

On Error GoTo Erro_RoteiroSRVOperMP_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194992
    Next
    
    For Each objRoteiroSRVOperMP In objRoteiroSRVOper.colMP

        objRoteiroSRVOperMP.lNumIntDocOper = objRoteiroSRVOper.lNumIntDoc

        'Le a tabelaRoteiroSRVOperMP
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteiroSRVOperMP WHERE NumIntDocOper= ? ", _
                    0, lNumIntDoc, objRoteiroSRVOperMP.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 194993
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194994
    
        'Se existir => UPDATE, senão => INSERT
        If lErro = AD_SQL_SUCESSO Then gError 194995

        If objRoteiroSRVOperMP.lNumIntDoc = 0 Then

            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_ROTEIROSRVOPERMP", lNumIntDoc)
            If lErro <> SUCESSO Then gError 194996
    
            objRoteiroSRVOperMP.lNumIntDoc = lNumIntDoc
            
        End If

        lErro = Comando_Executar(alComando(1), "INSERT INTO RoteiroSRVOperMP( NumIntDoc, NumIntDocOper, Seq, Produto, " & _
                "Quantidade, UM, Versao, Composicao, OBS) VALUES (?,?,?,?,?,?,?,?,?)", _
                objRoteiroSRVOperMP.lNumIntDoc, objRoteiroSRVOperMP.lNumIntDocOper, objRoteiroSRVOperMP.iSeq, objRoteiroSRVOperMP.sProduto, _
                objRoteiroSRVOperMP.dQuantidade, objRoteiroSRVOperMP.sUM, objRoteiroSRVOperMP.sVersao, objRoteiroSRVOperMP.iComposicao, objRoteiroSRVOperMP.sObs)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194997

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteiroSRVOperMP_Insere = SUCESSO

    Exit Function

Erro_RoteiroSRVOperMP_Insere:

    RoteiroSRVOperMP_Insere = gErr

    Select Case gErr

        Case 194991

        Case 194992
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194993, 194994
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPERMP", gErr)

        Case 194995
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTEIROSRVOPERMP_JA_CADASTRADO", gErr)

        Case 194996

        Case 194997
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ROTEIROSRVOPERMP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194998)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RoteiroSRVOperMaq_Exclui(ByVal objRoteiroSRVOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_RoteiroSRVOperMaq_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194924
    Next

    'Le a tabelaRoteiroSRVOperMaq
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteiroSRVOperMaq WHERE NumIntDocOper= ? ", _
                0, lNumIntDoc, objRoteiroSRVOper.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194925

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194926

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 194928
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RoteiroSRVOperMaq", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 194929
            
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194927
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteiroSRVOperMaq_Exclui = SUCESSO

    Exit Function

Erro_RoteiroSRVOperMaq_Exclui:

    RoteiroSRVOperMaq_Exclui = gErr

    Select Case gErr

        Case 194924
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194925, 194926, 194927
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPERMAQ", gErr)
        
        Case 194928
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ROTEIROSRVOPERMAQ", gErr)

        Case 194929
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ROTEIROSRVOPERMAQ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194930)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function RoteiroSRVOperMO_Exclui(ByVal objRoteiroSRVOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_RoteiroSRVOperMO_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194952
    Next

    'Le a tabelaRoteiroSRVOperMO
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteiroSRVOperMO WHERE NumIntDocOper= ? ", _
                0, lNumIntDoc, objRoteiroSRVOper.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194953

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194954

    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 194956
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RoteiroSRVOperMO", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 194957
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194955
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteiroSRVOperMO_Exclui = SUCESSO

    Exit Function

Erro_RoteiroSRVOperMO_Exclui:

    RoteiroSRVOperMO_Exclui = gErr

    Select Case gErr

        Case 194952
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194953, 194954, 194955
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPERMO", gErr)

        Case 194956
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ROTEIROSRVOPERMO", gErr)

        Case 194957
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ROTEIROSRVOPERMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194958)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function RoteiroSRVOperMP_Exclui(ByVal objRoteiroSRVOper As ClassRoteiroSRVOper) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_RoteiroSRVOperMP_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194980
    Next

    'Le a tabelaRoteiroSRVOperMP
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteiroSRVOperMP WHERE NumIntDocOper= ? ", _
                0, lNumIntDoc, objRoteiroSRVOper.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194981

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194982

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 194984
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RoteiroSRVOperMP", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 194985
   
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194983
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteiroSRVOperMP_Exclui = SUCESSO

    Exit Function

Erro_RoteiroSRVOperMP_Exclui:

    RoteiroSRVOperMP_Exclui = gErr

    Select Case gErr

        Case 194980
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194981, 194982, 194983
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSRVOPERMP", gErr)

        Case 194984
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ROTEIROSRVOPERMP", gErr)

        Case 194985
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ROTEIROSRVOPERMP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194986)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RelGarantia_Prepara(ByVal objGarantia As ClassGarantia, lNumIntRel As Long) As Long
'Monta a tabela temporária para o relatório de calendário de máquina
 
Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objFornecedor As New ClassFornecedor
Dim objProduto As New ClassProduto
Dim objFilialForn As New ClassFilialFornecedor
Dim sNF As String
Dim objTipoGarantia As New ClassTipoGarantia
Dim sLote As String
Dim objGarantiaLote As ClassGarantiaNumSerie
Dim objGarantiaServ As ClassGarantiaProduto

On Error GoTo Erro_RelGarantia_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 202238
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202239
    Next
 
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_REL_GARANTIA", lNumIntRel)
    If lErro <> SUCESSO Then gError 202240
    
    If objGarantia.lNumNotaFiscal <> 0 Then
        sNF = "Série: " & objGarantia.sSerie & " Número: " & CStr(objGarantia.lNumNotaFiscal)
    End If
    
    objProduto.sCodigo = objGarantia.sProduto
    
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 202241
   
    If objGarantia.lFornecedor <> 0 Then
    
        objFornecedor.lCodigo = objGarantia.lFornecedor
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 202242
        
        objFilialForn.lCodFornecedor = objGarantia.lFornecedor
        objFilialForn.iCodFilial = objGarantia.iFilialFornecedor
        
        'Verifica se já existe a Filial do Fornecedor lido na tabela FiliaisFornecedores
        lErro = CF("FilialFornecedor_Le", objFilialForn)
        If lErro <> SUCESSO And lErro <> 12929 Then gError 202243
        
        sNF = sNF & " Fornecedor: " & objFornecedor.sNomeReduzido & " Filial: " & objFilialForn.sNome

    End If
    
    If objGarantia.lTipoGarantia <> 0 Then
    
        objTipoGarantia.lCodigo = objGarantia.lTipoGarantia
        
        lErro = CF("TipoGarantia_Le", objTipoGarantia)
        If lErro <> SUCESSO And lErro <> 183849 Then gError 202244
    
    End If
    
    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
        sLote = objGarantia.sLote
    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
        If objGarantia.sLote <> "" Then
            sLote = objGarantia.sLote & SEPARADOR & objGarantia.iFilialOP
        End If
    ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
        iIndice = 0
        For Each objGarantiaLote In objGarantia.colNumSerie
            iIndice = iIndice + 1
            sLote = sLote & IIf(sLote <> "", IIf(iIndice <> objGarantia.colNumSerie.Count, ", ", " e "), "") & objGarantiaLote.sNumSerie
        Next
    Else
        sLote = ""
    End If
    
    'Insere em RelGarantia
    lErro = Comando_Executar(alComando(0), "INSERT INTO RelGarantia (NumIntRel,Codigo,Produto,DescProduto,DataVenda,Quantidade,Lote,NotaFiscal,TipoGarantia,TipoGarantiaDesc,UM, FilialEmpresa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
        lNumIntRel, objGarantia.lCodigo, objGarantia.sProduto, objProduto.sDescricao, objGarantia.dtDataVenda, objGarantia.dQuantidade, sLote, sNF, objGarantia.lTipoGarantia, objTipoGarantia.sDescricao, objProduto.sSiglaUMEstoque, objGarantia.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 202245
              
    If objGarantia.iGarantiaTotal = MARCADO Then
        Set objGarantiaServ = New ClassGarantiaProduto
        objGarantiaServ.iPrazo = objGarantia.iGarantiaTotalPrazo
        If objGarantia.colProduto.Count = 0 Then
            objGarantiaServ.lNumIntGarantia = -1
            objGarantiaServ.sProduto = "Integral"
        Else
            objGarantiaServ.lNumIntGarantia = -2
            objGarantiaServ.sProduto = "Demais Peças e Serviços"
        End If
        objGarantia.colProduto.Add objGarantiaServ
    End If
    
    iIndice = 0
    For Each objGarantiaServ In objGarantia.colProduto
        iIndice = iIndice + 1
        
        Set objProduto = New ClassProduto
       
        If objGarantiaServ.lNumIntGarantia >= 0 Then
        
            objProduto.sCodigo = objGarantia.sProduto
        
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 202246
        
        Else
            objProduto.sDescricao = objGarantiaServ.sProduto
        End If
    
        'Insere em RelGarantiaPeca
        lErro = Comando_Executar(alComando(0), "INSERT INTO RelGarantiaPecas (NumIntRel,Seq,Servico,ServicoDesc,Prazo) VALUES (?,?,?,?,?)", _
        lNumIntRel, iIndice, objProduto.sCodigo, objProduto.sDescricao, objGarantiaServ.iPrazo)
        If lErro <> SUCESSO Then gError 202247
    
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 202248
  
    RelGarantia_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelGarantia_Prepara:
 
    RelGarantia_Prepara = gErr
 
    Select Case gErr
 
        Case 202238
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 202239
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 202240 To 202244, 202246
 
        Case 202245
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELGARANTIA", gErr)
        
        Case 202247
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELGARANTIAPECA", gErr)
        
        Case 202248
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202249)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Function NFiscalFatPedSRV_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a gravação de uma nota fiscal fatura gerada a partir de um pedido de venda

Dim lErro As Long
Dim lTransacao As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NFiscalFatPedSRV_Grava
   
   'Abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 206181

    lErro = CF("NFiscalFatPedSRV_Grava_EmTrans", objNFiscal, objContabil)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
               
    If lErro = SUCESSO Then
               
        'Faz o commit
        lErro = Transacao_Commit()
        If lErro <> SUCESSO Then gError 206182
               
    Else
    
        If objNFiscal.lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 210147

            Call Transacao_Rollback
    
            lErro = CF("NFiscalPedido_AlteraDados", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 210148

        Else
            gError 210149
        
        End If
        
    End If
    
               
    NFiscalFatPedSRV_Grava = SUCESSO
    
    Exit Function
    
Erro_NFiscalFatPedSRV_Grava:

    NFiscalFatPedSRV_Grava = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206181
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 206182
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 210147 To 210149
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206183)
    
    End Select
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function NFiscalFatGarSRV_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a gravação de uma nota fiscal fatura gerada a partir de um pedido de venda

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_NFiscalFatGarSRV_Grava
   
   'Abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 206181

    lErro = CF("NFiscalFatGarSRV_Grava_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
               
    'Faz o commit
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 206182
               
    NFiscalFatGarSRV_Grava = SUCESSO
    
    Exit Function
    
Erro_NFiscalFatGarSRV_Grava:

    NFiscalFatGarSRV_Grava = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206181
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 206182
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206183)
    
    End Select
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function NFiscalFatPedSRV_Grava_EmTrans(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a gravação de uma nota fiscal fatura gerada a partir de um pedido de venda

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim alComando1(1 To 30) As Long
Dim alComando(1 To NUM_MAX_LCOMANDO_NFISCALFAT) As Long
Dim iIndice As Integer
Dim dQuantFaturada As Double
Dim dQuantNaoCancelada As Double
Dim objMovEstoque As ClassMovEstoque
Dim dQuantPedida As Double
Dim dQuantCancelada  As Double
Dim iStatus As Integer
Dim lNumIntDoc As Long
Dim bFaturouFaixa As Boolean, bBloqueio As Boolean

On Error GoTo Erro_NFiscalFatPedSRV_Grava_EmTrans
   
    'Abre os comandos
    For iIndice = LBound(alComando1) To UBound(alComando1)
        alComando1(iIndice) = Comando_Abrir()
        If alComando1(iIndice) = 0 Then gError 206184
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206185
    Next
        
    'Verifica se já existe nota com esses dados
    lErro = Comando_Executar(alComando1(30), "SELECT NumIntDoc FROM TipoDocInfoNFiscal WHERE FilialEmpresa = ? AND Serie = ? AND NumNotaFiscal = ? AND DataEmissao = ? AND Status <> ? AND (Tipo = ? OR Tipo = ?) ", lNumIntDoc, giFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, STATUS_EXCLUIDO, DOCINFO_NF_INT_ENTRADA, DOCINFO_NF_INT_SAIDA)
    If lErro <> AD_SQL_SUCESSO Then gError 206186
    
    lErro = Comando_BuscarPrimeiro(alComando1(30))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206187
    
    'Nova nota Fiscal
    If lErro = AD_SQL_SEM_DADOS Then
        
        objPedidoVenda.lCodigo = objNFiscal.lNumPedidoVenda
        objPedidoVenda.iFilialEmpresa = objNFiscal.iFilialPedido
    
        'Lê e faz lock no Pedido de Venda
        lErro = CF("PedidoServico_Lock", objPedidoVenda, alComando1(1))
        If lErro <> SUCESSO And lErro <> 36986 Then gError ERRO_SEM_MENSAGEM
        
        If lErro <> SUCESSO Then gError 206188
    
        'Se a Filial de faturamento for diferente da filialempresa atual --> erro.
        If objPedidoVenda.iFilialEmpresaFaturamento <> giFilialEmpresa Then gError 206189
    
        'Lê os bloqueios do Pedido de venda
        lErro = CF("BloqueiosGen_Testa", MAPBLOQGEN_TIPOTELA_PEDIDOSRV, objPedidoVenda, bBloqueio)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        If bBloqueio Then gError 206190
            
        'Nos ítens da NF coloca o NumIntDoc do PedidoVenda
        For iIndice = 1 To objNFiscal.colItensNF.Count
            objNFiscal.colItensNF(iIndice).lNumIntPedVenda = objPedidoVenda.lNumIntDoc
            objNFiscal.colItensNF(iIndice).iTipo = ITEMNF_TIPO_PECA
        Next
        
        For iIndice = 1 To objNFiscal.colItensSRV.Count
            objNFiscal.colItensSRV(iIndice).lNumIntPedVenda = objPedidoVenda.lNumIntDoc
            objNFiscal.colItensSRV(iIndice).iTipo = ITEMNF_TIPO_SERVICO
        Next
        
    End If
    
    'Chama a função de gravação de uma nota fiscal
    lErro = CF("NFiscalFatura_Grava_Trans", alComando(), objNFiscal)
    If lErro <> SUCESSO And lErro <> 61401 Then gError ERRO_SEM_MENSAGEM
    
    'Se for uma nota nova
    If lErro = SUCESSO Then
    
        lErro = ItensPSRV_AtualizaQuantFaturada(alComando(3), alComando(4), objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Atualiza estatística de QuantPedida em ProdutoFilial
        lErro = ProdutoFilial_NF_QuantPedidaSRV(objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Grava a Estatística do Cliente
        lErro = CF("FilialCliente_Grava_Estatistica", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Grava a Estatística do Produto Vendido
        lErro = CF("ProdutoVendido_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_DOC)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'Le Itens de PV
        lErro = CF("PedidoServico_Le_Itens", objPedidoVenda)
        If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
                   
        'Verifica se tem quantidade não cancelada
        lErro = PedSRV_QuantNaoCancelada(objPedidoVenda, dQuantNaoCancelada)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Se nao tiver quantidade nao cancelada, baixa o Pedido
        If Abs(dQuantNaoCancelada) < QTDE_ESTOQUE_DELTA Then
    
            lErro = PedidoServico_Baixar(objPedidoVenda)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
         
        End If
        
        'Verifica se o modulo de Livros Fiscais está Ativo
        If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                            
            'Grava o Livro Fiscal a partir da Nota Fiscal
            lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
    
'    Else
    
'horaentrada
        'trata da alteração dos dados da nota fiscal.
'        lErro = CF("NFiscal_Alteracao", objNFiscal)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
'    lErro = CF("ProjetoInfo_Grava", objNFiscal.objProjetoInfo, PRJ_CR_TIPO_NF, objNFiscal.lNumIntDoc)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("NF_Grava_RPS", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Grava os dados contábeis (contabilidade)
    lErro = objContabil.Contabil_Gravar_Registro(objNFiscal.lNumIntDoc, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataSaida, NAO_AVALIA_PELO_CUSTO_REAL_PRODUCAO, NAO_AVISA_LANCAMENTOS_CONTABILIZADOS, objNFiscal.lNumNotaFiscal, -objNFiscal.dValorLivroAux)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objNFiscal.objInfoAdic.lNumIntDoc = objNFiscal.lNumIntDoc
    objNFiscal.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_NF
    lErro = CF("InfoAdicionais_Grava", objNFiscal.objInfoAdic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Fecha os comandos
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscalFatPedSRV_Grava_EmTrans = SUCESSO
    
    Exit Function
    
Erro_NFiscalFatPedSRV_Grava_EmTrans:

    NFiscalFatPedSRV_Grava_EmTrans = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206184, 206185
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 206186, 206187
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case 206188
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
        
        Case 206189
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFATURAMENTO_DIFERENTE", gErr)
       
        Case 206190
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_BLOQUEIO", gErr, objPedidoVenda.lCodigo)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206191)
    
    End Select
    
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function NFiscalFatGarSRV_Grava_EmTrans(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a gravação de uma nota fiscal fatura gerada a partir de um pedido de venda

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim alComando1(0 To 10) As Long
Dim alComando(1 To NUM_MAX_LCOMANDO_NFISCALFAT) As Long
Dim iIndice As Integer
Dim dQuantFaturada As Double
Dim dQuantNaoCancelada As Double
Dim objMovEstoque As ClassMovEstoque
Dim dQuantPedida As Double
Dim dQuantCancelada  As Double
Dim iStatus As Integer
Dim lNumIntDoc As Long
Dim bFaturouFaixa As Boolean, bBloqueio As Boolean

On Error GoTo Erro_NFiscalFatGarSRV_Grava_EmTrans
   
    'Abre os comandos
    For iIndice = LBound(alComando1) To UBound(alComando1)
        alComando1(iIndice) = Comando_Abrir()
        If alComando1(iIndice) = 0 Then gError 206184
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206185
    Next
    
    'Chama a função de gravação de uma nota fiscal
    lErro = CF("NFiscalFatura_Grava_Trans", alComando(), objNFiscal)
    If lErro <> SUCESSO And lErro <> 61401 Then gError ERRO_SEM_MENSAGEM
    
    'Se for uma nota nova
    If lErro = SUCESSO Then
    
        'Grava a associação com os apontamentos
        lErro = NFiscalFatGarSRV_Grava_AP(alComando1(), objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Grava a Estatística do Cliente
        lErro = CF("FilialCliente_Grava_Estatistica", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Grava a Estatística do Produto Vendido
        lErro = CF("ProdutoVendido_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_DOC)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
        'Verifica se o modulo de Livros Fiscais está Ativo
        If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                            
            'Grava o Livro Fiscal a partir da Nota Fiscal
            lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
    
    Else
    
'horaentrada
        'trata da alteração dos dados da nota fiscal.
        lErro = CF("NFiscal_Alteracao", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
'    lErro = CF("ProjetoInfo_Grava", objNFiscal.objProjetoInfo, PRJ_CR_TIPO_NF, objNFiscal.lNumIntDoc)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("NF_Grava_RPS", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Grava os dados contábeis (contabilidade)
    lErro = objContabil.Contabil_Gravar_Registro(objNFiscal.lNumIntDoc, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataSaida, NAO_AVALIA_PELO_CUSTO_REAL_PRODUCAO, NAO_AVISA_LANCAMENTOS_CONTABILIZADOS, objNFiscal.lNumNotaFiscal, -objNFiscal.dValorLivroAux)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objNFiscal.objInfoAdic.lNumIntDoc = objNFiscal.lNumIntDoc
    objNFiscal.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_NF
    lErro = CF("InfoAdicionais_Grava", objNFiscal.objInfoAdic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Fecha os comandos
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscalFatGarSRV_Grava_EmTrans = SUCESSO
    
    Exit Function
    
Erro_NFiscalFatGarSRV_Grava_EmTrans:

    NFiscalFatGarSRV_Grava_EmTrans = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206184, 206185
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 206186, 206187
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case 206188
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
        
        Case 206189
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFATURAMENTO_DIFERENTE", gErr)
       
        Case 206190
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_BLOQUEIO", gErr, objPedidoVenda.lCodigo)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206191)
    
    End Select
    
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function ItensPSRV_AtualizaQuantFaturada(ByVal lComando As Long, ByVal lComando1 As Long, ByVal objNFiscal As ClassNFiscal, Optional ByVal bCancNF As Boolean = False) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iStatus As Integer
Dim dQuantFaturada As Double
Dim dQuantCancelada As Double
Dim dQuantPedida As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim alComando(0 To 1) As Long
Dim sUM As String
Dim dFator As Double
Dim objItemNF As ClassItemNF
Dim dQuantFatAMais As Double
Dim iTipo As Integer, colItens As New Collection
Dim dQtd As Double

On Error GoTo Erro_ItensPSRV_AtualizaQuantFaturada

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206192
    Next

    For iTipo = ITEMNF_TIPO_SERVICO To ITEMNF_TIPO_PECA

        If iTipo = ITEMNF_TIPO_SERVICO Then
            Set colItens = objNFiscal.colItensSRV
        Else
            Set colItens = objNFiscal.colItens
        End If
    
        'Para cada item de pedido de venda que foi utilizado na Nota Fiscal
        For iIndice = 1 To colItens.Count
            
            Set objItemNF = colItens.Item(iIndice)
            
            If bCancNF Then
                dQtd = -objItemNF.dQuantidade
            Else
                dQtd = objItemNF.dQuantidade
            End If
            
            lErro = CF("MovEstoque_ItemNF_Customizado", objItemNF, objNFiscal)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If Not objItemNF Is Nothing Then
            
                sUM = String(STRING_UM_SIGLA, 0)
                iStatus = STATUS_NAO_ATENDIDO
                
                'Lê o item de pedido de venda
                lErro = Comando_ExecutarPos(lComando, "SELECT QuantFaturada,QuantCancelada,Quantidade,UnidadeMed FROM ItensPedidoSRV WHERE NumIntDoc = ?", 0, dQuantFaturada, dQuantCancelada, dQuantPedida, sUM, objItemNF.lNumIntItemPedVenda)
                If lErro <> AD_SQL_SUCESSO Then gError 206193
                
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206194
                If lErro <> AD_SQL_SUCESSO Then gError 206195
                        
                lErro = CF("UM_Conversao", objItemNF.iClasseUM, sUM, objItemNF.sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                           
                If Arredonda_Estoque(dQuantPedida * dFator) - dQuantCancelada - dQuantFaturada - CDbl(Formata_Estoque(dQtd)) < QTDE_ESTOQUE_DELTA + (Arredonda_Estoque(dQuantPedida * dFator) * (objItemNF.dPercentMenosReceb)) Then iStatus = STATUS_ATENDIDO
        
                dQuantFatAMais = dQuantFaturada + CDbl(Formata_Estoque(dQtd)) - Arredonda_Estoque(dQuantPedida * dFator)
                If dQuantFatAMais < 0 Then dQuantFatAMais = 0
            
                'Atualiza a quantidade faturada
                lErro = Comando_ExecutarPos(lComando1, "UPDATE ItensPedidoSRV SET QuantFaturada = QuantFaturada + ?, Status =? ", lComando, Arredonda_Estoque(dQtd / dFator), iStatus)
                If lErro <> AD_SQL_SUCESSO Then gError 206196
    
            End If
    
        Next
        
    Next

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensPSRV_AtualizaQuantFaturada = SUCESSO
    
    Exit Function
    
Erro_ItensPSRV_AtualizaQuantFaturada:

    ItensPSRV_AtualizaQuantFaturada = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 206192
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
         Case 206193, 206194
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)
        
        Case 206195
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_NAO_CADASTRADO", gErr, objNFiscal.colItensNF(iIndice).lNumIntItemPedVenda)
            
        Case 206196
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206197)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function ItemOS_Gera_Operacoes(ByVal objItemOS As ClassItemOS) As Long

Dim lErro As Long
Dim objRoteiroSRV As New ClassRoteiroSRV
Dim objProduto As New ClassProduto
Dim dFatorUMProduto As Double, dFatorUMRoteiro As Double, dFator As Double
Dim objItemOSOperacoes As ClassItemOSOperacoes
Dim objItemOSOpMO As ClassItemOSOperacoesMO, objItemOSOpMaq As ClassItemOSOperacoesMaquinas, objItemOSOpMP As ClassItemOSOperacoesPecas
Dim objOperacao As ClassRoteiroSRVOper
Dim objOpMO As ClassRoteiroSRVOperMO, objOpMaq As ClassRoteiroSRVOperMaq, objOpMP As ClassRoteiroSRVOperMP
Dim objMaquina As ClassMaquinas
Dim objCompetenciaPadrao As New ClassCompetencias

On Error GoTo Erro_ItemOS_Gera_Operacoes

    Set objItemOS.colItemOSOperacoes = New Collection
    
    objRoteiroSRV.sServico = objItemOS.sServico
    objRoteiroSRV.sVersao = objItemOS.sVersao
    
    lErro = CF("RoteiroSRV_Le", objRoteiroSRV)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205833
    
    If lErro = SUCESSO Then
    
        objProduto.sCodigo = objItemOS.sServico
        
        'Lê o produto para descobrir as unidades de medidas associadas
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 205834
    
        'Descobre o fator de conversao da UM do Item da OP p/UM de estoque do produto
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemOS.sSiglaUM, objProduto.sSiglaUMEstoque, dFatorUMProduto)
        If lErro <> SUCESSO Then gError 205835
        
        'Descobre o fator de conversao da UM do Roteiro p/UM de estoque do produto
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objRoteiroSRV.sUM, objProduto.sSiglaUMEstoque, dFatorUMRoteiro)
        If lErro <> SUCESSO Then gError 205836
    
        dFator = (objItemOS.dQuantidade * dFatorUMProduto) / (objRoteiroSRV.dQuantidade * dFatorUMRoteiro)

        For Each objOperacao In objRoteiroSRV.colOperacoes
                        
            Set objItemOSOperacoes = New ClassItemOSOperacoes
            Set objMaquina = New ClassMaquinas
            
            objItemOS.colItemOSOperacoes.Add objItemOSOperacoes
            
            objItemOSOperacoes.lNumIntDocCompet = objOperacao.lNumIntDocCompet
            objItemOSOperacoes.lNumIntDocCT = objOperacao.lNumIntDocCT
            objItemOSOperacoes.sObservacao = objOperacao.sObservacao
            objItemOSOperacoes.iNivel = objOperacao.iNivel
            objItemOSOperacoes.iSeq = objOperacao.iSeq
            objItemOSOperacoes.iSeqPai = objOperacao.iSeqPai
            objItemOSOperacoes.sServico = objItemOS.sServico
            objItemOSOperacoes.sVersao = objItemOS.sVersao
            objItemOSOperacoes.dQuantidade = objItemOS.dQuantidade
            objItemOSOperacoes.sUMedida = objItemOS.sSiglaUM
            objItemOSOperacoes.iConsideraCarga = MARCADO
            
            For Each objOpMP In objOperacao.colMP
            
                Set objItemOSOpMP = New ClassItemOSOperacoesPecas
                
                objItemOSOpMP.sProduto = objOpMP.sProduto
                If objOpMP.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
                    objItemOSOpMP.dQuantidade = objOpMP.dQuantidade * dFator
                Else
                    objItemOSOpMP.dQuantidade = objOpMP.dQuantidade
                End If
                objItemOSOpMP.sUMProduto = objOpMP.sUM
                objItemOSOpMP.iComposicao = objOpMP.iComposicao
                objItemOSOpMP.sVersaoKitComp = objOpMP.sVersao
            
                objItemOSOperacoes.colPecas.Add objItemOSOpMP
            
            Next
            
            For Each objOpMO In objOperacao.colMO
            
                Set objItemOSOpMO = New ClassItemOSOperacoesMO
                
                objItemOSOpMO.dHoras = objOpMO.dHoras * dFator
                objItemOSOpMO.lCodigoMO = objOpMO.iCodMO
            
                objItemOSOperacoes.colMO.Add objItemOSOpMO
            
            Next
            
            For Each objOpMaq In objOperacao.colMaq
            
                Set objItemOSOpMaq = New ClassItemOSOperacoesMaquinas
                Set objMaquina = New ClassMaquinas
                
                objMaquina.iCodigo = objOpMaq.iCodMaq
                objMaquina.iFilialEmpresa = objOpMaq.iFilialEmpMaq
                
                lErro = CF("Maquinas_Le", objMaquina)
                If lErro <> SUCESSO And lErro <> 103090 Then gError 205837
                
                objItemOSOpMaq.dHoras = objOpMaq.dHoras * dFator
                objItemOSOpMaq.iQuantidade = objOpMaq.iQtd
                objItemOSOpMaq.lNumIntDocMaq = objMaquina.lNumIntDoc
            
                objItemOSOperacoes.colMaquinas.Add objItemOSOpMaq
            
            Next
            
        Next
        
    Else
    
        'Le a Competencia Padrao
        lErro = CF("Competencias_Le_Padrao", objCompetenciaPadrao)
        If lErro <> SUCESSO And lErro <> 137844 Then gError 206409
        
        If lErro = SUCESSO Then
        
            Set objItemOSOperacoes = New ClassItemOSOperacoes
        
            objItemOSOperacoes.lNumIntDocCompet = objCompetenciaPadrao.lNumIntDoc
            objItemOSOperacoes.lNumIntDocCT = objCompetenciaPadrao.lNumIntDocCT

            objItemOSOperacoes.iNivel = 0
            objItemOSOperacoes.iSeq = 1
            objItemOSOperacoes.iNivelRoteiro = 1
            objItemOSOperacoes.iSeqRoteiro = 1
            objItemOSOperacoes.sServico = objItemOS.sServico
            objItemOSOperacoes.sVersao = objItemOS.sVersao
            objItemOSOperacoes.dQuantidade = objItemOS.dQuantidade
            objItemOSOperacoes.sUMedida = objItemOS.sSiglaUM
            objItemOSOperacoes.iConsideraCarga = DESMARCADO
            
            objItemOS.colItemOSOperacoes.Add objItemOSOperacoes
        End If
    End If

    ItemOS_Gera_Operacoes = SUCESSO
    
    Exit Function
    
Erro_ItemOS_Gera_Operacoes:

    ItemOS_Gera_Operacoes = gErr
    
    Select Case gErr
    
        Case 205833 To 205837, 206409
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205838)
    
    End Select
    
    Exit Function

End Function

Public Function MovEstoqueAp_Grava(ByVal objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim alComando(1 To 6) As Long
Dim iIndice As Integer
Dim objOSAp As ClassOSAp
Dim objOSApMaq As ClassOSApMaq
Dim objOSApMO As ClassOSApMO
Dim lNumIntDoc As Long

On Error GoTo Erro_MovEstoqueAp_Grava
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205858
    Next

    Set objOSAp = objMovEstoque.objApontamento
    
    'Só grava pela primeira vez, não altera, logo não deixa excluir na gravação
'    lErro = OSApMaq_Exclui(alComando(1), alComando(2), objOSAp.lCodigo, objOSAp.iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 205853
'
'    lErro = OSApMO_Exclui(alComando(3), alComando(4), objOSAp.lCodigo, objOSAp.iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 205854
    
    For Each objOSApMaq In objOSAp.colMaq
        objOSApMaq.lCodigoAp = objOSAp.lCodigo
        objOSApMaq.iFilialEmpresa = objOSAp.iFilialEmpresa
        
        lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_INT_OSAPMAQ", lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objOSApMaq.lNumIntDoc = lNumIntDoc
        lErro = OSApMaq_Insere(alComando(5), objOSApMaq)
        If lErro <> SUCESSO Then gError 205855
    Next

    For Each objOSApMO In objOSAp.colMO
        objOSApMO.lCodigoAp = objOSAp.lCodigo
        objOSApMO.iFilialEmpresa = objOSAp.iFilialEmpresa
        
        lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_INT_OSAPMO", lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objOSApMO.lNumIntDoc = lNumIntDoc
        lErro = OSApMO_Insere(alComando(6), objOSApMO)
        If lErro <> SUCESSO Then gError 205856
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    MovEstoqueAp_Grava = SUCESSO

    Exit Function

Erro_MovEstoqueAp_Grava:

    MovEstoqueAp_Grava = gErr

    Select Case gErr
    
        Case 205853 To 205856

        Case 205858
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205857)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
End Function

Public Function OSApMaq_Insere(ByVal lComando As Long, ByVal objOSApMaq As ClassOSApMaq) As Long

Dim lErro As Long

On Error GoTo Erro_OSApMaq_Insere

    lErro = Comando_Executar(lComando, "INSERT INTO OSApMaq( CodigoAp, FilialEmpresa, Seq, NumIntDocMaq, " & _
            "HorasGastas, Quantidade, OS, ProdutoOS, NumIntDoc) VALUES (?,?,?,?,?,?,?,?,?)", _
            objOSApMaq.lCodigoAp, objOSApMaq.iFilialEmpresa, objOSApMaq.iSeq, objOSApMaq.lNumIntDocMaq, _
            objOSApMaq.dHorasGastas, objOSApMaq.iQuantidade, objOSApMaq.sOS, objOSApMaq.sProdutoOS, objOSApMaq.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205859

    OSApMaq_Insere = SUCESSO

    Exit Function

Erro_OSApMaq_Insere:

    OSApMaq_Insere = gErr

    Select Case gErr
        
        Case 205859
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OSAPMAQ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205860)

    End Select

    Exit Function

End Function

Public Function OSApMO_Insere(ByVal lComando As Long, ByVal objOSApMO As ClassOSApMO) As Long

Dim lErro As Long

On Error GoTo Erro_OSApMO_Insere

    lErro = Comando_Executar(lComando, "INSERT INTO OSApMO( CodigoAp, FilialEmpresa, Seq, CodigoMO, " & _
            "HorasGastas, OS, ProdutoOS, NumIntDoc) VALUES (?,?,?,?,?,?,?,?)", _
            objOSApMO.lCodigoAp, objOSApMO.iFilialEmpresa, objOSApMO.iSeq, objOSApMO.lCodigoMO, _
            objOSApMO.dHorasGastas, objOSApMO.sOS, objOSApMO.sProdutoOS, objOSApMO.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205861


    OSApMO_Insere = SUCESSO

    Exit Function

Erro_OSApMO_Insere:

    OSApMO_Insere = gErr

    Select Case gErr

        Case 205861
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OSAPMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205862)

    End Select

    Exit Function

End Function

Public Function OSApMO_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim iAux As Integer

On Error GoTo Erro_OSApMO_Exclui

    'Le a tabela OSApMO
    lErro = Comando_ExecutarPos(lComando1, "SELECT Seq FROM OSApMO WHERE CodigoAp= ? AND FilialEmpresa= ? ", _
                0, iAux, lCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 205863

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205864

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 205865
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM OSApMO", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 205866
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205867
        
    Loop

    OSApMO_Exclui = SUCESSO

    Exit Function

Erro_OSApMO_Exclui:

    OSApMO_Exclui = gErr

    Select Case gErr

        Case 205863, 205864, 205867
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OSAPMO", gErr)

        Case 205865
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_OSAPMO", gErr)

        Case 205866
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OSAPMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205868)

    End Select

    Exit Function

End Function

Public Function OSApMaq_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim iAux As Integer

On Error GoTo Erro_OSApMaq_Exclui

    'Le a tabela OSApMaq
    lErro = Comando_ExecutarPos(lComando1, "SELECT Seq FROM OSApMaq WHERE CodigoAp= ? AND FilialEmpresa= ? ", _
                0, iAux, lCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 205869

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205870

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 205871
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM OSApMaq", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 205872
        
        'Busca PróxiMaq
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205873
        
    Loop

    OSApMaq_Exclui = SUCESSO

    Exit Function

Erro_OSApMaq_Exclui:

    OSApMaq_Exclui = gErr

    Select Case gErr

        Case 205869, 205870, 205873
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OSAPMAQ", gErr)

        Case 205871
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_OSAPMAQ", gErr)

        Case 205872
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OSAPMaq", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205874)

    End Select

    Exit Function

End Function

Function MovEstoque_ApSRV_Grava(ByVal objMovEstoque As ClassMovEstoque, ByVal objContabil As ClassContabil) As Long
'Grava o Movimento de Estoque de Requisição de Produção

Dim lErro As Long
Dim lTransacao As Long
Dim objItemMovEstoque As ClassItemMovEstoque
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim objItemOS As ClassItemOS
Dim lCodigoMovEst As Long, lNumIntDoc As Long

On Error GoTo Erro_MovEstoque_ApSRV_Grava

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 205875
    
    'Abre Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205876
    Next
    
    'Le a tabela OSAp
    lErro = Comando_ExecutarPos(alComando(1), "SELECT CodigoMovEst FROM ApontamentoSRV WHERE FilialEmpresa = ? AND Codigo = ?", 0, lCodigoMovEst, objMovEstoque.objApontamento.iFilialEmpresa, objMovEstoque.objApontamento.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 205841

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205841

    If lErro <> AD_SQL_SEM_DADOS Then
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ApontamentoSRV SET Data = ?, Hora = ?", alComando(1), objMovEstoque.objApontamento.dtData, CDbl(objMovEstoque.objApontamento.dtHora))
        If lErro <> AD_SQL_SUCESSO Then gError 205841
        
        If objMovEstoque.colItens.Count > 0 Then objMovEstoque.lCodigo = lCodigoMovEst
       
    Else
    
        If objMovEstoque.colItens.Count > 0 Then
            lErro = CF("MovEstoque_Automatico_EmTransacao", objMovEstoque.iFilialEmpresa, lCodigoMovEst)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            objMovEstoque.lCodigo = lCodigoMovEst
            
            objMovEstoque.objApontamento.lCodigoMovEst = lCodigoMovEst
        End If
    
        lErro = CF("Config_ObterNumInt", "SRVConfig", "NUM_PROX_INT_APONTAMENTOSRV", lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objMovEstoque.objApontamento.lNumIntDoc = lNumIntDoc
        
        lErro = Comando_Executar(alComando(2), "INSERT INTO ApontamentoSRV (NumIntDoc,Codigo,CodigoMovEst,FilialEmpresa,Data,Hora,Cliente) VALUES (?,?,?,?,?,?,?)", objMovEstoque.objApontamento.lNumIntDoc, objMovEstoque.objApontamento.lCodigo, objMovEstoque.lCodigo, objMovEstoque.iFilialEmpresa, objMovEstoque.dtData, CDbl(objMovEstoque.dtHora), objMovEstoque.lCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 205841
        
        lErro = CF("MovEstoqueAp_Grava", objMovEstoque)
        If lErro <> SUCESSO Then gError 205879
        
    End If
    
    If objMovEstoque.colItens.Count > 0 Then
    
        For Each objItemMovEstoque In objMovEstoque.colItens
            Set objItemOS = New ClassItemOS
            objItemOS.sServico = objItemMovEstoque.sProdutoOP
            objItemOS.sCodigo = objItemMovEstoque.sOPCodigo
            objItemOS.iFilialEmpresa = objMovEstoque.iFilialEmpresa
            lErro = CF("ItemOS_Le", objItemOS)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205878
            
            'Armazena referencia p/ ItemOP em ItemMovEstoque
            objItemMovEstoque.lNumIntDocOrigem = objItemOS.lNumIntDoc
            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMOS
        Next
       
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then gError 205878
        
    End If
        
    
    If Not (objContabil Is Nothing) Then
    
        'Grava os dados contábeis (contabilidade)
        lErro = objContabil.Contabil_Gravar_Registro(objMovEstoque.objApontamento.lNumIntDoc, 0, 0, objMovEstoque.objApontamento.dtData)
        If lErro <> SUCESSO Then gError 205880

    End If
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirmacao da transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 205881

    MovEstoque_ApSRV_Grava = SUCESSO

    Exit Function

Erro_MovEstoque_ApSRV_Grava:

    MovEstoque_ApSRV_Grava = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 205875
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 205876
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205877 To 205880

        Case 205881
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 205841
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTAMENTOSRV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205882)

    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function MovEstoque_ApSRV_Exclui(ByVal objMovEstoque As ClassMovEstoque, ByVal objContabil As ClassContabil) As Long
'Exclui a Requisição de producao

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 7) As Long
Dim iIndice As Integer, lCodigoMovEst As Long, lNumIntDocAP As Long, lNumNF As Long

On Error GoTo Erro_MovEstoque_ApSRV_Exclui

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 205883

    'Abertura comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205884
    Next

    If objMovEstoque.iFilialEmpresa = 0 Then objMovEstoque.iFilialEmpresa = giFilialEmpresa

    'Le a tabela OSAp
    lErro = Comando_ExecutarPos(alComando(5), "SELECT CodigoMovEst, NumIntDoc FROM ApontamentoSRV WHERE Codigo= ? AND FilialEmpresa= ? ", _
                0, lCodigoMovEst, lNumIntDocAP, objMovEstoque.objApontamento.lCodigo, objMovEstoque.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 205869

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205870

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 205871

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(5))
    If lErro <> SUCESSO Then gError 205871
    
    'Le a tabela ItemNFItemAP
    lErro = Comando_Executar(alComando(7), "SELECT N.NumNotaFiscal FROM ItemNFItemAPSRV AS X,ItensNFiscal AS I, NFiscal AS N WHERE N.NumIntDoc = I.NumIntNF AND I.NumIntDoc = X.NumIntItemNF AND N.Status <> 7 AND X.NumIntDocAP = ? ", _
                lNumNF, lNumIntDocAP)
    If lErro <> AD_SQL_SUCESSO Then gError 205869

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(7))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205870

    'Se não existir => ERRO
    If lErro <> AD_SQL_SEM_DADOS Then gError 206432

    'Lê os ítens do Movimento de Estoque
    lErro = CF("MovimentoEstoqueAp_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205885
       
    lErro = OSApMaq_Exclui(alComando(1), alComando(2), objMovEstoque.objApontamento.lCodigo, objMovEstoque.objApontamento.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 205886
    
    lErro = OSApMO_Exclui(alComando(3), alComando(4), objMovEstoque.objApontamento.lCodigo, objMovEstoque.objApontamento.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 205887

    If objMovEstoque.lCodigo <> 0 Then
        
        'Redefine os movimento para fazer o cancelamento
        objMovEstoque.iOperacao = EXCLUSAO_MOVIMENTO
            
        'Grava o movimento de estoque
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then gError 89878

    End If
    
    If Not (objContabil Is Nothing) Then

        'exclui os dados contábeis pendentes ou atualizados (se houverem) relativos a lNumIntDocOrigem
        lErro = objContabil.Contabil_Exclui(objMovEstoque.objApontamento.lNumIntDoc, CONTABIL_EXCLUI_EFETIVO)
        If lErro <> SUCESSO Then gError 89879
        
    End If

    lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM ApontamentoSRV", alComando(5))
    If lErro <> AD_SQL_SUCESSO Then gError 205872
    
    'Confirmacao da transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 205889

    'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    MovEstoque_ApSRV_Exclui = SUCESSO

    Exit Function

Erro_MovEstoque_ApSRV_Exclui:

    MovEstoque_ApSRV_Exclui = gErr

    Select Case gErr

        Case 205883
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 205884
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205885 To 205888

        Case 205889
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 205869 To 205872
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTAMENTOSRV", gErr)
            
        Case 206432
            Call Rotina_Erro(vbOKOnly, "ERRO_APONTAMENTOSRV_FATURADO", gErr, lNumNF)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205890)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function

Function ProdutoFilial_NF_QuantPedidaSRV(ByVal objNFiscal As ClassNFiscal, Optional ByVal bSubtrai As Boolean = True) As Long
'Atualiza Subtraindo a quantPedida do Produto em UM de Venda na tabela ProdutosFilial
 
Dim lErro As Long
Dim dQuantPedida As Double
Dim dQuantItem As Double
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim objProdutoINF As New ClassProduto
Dim dFator As Double
Dim objItemPV As ClassItemPedido
Dim dQtdePV As Double
 
On Error GoTo Erro_ProdutoFilial_NF_QuantPedidaSRV

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206173
    Next
    
    For Each objItemNF In objNFiscal.colItensNF
    
        lErro = CF("MovEstoque_ItemNF_Customizado", objItemNF, objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If Not objItemNF Is Nothing Then
    
            objProdutoINF.sCodigo = objItemNF.sProduto
            lErro = CF("Produto_Le", objProdutoINF)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            If lErro <> SUCESSO Then gError 206174
    
            For iIndice = 0 To objItemNF.colItensRomaneioGrade.Count
           
                If objItemNF.colItensRomaneioGrade.Count = 0 Then
           
                    objProduto.sCodigo = objItemNF.sProduto
                    dQuantItem = objItemNF.dQuantidade
                    
                ElseIf iIndice = 0 Then
                    iIndice = 1
                    objProduto.sCodigo = objItemNF.colItensRomaneioGrade(iIndice).sProduto
                    dQuantItem = objItemNF.colItensRomaneioGrade(iIndice).dQuantidade
                    
                End If
            
                If Not bSubtrai Then dQuantItem = -dQuantItem
            
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                If lErro <> SUCESSO Then gError 206175
            
                If objProdutoINF.iKitVendaComp <> MARCADO Then
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
                Else
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, objProduto.sSiglaUMVenda, dFator)
                End If
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'Calcula a quantidade a ser Subtraida a quantidade pedida do produto
                dQuantItem = dQuantItem * dFator
                                                   
                'Busca o Registro de Produto do Item e FiliaisProduto
                lErro = Comando_ExecutarPos(alComando(0), "SELECT QuantPedida FROM ProdutosFilial WHERE Produto = ? AND FilialEmpresa = ?", 0, dQuantPedida, objProduto.sCodigo, objNFiscal.iFilialPedido)
                If lErro <> AD_SQL_SUCESSO Then gError 206176
        
                lErro = Comando_BuscarPrimeiro(alComando(0))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206177
                If lErro <> AD_SQL_SUCESSO Then gError 206178 'Não encontrou
                              
                'Atualiza a quantidade pedida em ProdutosFilial para o Produto
                lErro = Comando_ExecutarPos(alComando(1), "UPDATE ProdutosFilial SET QuantPedida = QuantPedida - ?", alComando(0), dQuantItem)
                If lErro <> AD_SQL_SUCESSO Then gError 206179
            Next
            
        End If
            
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ProdutoFilial_NF_QuantPedidaSRV = SUCESSO
    
    Exit Function
    
Erro_ProdutoFilial_NF_QuantPedidaSRV:

    ProdutoFilial_NF_QuantPedidaSRV = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 206173
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 206174, 206175
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 206176, 206177
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", gErr, objNFiscal.iFilialEmpresa, objItemNF.sProduto)
        
        Case 206178
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", gErr, objItemNF.sProduto, objNFiscal.iFilialEmpresa)
        
        Case 206179
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PRODUTOSFILIAL", gErr, objNFiscal.iFilialEmpresa, objItemNF.sProduto)
                               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206180)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function CancelamentoNFiscal_Atualiza_PedidoSRV(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim objPSRV As New ClassPedidoDeVenda

On Error GoTo Erro_CancelamentoNFiscal_Atualiza_PedidoSRV

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206410
    Next
    
    lErro = ItensPSRV_AtualizaQuantFaturada(alComando(1), alComando(2), objNFiscal, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Atualiza estatística de QuantPedida em ProdutoFilial
    lErro = ProdutoFilial_NF_QuantPedidaSRV(objNFiscal, False)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objPSRV.lCodigo = objNFiscal.lNumPedidoVenda
    objPSRV.iFilialEmpresa = objNFiscal.iFilialPedido
    
    lErro = PedidoServico_Baixar(objPSRV, STATUS_ABERTO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    CancelamentoNFiscal_Atualiza_PedidoSRV = SUCESSO
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
Erro_CancelamentoNFiscal_Atualiza_PedidoSRV:

    CancelamentoNFiscal_Atualiza_PedidoSRV = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206410
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206411)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function ApontamentoSRV_Automatico(lNumAp As Long) As Long

Dim lErro As Long

On Error GoTo Erro_ApontamentoSRV_Automatico

    lErro = CF("Config_ObterAutomatico", "SRVConfig", "NUM_PROX_APONTAMENTOSRV", "ApontamentoSRV", "Codigo", lNumAp)
    If lErro <> SUCESSO Then gError 188770

    ApontamentoSRV_Automatico = SUCESSO

    Exit Function

Erro_ApontamentoSRV_Automatico:

    ApontamentoSRV_Automatico = gErr

    Select Case gErr

        Case 188770

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188771)

    End Select

    Exit Function

End Function

Function NFiscalFatGarSRV_Grava_AP(alComando() As Long, objNF As ClassNFiscal) As Long

Dim lErro As Long
Dim colAPDet As Collection
Dim objDet As ClassOSApDet
Dim objItemNF As ClassItemNF

On Error GoTo Erro_NFiscalFatGarSRV_Grava_AP

    Set colAPDet = objNF.objInfoUsu

    For Each objDet In colAPDet
    
        If objDet.iPeca = MARCADO Then
            Set objItemNF = objNF.colItens(objDet.iItemNF)
        Else
            Set objItemNF = objNF.colItensSRV(objDet.iItemNF)
        End If
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO ItemNFItemAPSRV (NumIntItemNF,TipoNumIntDocAP,NumIntDocItemAP,Quantidade,UM,NumIntDocAP) VALUES (?,?,?,?,?,?)", _
        objItemNF.lNumIntDoc, objDet.iTipo, objDet.lNumIntDoc, objDet.dQuantidade, objDet.sUM, objDet.lNumIntDocAP)
        If lErro <> AD_SQL_SUCESSO Then gError 183667
    
    Next

    NFiscalFatGarSRV_Grava_AP = SUCESSO

    Exit Function

Erro_NFiscalFatGarSRV_Grava_AP:

    NFiscalFatGarSRV_Grava_AP = gErr

    Select Case gErr

        Case 188770

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188771)

    End Select

    Exit Function

End Function

Function NotaFiscalSRV_Excluir(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iAux As Integer
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim colItens As New Collection

On Error GoTo Erro_NotaFiscalSRV_Excluir

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206440
    Next
    
    For Each objItemNF In objNF.colItensNF
        colItens.Add objItemNF
    Next
    For Each objItemNF In objNF.colItensSRV
        colItens.Add objItemNF
    Next
    
    For Each objItemNF In colItens

        'Le a tabela OSApMO
        lErro = Comando_ExecutarPos(alComando(1), "SELECT TipoNumIntDocAP FROM ItemNFItemApSRV WHERE NumIntItemNF = ? ", _
                    0, iAux, objItemNF.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 206441
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206442
    
        'Se não existir => ERRO
        Do While lErro <> AD_SQL_SEM_DADOS
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(1))
            If lErro <> SUCESSO Then gError 206443
        
            lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ItemNFItemApSRV", alComando(1))
            If lErro <> AD_SQL_SUCESSO Then gError 206444
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206445
            
        Loop
        
    Next

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NotaFiscalSRV_Excluir = SUCESSO
       
    Exit Function
    
Erro_NotaFiscalSRV_Excluir:

    NotaFiscalSRV_Excluir = gErr
    
    Select Case gErr
        
        Case 206440
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
               
        Case 206441, 206442, 206445
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFITEMAPSRV", gErr)

        Case 206443
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ITEMNFITEMAPSRV", gErr)

        Case 206444
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMNFITEMAPSRV", gErr)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206446)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function SolicitacaoSRV_Encerra(ByVal lNumIntDoc As Long) As Long
'Altera uma Solicitacao de Servico

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim lStatusBx As Long

On Error GoTo Erro_SolicitacaoSRV_Encerra

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211208
    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ? AND Valor = ?", lStatusBx, CAMPOSGENERICOS_SOLICSRV_STATUSITEM, STRING_BAIXADA)
    If lErro <> AD_SQL_SUCESSO Then gError 211209

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211210

    If lErro = AD_SQL_SUCESSO Then
           
        lErro = Comando_Executar(alComando(2), "UPDATE ItensSolicSRV SET StatusItem = ? WHERE NumIntSolicSRV = ?", lStatusBx, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 211211
        
    End If
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    SolicitacaoSRV_Encerra = SUCESSO

    Exit Function

Erro_SolicitacaoSRV_Encerra:

    SolicitacaoSRV_Encerra = gErr

    Select Case gErr

        Case 211208
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211209, 211210
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSGENERICOSVALORES", gErr, CAMPOSGENERICOS_SOLICSRV_STATUSITEM)
        
        Case 211211
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITENSSOLICSRV", gErr, lNumIntDoc)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211212)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function ItensPSRV_DesfazBaixa(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iStatus As Integer
Dim dQuantFaturada As Double
Dim dQuantCancelada As Double
Dim dQuantPedida As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim alComando(0 To 2) As Long
Dim sUM As String
Dim dFator As Double
Dim objItemPV As ClassItemPedido
Dim dQuantFatAMais As Double
Dim iTipo As Integer, colItens As New Collection
Dim dQtd As Double
Dim objProduto As New ClassProduto
Dim sProduto As String

On Error GoTo Erro_ItensPSRV_DesfazBaixa

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 210635
    Next

    sUM = String(STRING_UM_SIGLA, 0)
    iStatus = STATUS_NAO_ATENDIDO
    sProduto = String(STRING_PRODUTO, 0)
    
    'Lê o item de pedido de venda
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Produto, QuantFaturada,QuantCancelada,Quantidade,UnidadeMed FROM ItensPedidoSRV WHERE FilialEmpresa = ? AND CodPedido = ?", 0, sProduto, dQuantFaturada, dQuantCancelada, dQuantPedida, sUM, objPV.iFilialEmpresa, objPV.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 210636
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210637
            
            
    Do While lErro = AD_SQL_SUCESSO
            
        objProduto.sCodigo = sProduto
            
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        If lErro <> SUCESSO Then gError 210641
    
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dQtd = (dQuantPedida - dQuantFaturada - dQuantCancelada) * dFator
        
        'Atualiza a quantidade pedida em ProdutosFilial para o Produto
        lErro = Comando_Executar(alComando(2), "UPDATE ProdutosFilial SET QuantPedida = QuantPedida + ? WHERE Produto = ? AND FilialEmpresa = ?", dQtd, sProduto, objPV.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 210642
            
        'Atualiza o status
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensPedidoSRV SET Status =? ", alComando(0), iStatus)
        If lErro <> AD_SQL_SUCESSO Then gError 210638

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210639


    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensPSRV_DesfazBaixa = SUCESSO
    
    Exit Function
    
Erro_ItensPSRV_DesfazBaixa:

    ItensPSRV_DesfazBaixa = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 210635
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
         Case 210636, 210637, 210639
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDOSRV2", gErr)
        
        Case 210638
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDOSRV", gErr)

        Case 210641
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 210642
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PRODUTOSFILIAL", gErr, objPV.iFilialEmpresa, sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210640)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function PedidoSRV_Desfaz_Baixa(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lTransacao As Long

On Error GoTo Erro_PedidoSRV_Desfaz_Baixa

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 210643

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 210644
    Next
    
    lErro = ItensPSRV_DesfazBaixa(objPV)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = PedidoServico_Baixar(objPV, STATUS_ABERTO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 210645

    PedidoSRV_Desfaz_Baixa = SUCESSO
    
    Exit Function
    
Erro_PedidoSRV_Desfaz_Baixa:

    PedidoSRV_Desfaz_Baixa = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 210643
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 210644
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
               
        Case 210645
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210646)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function PedidosSRVBaixar_Batch(ByVal colPVInfo As Collection) As Long
'recebe colecao de ClassPVInfo
'transfere pedidos de venda p/pedidos de venda baixados

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 1) As Long
Dim lTransacao As Long
Dim objPVInfo As ClassPVInfo
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_PedidosSRVBaixar_Batch

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 210680
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 210681

    'Percorre a coleção de PedVendas
    For Each objPVInfo In colPVInfo
    
        If objPVInfo.iMarcada = MARCADO Then
        
            Set objPedidoVenda = New ClassPedidoDeVenda
        
            'Passa o identificador do Pedido
            objPedidoVenda.iFilialEmpresa = giFilialEmpresa
            objPedidoVenda.lCodigo = objPVInfo.lCodPedido
            
            'Le o Pedido
            
            lErro = CF("PedidoServico_Lock", objPedidoVenda, alComando(1))
            If lErro <> SUCESSO Then gError 210682
            
            'Lê os ítens do Pedido de Vendas
            lErro = CF("PedidoServico_Le_Itens", objPedidoVenda)
            If lErro <> SUCESSO Then gError 210683
            
            'baixa o pedido
            lErro = CF("PedidoServico_Baixar_Pedido", objPedidoVenda)
            If lErro <> SUCESSO Then gError 210684
            
        End If
            
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 210685

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    PedidosSRVBaixar_Batch = SUCESSO

    Exit Function

Erro_PedidosSRVBaixar_Batch:

    PedidosSRVBaixar_Batch = gErr

    Select Case gErr

        Case 210680
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 210681
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 210682 To 210684
        
        Case 210685
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210686)

    End Select
    
    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


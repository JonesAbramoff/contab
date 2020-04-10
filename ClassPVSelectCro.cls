VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function OPFichaProcesso_Le(objOP As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tOPFichaProcesso As typeOPFichaProcesso
Dim objOPFichaProcesso As New ClassOPFichaProcesso
Dim iIndice As Integer

On Error GoTo Erro_OPFichaProcesso_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 131191

    With tOPFichaProcesso

        'Aloca espaço nos buffers
        .sCodigoOP = String(STRING_ORDEM_DE_PRODUCAO, 0)
        .sAmostra = String(STRING_FICHAPROCESSO_AMOSTRA, 0)
        .sObservacao = String(STRING_FICHAPROCESSO_OBSERVACAO, 0)
        .sTela = String(STRING_FICHAPROCESSO_TELA, 0)

        'Le a Ficha de Processo
        lErro = Comando_Executar(lComando, "SELECT CodigoOP, FilialEmpresa, CargaDrays, Embalagem, Mistura, TempoMistura, " & _
                                "ProcessoAgua, ProcessoOleo, Amperagem, VelRotoresDe, VelRotoresAte, AberturaGate, TempCamaraZ1De, TempCamaraZ1Ate, TempCamaraZ2De, TempCamaraZ2Ate, " & _
                                "TempCamaraZ3De, TempCamaraZ3Ate, Tela, TempExtrusoraZ1De, TempExtrusoraZ1Ate, TempExtrusoraZ2De, TempExtrusoraZ2Ate, TempExtrusoraZ3De, TempExtrusoraZ3Ate, " & _
                                "TempExtrusoraZ4De, TempExtrusoraZ4Ate, TempExtrusoraZ5De, TempExtrusoraZ5Ate, TempExtrusoraZ6De, TempExtrusoraZ6Ate, TempExtrusoraZ7De, TempExtrusoraZ7Ate, " & _
                                "VelVariadorDe, VelVariadorAte, Observacao, Amostra " & _
                                "FROM OPFichaProcesso WHERE CodigoOP = ? AND FilialEmpresa = ? ", _
                                .sCodigoOP, .iFilialEmpresa, .iCargaDrays, .iEmbalagem, .iMistura, .iTempoMistura, _
                                .dProcessoAgua, .dProcessoOleo, .iAmperagem, .dVelRotoresDe, .dVelRotoresAte, .dAberturaGate, .dTempCamaraZDe(1), .dTempCamaraZAte(1), .dTempCamaraZDe(2), .dTempCamaraZAte(2), .dTempCamaraZDe(3), _
                                .dTempCamaraZAte(3), .sTela, .dTempExtrusoraZDe(1), .dTempExtrusoraZAte(1), .dTempExtrusoraZDe(2), .dTempExtrusoraZAte(2), .dTempExtrusoraZDe(3), .dTempExtrusoraZAte(3), .dTempExtrusoraZDe(4), .dTempExtrusoraZAte(4), _
                                .dTempExtrusoraZDe(5), .dTempExtrusoraZAte(5), .dTempExtrusoraZDe(6), .dTempExtrusoraZAte(6), .dTempExtrusoraZDe(7), .dTempExtrusoraZAte(7), .dVelVariadorDe, .dVelVariadorAte, .sObservacao, .sAmostra, _
                                objOP.sCodigo, objOP.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 131192

    End With

    'Busca o item
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131193
    If lErro = AD_SQL_SEM_DADOS Then gError 131194

    With objOPFichaProcesso

        .iCargaDrays = tOPFichaProcesso.iCargaDrays
        .iEmbalagem = tOPFichaProcesso.iEmbalagem
        .iFilialEmpresa = tOPFichaProcesso.iFilialEmpresa
        .iMistura = tOPFichaProcesso.iMistura
        .iTempoMistura = tOPFichaProcesso.iTempoMistura
        .sCodigoOP = tOPFichaProcesso.sCodigoOP

        .dProcessoAgua = tOPFichaProcesso.dProcessoAgua
        .dProcessoOleo = tOPFichaProcesso.dProcessoOleo
        .iAmperagem = tOPFichaProcesso.iAmperagem
        .dVelRotoresDe = tOPFichaProcesso.dVelRotoresDe
        .dVelRotoresAte = tOPFichaProcesso.dVelRotoresAte
        .dAberturaGate = tOPFichaProcesso.dAberturaGate

        For iIndice = 1 To FICHAPROCESSO_NUM_ZONAS_TC
            .dTempCamaraZDe(iIndice) = tOPFichaProcesso.dTempCamaraZDe(iIndice)
            .dTempCamaraZAte(iIndice) = tOPFichaProcesso.dTempCamaraZAte(iIndice)
        Next

        .sTela = tOPFichaProcesso.sTela

        For iIndice = 1 To FICHAPROCESSO_NUM_ZONAS_TE
            .dTempExtrusoraZDe(iIndice) = tOPFichaProcesso.dTempExtrusoraZDe(iIndice)
            .dTempExtrusoraZAte(iIndice) = tOPFichaProcesso.dTempExtrusoraZAte(iIndice)
        Next

        .dVelVariadorDe = tOPFichaProcesso.dVelVariadorDe
        .dVelVariadorAte = tOPFichaProcesso.dVelVariadorAte
        .sObservacao = tOPFichaProcesso.sObservacao
        .sAmostra = tOPFichaProcesso.sAmostra

        Set objOP.objInfoUsu = objOPFichaProcesso

    End With

    'Fecha comando
    Call Comando_Fechar(lComando)

    OPFichaProcesso_Le = SUCESSO

    Exit Function

Erro_OPFichaProcesso_Le:

    OPFichaProcesso_Le = gErr

    Select Case gErr

        Case 131191
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131192, 131193
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPFICHAPROCESSO", gErr)

        Case 131194

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemOPCarga_Le(ByVal objItemOPCarga As ClassItemOPCarga) As Long

Dim lErro As Long
Dim lComando As Long
Dim tItemOPCarga As typeItemOPCarga

On Error GoTo Erro_ItemOPCarga_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 131182

    With tItemOPCarga

        'Aloca espaço nos buffers
        .sProdutoBase = String(STRING_PRODUTO, 0)

        'Le a Carga do ItemOP
        lErro = Comando_Executar(lComando, "SELECT NumIntItemOP, ProdutoBase, KgProdBase, KgCarga, QtdCarga FROM ItensOPCarga WHERE NumIntItemOP = ? ", _
                                                .lNumIntItemOP, .sProdutoBase, .dKgProdBase, .dKgCarga, .dQtdCarga, objItemOPCarga.lNumIntItemOP)
        If lErro <> AD_SQL_SUCESSO Then gError 131183

    End With

    'Busca o item
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131184

    If lErro = AD_SQL_SEM_DADOS Then gError 131199

    With tItemOPCarga

        objItemOPCarga.sProdutoBase = .sProdutoBase
        objItemOPCarga.dKgProdBase = .dKgProdBase
        objItemOPCarga.dKgCarga = .dKgCarga
        objItemOPCarga.dQtdCarga = .dQtdCarga

    End With

    lErro = ItemOPCargaInsumo_Le(objItemOPCarga)
    If lErro <> SUCESSO Then gError 131186

    'Fecha comando
    Call Comando_Fechar(lComando)

    ItemOPCarga_Le = SUCESSO

    Exit Function

Erro_ItemOPCarga_Le:

    ItemOPCarga_Le = gErr

    Select Case gErr

        Case 131182
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131183, 131184
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPCARGA", gErr)

        Case 131185, 131186, 131199

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemOPCargaInsumo_Le(ByVal objItemOPCarga As ClassItemOPCarga) As Long

Dim lErro As Long
Dim lComando As Long
Dim tItemOPCargaInsumo As typeItemOPCargaInsumo
Dim objItemOPCargaInsumo As ClassItemOPCargaInsumo
Dim sDescricao As String

On Error GoTo Erro_ItemOPCargaInsumo_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 131187

    With tItemOPCargaInsumo

        'Aloca espaço nos buffers
        .sProduto = String(STRING_PRODUTO, 0)
        .sLote = String(STRING_LOTE_RASTREAMENTO, 0)
        sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

        'Le a Carga do ItemOP
        lErro = Comando_Executar(lComando, "SELECT I.NumIntDoc, I.NumIntItemOP, I.Produto, I.KgCarga, I.Seq, P.Descricao, I.Fornecedor, I.Lote FROM ItensOPCargaInsumo AS I, Produtos AS P  WHERE I.Produto = P.Codigo AND I.NumIntItemOP = ? ORDER BY Seq ", _
                                                .lNumIntDoc, .lNumIntItemOP, .sProduto, .dKgCarga, .iSeq, sDescricao, .lFornecedor, .sLote, objItemOPCarga.lNumIntItemOP)
        If lErro <> AD_SQL_SUCESSO Then gError 131188

    End With

    'Busca o item
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131189

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItemOPCargaInsumo = New ClassItemOPCargaInsumo

        With tItemOPCargaInsumo

            objItemOPCargaInsumo.lNumIntDoc = .lNumIntDoc
            objItemOPCargaInsumo.lNumIntItemOP = .lNumIntItemOP
            objItemOPCargaInsumo.sProduto = .sProduto
            objItemOPCargaInsumo.dKgCarga = .dKgCarga
            objItemOPCargaInsumo.iSeq = .iSeq
            objItemOPCargaInsumo.sDescricao = sDescricao
            objItemOPCargaInsumo.lFornecedor = .lFornecedor
            objItemOPCargaInsumo.sLote = .sLote

        End With

        objItemOPCarga.colInsumos.Add objItemOPCargaInsumo

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131190

    Loop

    'Fecha comando
    Call Comando_Fechar(lComando)

    ItemOPCargaInsumo_Le = SUCESSO

    Exit Function

Erro_ItemOPCargaInsumo_Le:

    ItemOPCargaInsumo_Le = gErr

    Select Case gErr

        Case 131187
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131188, 131189, 131190
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPCARGAINSUMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'#####################################################################
'Inserido por Wagner 21/09/05
Function PV_Le_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim sTabela As String
Dim sObservacao As String
Dim objPVCro As New ClassPedidoDeVendaCro

On Error GoTo Erro_PV_Le_Customizado

     'Abre Comando
     For iIndice = LBound(alComando) To UBound(alComando)
         alComando(iIndice) = Comando_Abrir()
         If alComando(iIndice) = 0 Then gError 140306
     Next
    
     sTabela = "PedidosDeVenda"
    
     'Tenta ler Pedido de Venda e Pedido de Venda Baixado
     For iIndice = 1 To 2
     
        sObservacao = String(STRING_PEDIDOVENDA_OBSERVACAO, 0)
        
        'Le o Pedido de Venda
        lErro = Comando_Executar(alComando(1), "SELECT Observacao FROM " & sTabela & " WHERE NumIntDoc= ? ", sObservacao, objPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 140307
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140308
    
        If lErro = AD_SQL_SUCESSO Then Exit For
                 
        sTabela = "PedidosDeVendaBaixados"
    
    Next
    
    'Não encontrou
    If lErro <> AD_SQL_SUCESSO Then gError 140309
    
    objPVCro.sObservacao = sObservacao
    
    Set objPV.objInfoUsu = objPVCro
          
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PV_Le_Customizado = SUCESSO

    Exit Function

Erro_PV_Le_Customizado:

    PV_Le_Customizado = gErr

        Select Case gErr

        Case 140306
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140307, 140308
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", gErr)

        Case 140309
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPV.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function
'#####################################################################

'###########################################################
'Inserido por Wagner 09/12/05
Public Function ItensOV_Le_Customizado(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim lComando As Long
Dim iEscolhido As Integer
Dim sConcorrente As String
Dim iCondPagto As Integer
Dim objItensOVCro As ClassItemOVCro
Dim objItensOV As ClassItemOV

On Error GoTo Erro_ItensOV_Le_Customizado

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 141360
    
    For Each objItensOV In objOrcamentoVenda.colItens
    
        sConcorrente = String(STRING_OBSERVACAO_OBSERVACAO, 0)

        '############################################
        'Alterado por Wagner 09/12/05
        'Lê os itens do orçamento passado
        lErro = Comando_Executar(lComando, "SELECT CondPagto, Escolhido, Concorrente FROM ItensOrcamentoVenda WHERE NumIntDoc = ? ", iCondPagto, iEscolhido, sConcorrente, objItensOV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 141361
        '############################################

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 141362
    
        If lErro = AD_SQL_SEM_DADOS Then gError 141363
        
        Set objItensOVCro = New ClassItemOVCro

        objItensOVCro.iCondPagto = iCondPagto
        objItensOVCro.iEscolhido = iEscolhido
        objItensOVCro.sConcorrente = sConcorrente

        Set objItensOV.objInfoUsu = objItensOVCro

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensOV_Le_Customizado = SUCESSO

    Exit Function

Erro_ItensOV_Le_Customizado:

    ItensOV_Le_Customizado = gErr

    Select Case gErr

        Case 131360
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131361, 131362
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOV", gErr)

        Case 131363
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTO_VENDA_SEM_ITENS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function
'########################################################################

'###########################################################
'Inserido por Wagner 09/12/05
'Public Function RastroLoteLista_Le(ByVal lFornecedor As Long, ByVal sProduto As String, ByVal sLote As String) As Long
'
'Dim lErro As Long
'Dim lComando As Long
'Dim iAux As Integer
'
'On Error GoTo Erro_RastroLoteLista_Le
'
'    'Abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 141748
'
'    lErro = Comando_Executar(lComando, "SELECT 1 FROM RastroLoteLista WHERE Fornecedor = ? AND Produto = ? AND Lote = ? ", iAux, lFornecedor, sProduto, sLote)
'    If lErro <> AD_SQL_SUCESSO Then gError 141749
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 141750
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 141751
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    RastroLoteLista_Le = SUCESSO
'
'    Exit Function
'
'Erro_RastroLoteLista_Le:
'
'    RastroLoteLista_Le = gErr
'
'    Select Case gErr
'
'        Case 141748
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 141749, 141750
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)
'
'        Case 141751
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'########################################################################

'#####################################################################
'Inserido por Wagner 09/02/06
Public Function PV_Verifica_Bloqueios(ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer, iBloqueiosAtivos As Integer, lComando As Long) As Long
'Verifica se tem bloqueios

Dim lErro As Long
Dim lCount As Long

On Error GoTo Erro_PV_Verifica_Bloqueios

    If lComando = 0 Then
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 141736
    End If

    'Busca Bloqueios de Crédito liberado
    lErro = Comando_Executar(lComando, "SELECT Count(*) FROM BloqueiosPV WHERE FilialEmpresa = ? AND PedidoDeVenda = ? AND DataLib = ? AND TipoDeBloqueio IN (4,5,6)", lCount, iFilialEmpresa, lCodigo, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 141737

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141738

    iBloqueiosAtivos = 0

    'se o bloqueio de credito estiver aprovado ==> retona a flag indicativa
    If lErro = AD_SQL_SUCESSO Then iBloqueiosAtivos = lCount

    'MsgBox (CStr(iFilialEmpresa) & " " & CStr(lCodigo) & " " & CStr(lCount) & " " & iBloqueiosAtivos)

    PV_Verifica_Bloqueios = SUCESSO

    Exit Function

Erro_PV_Verifica_Bloqueios:

    PV_Verifica_Bloqueios = gErr

    Select Case gErr

        Case 141736
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 141737, 141738
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPV", gErr, lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function
'#####################################################################



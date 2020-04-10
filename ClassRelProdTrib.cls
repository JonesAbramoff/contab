VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelProdTrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sUltProduto As String
Private lUltCliente As Long
Private iUltFilialCli As Integer
Private iUltFilialEmp As Integer
Private dUltPreco As Double

Public dValorAnt As Double

Public Function Obter_ICMSSTValor(ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dPreco As Double, dValor As Double) As Long

Dim lErro As Long
Dim objNFiscal As ClassNFiscal
Dim objItemNF As ClassItemNF
Dim dValorTotal As Double, objTribTab As ClassTribTab
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objProduto As ClassProduto
Dim iTipoNF As Integer, sSigla As String

On Error GoTo Erro_Obter_ICMSSTValor

    If sProduto = sUltProduto And lCliente = lUltCliente And iUltFilialCli = iFilialCli And iUltFilialEmp = iFilialEmpresa And Abs(dUltPreco - dPreco) < DELTA_VALORMONETARIO Then
    
        dValor = dValorAnt
    
    Else
    
        Set objNFiscal = New ClassNFiscal
            
        Call objNFiscal.Inicializa_Tributacao
        Call objNFiscal.objTributacaoNF.Coloca_Auto
        
        objNFiscal.iCondicaoPagto = 0
        
        objNFiscal.lCliente = lCliente
        objNFiscal.iFilialCli = iFilialCli
        
        If objNFiscal.lCliente = 0 Then
            objNFiscal.lCliente = CLIENTE_LOJA_TRIB
            objNFiscal.iFilialCli = FILIAL_MATRIZ
            iTipoNF = DOCINFO_TNFISFV
            sSigla = "TNFISFV"
        Else
            iTipoNF = DOCINFO_NFISFV
            sSigla = "NFISFV"
        End If
        
        objNFiscal.dtDataReferencia = gdtDataAtual
        objNFiscal.dtDataEmissao = gdtDataAtual
        objNFiscal.dtDataSaida = gdtDataAtual
        objNFiscal.dtDataEntrada = DATA_NULA
        objNFiscal.dtDataVencimento = DATA_NULA
        objNFiscal.iFilialEmpresa = iFilialEmpresa
        objNFiscal.iFilialEntrega = objNFiscal.iFilialCli
        
        objNFiscal.iTipoDocInfo = iTipoNF
        objNFiscal.iTipoNFiscal = iTipoNF
        objNFiscal.iStatus = 1
        objNFiscal.iRecibo = 0
        
        Set objItemNF = New ClassItemNF
        Call objItemNF.Inicializa_Tributacao
        Call objItemNF.objTributacaoItemNF.Coloca_Auto
        
        objItemNF.dCusto = 0
        objItemNF.dQuantidade = 1
        objItemNF.dPrecoUnitario = dPreco
        objItemNF.dValorTotal = dPreco
        objItemNF.sProduto = sProduto
        objItemNF.iItem = 1
    
        objNFiscal.ColItensNF.Add1 objItemNF
        
        Set objProduto = New ClassProduto
        objProduto.sCodigo = objItemNF.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
        
        objItemNF.objTributacaoItemNF.sProduto = objProduto.sCodigo
        objItemNF.objTributacaoItemNF.iExTIPI = objProduto.iExTIPI
        objItemNF.objTributacaoItemNF.sGenero = objProduto.sGenero
        objItemNF.objTributacaoItemNF.iProdutoEspecifico = objProduto.iProdutoEspecifico
        objItemNF.objTributacaoItemNF.sUMTrib = objProduto.sSiglaUMTrib
        objItemNF.objTributacaoItemNF.sIPICodProduto = objProduto.sIPICodigo
    
        objItemNF.objTributacaoItemNF.dQtdTrib = objItemNF.dQuantidade
        objItemNF.objTributacaoItemNF.dValorUnitTrib = objItemNF.dPrecoUnitario
        objItemNF.objTributacaoItemNF.dPrecoTotal = objItemNF.dValorTotal
        objItemNF.objTributacaoItemNF.sProdutoDescricao = objProduto.sDescricao
        objItemNF.objTributacaoItemNF.sISSQN = objProduto.sISSQN
        
        objNFiscal.dValorProdutos = objItemNF.dValorTotal
        
        Set objTribTab = New ClassTribTab
        
        lErro = objTribTab.TributacaoNF_Reset(objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'obtem dados do contrato
        Call objTribTab.TipoNFiscal_Definir(iTipoNF, sSigla)
        Call objTribTab.DataEmissao_Alterada(gdtDataAtual)
        Call objTribTab.Serie_Alterada(objNFiscal.sSerie)
        
        lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 1)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objTribTab.PreencherNF(objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Set objTribTab = Nothing
        
        objNFiscal.sNaturezaOp = objNFiscal.objTributacaoNF.sNaturezaOpInterna
        
        'acertar valor total e da parcela, considerando os tributos, retencóes,...
        objNFiscal.dValorTotal = dValorTotal + objNFiscal.objTributacaoNF.dICMSSubstValor + objNFiscal.objTributacaoNF.dIPIValor + objNFiscal.dValorFrete + objNFiscal.dValorSeguro + objNFiscal.dValorOutrasDespesas + IIf(objNFiscal.objTributacaoNF.iISSIncluso = 0, objNFiscal.objTributacaoNF.dISSValor, 0)
                     
        sUltProduto = sProduto
        lUltCliente = lCliente
        iUltFilialCli = iFilialCli
        iUltFilialEmp = iFilialEmpresa

        dValor = objItemNF.objTributacao.dICMSSubstValor
        dValorAnt = dValor
        
    End If

    Obter_ICMSSTValor = SUCESSO
    
    Exit Function
    
Erro_Obter_ICMSSTValor:

    Obter_ICMSSTValor = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154188)

    End Select
    
    Resume Next

    Exit Function
    
End Function

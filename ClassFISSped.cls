VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFISSped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    
Private Const STRING_CIAP_CODBEM = 60
Private Const STRING_CIAP_G125_TIPOMOV = 10

Private gsPerfil As String

Private giVersao As Integer
Private gsCodIBGE As String

Function Gera_Sped_Fiscal(ByVal sDiretorio As String, ByVal iFilialEmpresa As Integer, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal objTela As Object, ByVal iIncluiRegInv As Integer, ByVal iMotivoRegInv As Integer, ByVal dtDataInv As Date, ByVal iFiltroNatureza As Integer, ByVal iIncluiRCPE As Integer) As Long

Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_Gera_Sped_Fiscal

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 203288

    lErro = CF("Gera_Sped_Fiscal_Trans", sDiretorio, iFilialEmpresa, dtDataIni, dtDataFim, objTela, iIncluiRegInv, iMotivoRegInv, dtDataInv, iFiltroNatureza, iIncluiRCPE)
    If lErro <> AD_SQL_SUCESSO Then gError 210135

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 203369
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_SPEDFISCAL_SUCESSO", sDiretorio)
    
    Gera_Sped_Fiscal = SUCESSO
    
    Exit Function
    
Erro_Gera_Sped_Fiscal:

    Gera_Sped_Fiscal = gErr
    
    Select Case gErr
    
        Case 203288
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 203369
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 210135
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210136)
     
    End Select
     
    Call Transacao_Rollback

End Function

Function Gera_Sped_Fiscal_Trans(ByVal sDiretorio As String, ByVal iFilialEmpresa As Integer, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal objTela As Object, ByVal iIncluiRegInv As Integer, ByVal iMotivoRegInv As Integer, ByVal dtDataInv As Date, ByVal iFiltroNatureza As Integer, ByVal iIncluiRCPE As Integer) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim lNumIntSped As Long
Dim iCodVersao As Integer
Dim lTransacao As Long
Dim alComando(1 To 41) As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objSpedFiscal As ClassSpedFiscal
Dim objSpedRegistro As ClassSpedRegistro
Dim colSpedRegistro As New Collection
Dim tNFiscal As typeNFiscal
Dim tTipoDocInfo As typeTipoDocInfo
Dim iNFe As Integer
Dim sCodParticipante As String
Dim tTributacaoNF As typeTributacaoNF
Dim colMensagem As New Collection
Dim colC101 As New Collection
Dim lSeqPai As Long, lSeqPaiD2 As Long
Dim objSpedRegistroAux As ClassSpedRegistro
Dim objSpedRegistro0150 As New ClassSpedRegistro
Dim objSpedRegistro0175 As New ClassSpedRegistro
Dim objSpedRegistro0190 As New ClassSpedRegistro
Dim objSpedRegistro0200 As New ClassSpedRegistro
Dim objSpedRegistro0205 As New ClassSpedRegistro
Dim objSpedRegistro0210 As New ClassSpedRegistro
Dim objSpedRegistro0220 As New ClassSpedRegistro
Dim objSpedRegistro0300 As New ClassSpedRegistro
Dim objSpedRegistro0305 As New ClassSpedRegistro
Dim objSpedRegistro0450 As New ClassSpedRegistro
Dim objSpedRegistro0500 As New ClassSpedRegistro
Dim objSpedRegistro0600 As New ClassSpedRegistro
Dim objSpedRegistro0990 As New ClassSpedRegistro
Dim objSpedRegistroC001 As New ClassSpedRegistro
Dim objSpedRegistroC100 As New ClassSpedRegistro
Dim objSpedRegistroC101 As New ClassSpedRegistro
Dim objSpedRegistroC110 As New ClassSpedRegistro
Dim objSpedRegistroC111 As New ClassSpedRegistro
Dim objSpedRegistroC112 As New ClassSpedRegistro
Dim objSpedRegistroC113 As New ClassSpedRegistro
Dim objSpedRegistroC114 As New ClassSpedRegistro
Dim objSpedRegistroC115 As New ClassSpedRegistro
Dim objSpedRegistroC120 As New ClassSpedRegistro
Dim objSpedRegistroC130 As New ClassSpedRegistro
Dim objSpedRegistroC140 As New ClassSpedRegistro
Dim objSpedRegistroC141 As New ClassSpedRegistro
Dim objSpedRegistroC160 As New ClassSpedRegistro
Dim objSpedRegistroC170 As New ClassSpedRegistro
Dim objSpedRegistroC172 As New ClassSpedRegistro
Dim objSpedRegistroC190 As New ClassSpedRegistro
Dim objSpedRegistroC191 As New ClassSpedRegistro
Dim objSpedRegistroC195 As New ClassSpedRegistro

Dim objSpedRegistroC300 As New ClassSpedRegistro
Dim objSpedRegistroC320 As New ClassSpedRegistro
Dim objSpedRegistroC321 As New ClassSpedRegistro

Dim objSpedRegistroC400 As New ClassSpedRegistro
Dim objSpedRegistroC405 As New ClassSpedRegistro
Dim objSpedRegistroC420 As New ClassSpedRegistro
Dim objSpedRegistroC425 As New ClassSpedRegistro
Dim objSpedRegistroC460 As New ClassSpedRegistro
Dim objSpedRegistroC470 As New ClassSpedRegistro
Dim objSpedRegistroC490 As New ClassSpedRegistro
Dim objSpedRegistroC500 As New ClassSpedRegistro
Dim objSpedRegistroC590 As New ClassSpedRegistro
Dim objSpedRegistroC990 As New ClassSpedRegistro
Dim objSpedRegistroD001 As New ClassSpedRegistro
Dim objSpedRegistroD100 As New ClassSpedRegistro
Dim objSpedRegistroD190 As New ClassSpedRegistro
Dim objSpedRegistroD500 As New ClassSpedRegistro
Dim objSpedRegistroD590 As New ClassSpedRegistro
Dim objSpedRegistroD990 As New ClassSpedRegistro
Dim objSpedRegistroE001 As New ClassSpedRegistro
Dim objSpedRegistroE100 As New ClassSpedRegistro
Dim objSpedRegistroE110 As New ClassSpedRegistro
Dim objSpedRegistroE111 As New ClassSpedRegistro
Dim objSpedRegistroE116 As New ClassSpedRegistro
Dim objSpedRegistroE200 As New ClassSpedRegistro
Dim objSpedRegistroE210 As New ClassSpedRegistro
Dim objSpedRegistroE220 As New ClassSpedRegistro
Dim objSpedRegistroE250 As New ClassSpedRegistro
Dim objSpedRegistroE300 As New ClassSpedRegistro
Dim objSpedRegistroE310 As New ClassSpedRegistro
Dim objSpedRegistroE311 As New ClassSpedRegistro
Dim objSpedRegistroE312 As New ClassSpedRegistro
Dim objSpedRegistroE313 As New ClassSpedRegistro
Dim objSpedRegistroE316 As New ClassSpedRegistro
Dim objSpedRegistroE500 As New ClassSpedRegistro
Dim objSpedRegistroE510 As New ClassSpedRegistro
Dim objSpedRegistroE520 As New ClassSpedRegistro
Dim objSpedRegistroE530 As New ClassSpedRegistro
Dim objSpedRegistroE990 As New ClassSpedRegistro
Dim objSpedRegistroG001 As New ClassSpedRegistro
Dim objSpedRegistroG110 As New ClassSpedRegistro
Dim objSpedRegistroG125 As New ClassSpedRegistro
Dim objSpedRegistroG126 As New ClassSpedRegistro
Dim objSpedRegistroG130 As New ClassSpedRegistro
Dim objSpedRegistroG140 As New ClassSpedRegistro
Dim objSpedRegistroG990 As New ClassSpedRegistro
Dim objSpedRegistroH001 As New ClassSpedRegistro
Dim objSpedRegistroH005 As New ClassSpedRegistro
Dim objSpedRegistroH010 As New ClassSpedRegistro
Dim objSpedRegistroH990 As New ClassSpedRegistro
Dim objSpedRegistroK001 As New ClassSpedRegistro
Dim objSpedRegistroK100 As New ClassSpedRegistro
Dim objSpedRegistroK200 As New ClassSpedRegistro
Dim objSpedRegistroK220 As New ClassSpedRegistro
Dim objSpedRegistroK230 As New ClassSpedRegistro
Dim objSpedRegistroK235 As New ClassSpedRegistro
Dim objSpedRegistroK250 As New ClassSpedRegistro
Dim objSpedRegistroK255 As New ClassSpedRegistro
Dim objSpedRegistroK990 As New ClassSpedRegistro
Dim objSpedRegistro1001 As New ClassSpedRegistro
Dim objSpedRegistro1010 As New ClassSpedRegistro
Dim objSpedRegistro1100 As New ClassSpedRegistro
Dim objSpedRegistro1105 As New ClassSpedRegistro
Dim objSpedRegistro1400 As New ClassSpedRegistro
Dim objSpedRegistro1990 As New ClassSpedRegistro
Dim objSpedRegistro9001 As New ClassSpedRegistro
Dim objSpedRegistro9900 As New ClassSpedRegistro
Dim objSpedRegistro9990 As New ClassSpedRegistro
Dim objSpedRegistro9999 As New ClassSpedRegistro
Dim colUM As New Collection
Dim iCodigoECF As Integer
Dim iCaixa As Integer
Dim sModeloECF As String
Dim sNumSerieECF As String
Dim tArqMestre As typeLojaArqFisMestre
Dim lSeqECF As Long
Dim lSeqReducaoZ As Long
Dim colSpedC190 As New Collection
Dim colSpedC490 As New Collection
Dim dValDebICMS_objSpedC190 As Double
Dim dValCredICMS_objSpedC190 As Double
Dim dValDebICMS_objSpedC490 As Double
Dim dValCredICMS_objSpedC590 As Double
Dim dValCredICMS_objSpedD190 As Double
Dim dValCredICMS_objSpedD590 As Double
Dim colSpedE510 As New Collection
Dim dValDebIPI As Double
Dim dValCredIPI As Double
Dim dValorTotalEst As Double
Dim iAchouRedz As Integer
Dim dValorICMSRecolher As Double
Dim lTotal As Long
Dim lContador As Long
Dim iAno As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim lSeqE100 As Long
Dim lSeqE500 As Long
Dim colSpedC190UF As New Collection
Dim colSpedC590UF As New Collection
Dim colSpedD590UF As New Collection
Dim objSpedAux As New ClassFISSpedAux
Dim lNumNFInut As Long, lNumNFDeInut As Long, lNumNFAteInut As Long
Dim sSerieNFInut As String, dtDataNFInut As Date
Dim sNumDe As String, sNumAte As String
Dim lD2NumNFIni As Long, lD2NumNFFim As Long, sD2Serie As String, dtD2Data As Date, dD2Total As Double
Dim sD2Produto As String, sD2UM As String, dD2ProdQtde As Double, dD2ProdTotal As Double, dD2ProdDesc As Double
Dim dVendaBruta As Double

On Error GoTo Erro_Gera_Sped_Fiscal_Trans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203289
    Next

    If Not objTela Is Nothing Then
    objTela.dValorAtual = 0
    objTela.iCancelaBatch = 0
    End If

    objSpedRegistro0150.sTipo = "0150"
    objSpedRegistro0175.sTipo = "0175"
    objSpedRegistro0190.sTipo = "0190"
    objSpedRegistro0200.sTipo = "0200"
    objSpedRegistro0205.sTipo = "0205"
    objSpedRegistro0220.sTipo = "0220"
    objSpedRegistro0450.sTipo = "0450"


    'se for uma nota fiscal de entrada ou saida
    lErro = Comando_Executar(alComando(6), "SELECT Count(*) From NFiscal AS N, TiposDocInfo AS T WHERE TipoNFiscal = Codigo AND  ((Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) Or (Tipo = 1 Or Tipo =3) AND DataEntrada >= ? AND DataEntrada <= ? ) AND FilialEmpresa = ? ", lContador, dtDataIni, dtDataFim, dtDataIni, dtDataFim, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 204129

    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204130


    If Not objTela Is Nothing Then
    objTela.dValorTotal = lContador
    End If
    
    lContador = 0

    lErro = Comando_Executar(alComando(7), "SELECT Count(*)  FROM ECF AS E, ImpressoraECF AS I, ModeloECF AS M WHERE E.FilialEmpresa = ? AND E.ImpressoraECF = I.Codigo AND E.FilialEmpresa = I.FilialEmpresa AND I.CodModelo = M.Codigo", _
    lContador, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 204131

    lErro = Comando_BuscarPrimeiro(alComando(7))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204132


    If Not objTela Is Nothing Then
    objTela.dValorTotal = objTela.dValorTotal + lContador
    End If
    
    If Month(dtDataIni) = 2 Then

        lContador = 0

        iAno = Year(dtDataIni)

        lErro = Comando_Executar(alComando(3), "SELECT Count(*) FROM SldMesEst, Produtos WHERE Ano = ? AND FilialEmpresa = ? ", lContador, iAno, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 204127
    
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204128

    If Not objTela Is Nothing Then
        objTela.dValorTotal = objTela.dValorTotal + lContador
    End If

    End If

    lErro = CF("Config_ObterNumInt", "FISConfig", "NUM_PROX_SPEDFISCAL", lNumIntSped)
    If lErro <> SUCESSO Then gError 203290
    
    lErro = Comando_Executar(alComando(2), "SELECT Codigo From SpedFisVersaoLeiaute WHERE DataInicio <= ? ORDER BY DataInicio DESC ", iCodVersao, dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 203291
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203292
    
    'versao nao cadastrada
    If lErro <> AD_SQL_SUCESSO Then gError 203293
    
    giVersao = iCodVersao
    
    objFilialEmpresa.iCodFilial = iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 203294

    If lErro <> SUCESSO Then gError 203295
    
    gsPerfil = objFilialEmpresa.sSpedFiscalPerfil
    
    '******** registro 0000 *****************
    lErro = SpedFiscal_0000(iFilialEmpresa, lNumIntSped, objSpedFiscal, colSpedRegistro, iCodVersao, dtDataIni, dtDataFim, objFilialEmpresa)
    If lErro <> SUCESSO Then gError 203296
    
    '******** registro 0001 *****************
    lErro = SpedFiscal_0001(objSpedFiscal, colSpedRegistro)
    If lErro <> SUCESSO Then gError 203297

    If objFilialEmpresa.iTipoTribAtividade = 2 And giVersao >= 14 Then
        lErro = SpedFiscal_0002(objSpedFiscal, colSpedRegistro)
        If lErro <> SUCESSO Then gError 203297
    End If
    
    '******** registro 0005 *****************
    lErro = SpedFiscal_0005(objFilialEmpresa, objSpedFiscal, colSpedRegistro)
    If lErro <> SUCESSO Then gError 203298
    
    '******** registro 0015 *****************
    lErro = SpedFiscal_0015(objFilialEmpresa, objSpedFiscal, colSpedRegistro)
    If lErro <> SUCESSO Then gError 203299
    
    '******** registro 0100 *****************
    lErro = SpedFiscal_0100(objFilialEmpresa, objSpedFiscal, colSpedRegistro)
    If lErro <> SUCESSO Then gError 203300
    
    With objSpedAux
    
        Set .objSpedFiscal = objSpedFiscal
        
        .dtDataIni = dtDataIni
        .dtDataFim = dtDataFim
        .iFilialEmpresa = iFilialEmpresa
        .iIncluiRCPE = iIncluiRCPE
        
        Set .objSpedRegistroG001 = objSpedRegistroG001
        Set .objSpedRegistroG110 = objSpedRegistroG110
        Set .objSpedRegistroG125 = objSpedRegistroG125
        Set .objSpedRegistroG126 = objSpedRegistroG126
        Set .objSpedRegistroG130 = objSpedRegistroG130
        Set .objSpedRegistroG140 = objSpedRegistroG140
        Set .objSpedRegistroG990 = objSpedRegistroG990
        
        Set .objSpedRegistro0150 = objSpedRegistro0150
        Set .objSpedRegistro0175 = objSpedRegistro0175
        Set .objSpedRegistro0190 = objSpedRegistro0190
        Set .objSpedRegistro0200 = objSpedRegistro0200
        Set .objSpedRegistro0205 = objSpedRegistro0205
        Set .objSpedRegistro0210 = objSpedRegistro0210
        Set .objSpedRegistro0220 = objSpedRegistro0220
        Set .objSpedRegistro0300 = objSpedRegistro0300
        Set .objSpedRegistro0305 = objSpedRegistro0305
        Set .objSpedRegistro0500 = objSpedRegistro0500
        Set .objSpedRegistro0600 = objSpedRegistro0600
        
        Set .objSpedRegistro1001 = objSpedRegistro1001
        Set .objSpedRegistro1010 = objSpedRegistro1010
        Set .objSpedRegistro1100 = objSpedRegistro1100
        Set .objSpedRegistro1105 = objSpedRegistro1105
        Set .objSpedRegistro1400 = objSpedRegistro1400
        
        Set .objSpedRegistroK001 = objSpedRegistroK001
        Set .objSpedRegistroK100 = objSpedRegistroK100
        Set .objSpedRegistroK200 = objSpedRegistroK200
        Set .objSpedRegistroK220 = objSpedRegistroK220
        Set .objSpedRegistroK230 = objSpedRegistroK230
        Set .objSpedRegistroK235 = objSpedRegistroK235
        Set .objSpedRegistroK250 = objSpedRegistroK250
        Set .objSpedRegistroK255 = objSpedRegistroK255
        Set .objSpedRegistroK990 = objSpedRegistroK990
        
        'E Período de Apuração do ICMS Diferencial de Alíquota – UF Origem/Destino EC 87/15 E300 2 1:N
        'E Apuração do ICMS Diferencial de Alíquota – UF Origem/Destino EC 87/15 E310 3 1:1
        'E Ajuste/Benefício/Incentivo da Apuração do ICMS Diferencial de Alíquota – UF Origem/Destino EC 87/15 E311 4 1:N
        'E Informações Adicionais dos Ajustes da Apuração do ICMS Diferencial de Alíquota – UF Origem/Destino EC 87/15 E312 5 1:N
        'E Informações Adicionais da Apuração do ICMS Diferencial de Alíquota – UF Origem/Destino EC 87/15 Identificação dos Documentos Fiscais E313 5 1:N
        'E Obrigações do ICMS recolhido ou a recolher – Diferencial de Alíquota – UF Origem/Destino EC 87/15 E316 4 1:N
        Set .objSpedRegistroE300 = objSpedRegistroE300
        Set .objSpedRegistroE310 = objSpedRegistroE310
        Set .objSpedRegistroE311 = objSpedRegistroE311
        Set .objSpedRegistroE312 = objSpedRegistroE312
        Set .objSpedRegistroE313 = objSpedRegistroE313
        Set .objSpedRegistroE316 = objSpedRegistroE316
        
        Set .objSpedRegistroC101 = objSpedRegistroC101
        Set .objSpedRegistroC191 = objSpedRegistroC191
        
        Set .colUM = colUM
    
    End With
    
    '******** registro C001 *****************
    lErro = SpedFiscal_C001(objSpedFiscal, objSpedRegistroC001, dtDataIni, dtDataFim, iFilialEmpresa)
    If lErro <> SUCESSO Then gError 204141
    
    '***** pesquisa as notas fiscais ***************
    
    tNFiscal.sSerie = String(STRING_SERIE, 0)
    tNFiscal.sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
    tNFiscal.sPlaca = String(STRING_NFISCAL_PLACA, 0)
    tNFiscal.sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
    tNFiscal.sChvNFe = String(STRING_NFE_CHNFE, 0)
    tTipoDocInfo.sNomeTelaNFiscal = String(STRING_NOME, 0)
    
    
    'se for uma nota fiscal  saida
    lErro = Comando_Executar(alComando(4), "SELECT TipoNFiscal, NumIntDoc, ModeloArqICMS, ChvNFe, Tipo, Cliente, FilialCli, Fornecedor, FilialForn, Status, Complementar, Serie, NumNotaFiscal, DataEmissao, DataEntrada, DataSaida, ValorTotal, ClasseDocCPR, NumIntDocCPR, ValorDesconto, " & _
        "ValorProdutos, ValorFrete, FreteRespons, ValorSeguro, ValorOutrasDespesas, MensagemNota, FilialEntrega, CodTransportadora, Placa, PlacaUF, VolumeQuant, PesoLiq, PesoBruto, T.TipoMovtoEstoque, NomeTelaNFiscal, T.Faturavel FROM NFiscal AS N, TiposDocInfo AS T WHERE ModeloArqICMS <> 0 AND TipoNFiscal = Codigo AND ((Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) Or ((Tipo = 1 Or Tipo = 3) AND DataEntrada >= ? AND DataEntrada <= ?)) AND N.FilialEmpresa = ? ORDER BY CASE WHEN Tipo=2 THEN DataEmissao ELSE DataEntrada END, Serie, NumNotaFiscal", _
        tNFiscal.iTipoNFiscal, tNFiscal.lNumIntDoc, tTipoDocInfo.iModeloArqICMS, tNFiscal.sChvNFe, tTipoDocInfo.iTipo, tNFiscal.lCliente, tNFiscal.iFilialCli, tNFiscal.lFornecedor, tNFiscal.iFilialForn, tNFiscal.iStatus, tTipoDocInfo.iComplementar, tNFiscal.sSerie, tNFiscal.lNumNotaFiscal, tNFiscal.dtDataEmissao, tNFiscal.dtDataEntrada, tNFiscal.dtDataSaida, tNFiscal.dValorTotal, _
        tNFiscal.iClasseDocCPR, tNFiscal.lNumIntDocCPR, tNFiscal.dValorDesconto, tNFiscal.dValorProdutos, tNFiscal.dValorFrete, tNFiscal.iFreteRespons, tNFiscal.dValorSeguro, tNFiscal.dValorOutrasDespesas, tNFiscal.sMensagemNota, tNFiscal.iFilialEntrega, tNFiscal.iCodTransportadora, tNFiscal.sPlaca, tNFiscal.sPlacaUF, _
        tNFiscal.lVolumeQuant, tNFiscal.dPesoLiq, tNFiscal.dPesoBruto, tTipoDocInfo.iTipoMovtoEstoque, tTipoDocInfo.sNomeTelaNFiscal, tTipoDocInfo.iFaturavel, dtDataIni, dtDataFim, dtDataIni, dtDataFim, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203301
    
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203302
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = SpedFiscal_Trata_ItemNF(objTela, objFilialEmpresa, dtDataIni, dtDataFim, objSpedFiscal, tNFiscal, objSpedRegistro0150, objSpedRegistro0175, objSpedRegistro0190, objSpedRegistro0200, objSpedRegistro0205, objSpedRegistro0220, tTipoDocInfo, alComando, objSpedRegistroC100, objSpedRegistro0450, objSpedRegistroC110, _
        objSpedRegistroC111, objSpedRegistroC112, objSpedRegistroC113, objSpedRegistroC114, objSpedRegistroC115, objSpedRegistroC120, objSpedRegistroC130, objSpedRegistroC140, objSpedRegistroC141, objSpedRegistroC160, _
        objSpedRegistroC170, objSpedRegistroC172, objSpedRegistroC190, colSpedE510, colSpedC190, dValDebICMS_objSpedC190, dValCredICMS_objSpedC190, colMensagem, colUM, colSpedC190UF, objSpedAux)
        If lErro <> SUCESSO Then gError 204914
        
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203329
            
    Loop
    
    'NFs Inutilizadas
    sSerieNFInut = String(STRING_SERIE, 0)
    sNumDe = String(STRING_MAXIMO, 0)
    sNumAte = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(35), "SELECT Data, Serie, nNFIni, nNFFim  FROM NFeFedRetInutNFe WHERE Data BETWEEN ? AND ? AND tpAmb = ? AND cStat = ? AND FilialEmpresa = ? ORDER BY Data, nNFIni ", dtDataNFInut, sSerieNFInut, sNumDe, sNumAte, dtDataIni, dtDataFim, 0, "102", iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203301
    
    lErro = Comando_BuscarPrimeiro(alComando(35))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203302
    
    Do While lErro = AD_SQL_SUCESSO
    
        lNumNFDeInut = StrParaLong(sNumDe)
        lNumNFAteInut = StrParaLong(sNumAte)
    
        For lNumNFInut = lNumNFDeInut To lNumNFAteInut
        
            objSpedFiscal.sTipo = "C100"
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            
            objSpedRegistroC100.sTipo = objSpedFiscal.sTipo
            objSpedRegistroC100.iCount = objSpedRegistroC100.iCount + 1
        
            objSpedFiscal.sRegistro = "|C100|1|0||55|05|" & sSerieNFInut & "|" & lNumNFInut & "||||||||||||||||||||||"
    
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
            If lErro <> AD_SQL_SUCESSO Then gError 203315
        
        Next
    
        lErro = Comando_BuscarProximo(alComando(35))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203329
            
    Loop
    'NFs Inutilizadas
   
    If gsPerfil = "B" Then
    
        '********************** registro C300 *******************************
        sD2Serie = String(STRING_SERIE, 0)
        lErro = Comando_Executar(alComando(37), "SELECT MIN(NumNotaFiscal), MAX(NumNotaFiscal), SUM(ValorTotal), Serie, DataEmissaoNF FROM CupomFiscal WHERE NFeChaveAcesso = '' AND Status <> 7 AND FilialEmpresa = ? AND DataEmissaoNF >= ? AND DataEmissaoNF <= ? AND Tipo = 1 AND NumNotaFiscal <> 0 GROUP BY Serie, DataEmissaoNF", _
            lD2NumNFIni, lD2NumNFFim, dD2Total, sD2Serie, dtD2Data, objSpedFiscal.iFilialEmpresa, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203330
    
        lErro = Comando_BuscarProximo(alComando(37))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203331
        
        Do While lErro = AD_SQL_SUCESSO
        
            objSpedFiscal.sTipo = "C300"
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            
            objSpedFiscal.sRegistro = "|C300|02|D|" & sD2Serie & "|" & CStr(lD2NumNFIni) & "|" & CStr(lD2NumNFFim) & "|" & Format(dtD2Data, "ddmmyyyy") & "|" & Format(dD2Total, "Fixed") & "||||"
            
            objSpedRegistroC300.sTipo = objSpedFiscal.sTipo
            objSpedRegistroC300.iCount = objSpedRegistroC300.iCount + 1
                    
            lErro = Comando_Executar(alComando(38), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
            If lErro <> AD_SQL_SUCESSO Then gError 203315
            
            objSpedFiscal.sTipo = "C320"
            objSpedFiscal.lSeqPai = objSpedFiscal.lSeq
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            
            objSpedFiscal.sRegistro = "|C320|090|5929||" & Format(dD2Total, "Fixed") & "|0|0|0||"
            
            objSpedRegistroC320.sTipo = objSpedFiscal.sTipo
            objSpedRegistroC320.iCount = objSpedRegistroC320.iCount + 1
                    
            lErro = Comando_Executar(alComando(40), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, objSpedFiscal.lSeqPai)
            If lErro <> AD_SQL_SUCESSO Then gError 203315
            
            lSeqPaiD2 = objSpedFiscal.lSeq
            
            sD2Produto = String(STRING_PRODUTO, 0)
            sD2UM = String(STRING_UM_SIGLA, 0)
            lErro = Comando_Executar(alComando(39), "SELECT I.Produto, MAX(I.UnidadeMed), SUM(I.Quantidade), SUM(I.Quantidade*I.PrecoUnitario), SUM(I.ValorDesconto) FROM CupomFiscal C, ItensCupomFiscal I WHERE C.NFeChaveAcesso = '' AND I.NumIntCupom = C.NumIntDoc AND I.FilialEmpresa = C.FilialEmpresa AND I.Status <> 7 AND C.Status <> 7 AND C.NumNotaFiscal <> 0 AND C.DataEmissaoNF = ? AND C.Serie = ? AND C.FilialEmpresa = ? AND C.Tipo = 1 GROUP BY I.Produto", _
                sD2Produto, sD2UM, dD2ProdQtde, dD2ProdTotal, dD2ProdDesc, dtD2Data, sD2Serie, objSpedFiscal.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 203330
        
            lErro = Comando_BuscarProximo(alComando(39))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203331
            
            Do While lErro = AD_SQL_SUCESSO
                
                objSpedFiscal.sTipo = "C321"
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedFiscal.lSeqPai = lSeqPaiD2
                
                objSpedFiscal.sRegistro = "|C321|" & sD2Produto & "|" & Format(dD2ProdQtde, "##############0.0##") & "|" & sD2UM & "|" & Format(dD2ProdTotal, "Fixed") & "|" & Format(dD2ProdDesc, "Fixed") & "|||||"
                
                objSpedRegistroC321.sTipo = objSpedFiscal.sTipo
                objSpedRegistroC321.iCount = objSpedRegistroC321.iCount + 1
                
                lErro = Comando_Executar(alComando(41), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, objSpedFiscal.lSeqPai)
                If lErro <> AD_SQL_SUCESSO Then gError 203315
            
                lErro = Comando_BuscarProximo(alComando(39))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203331
            
            Loop
                
            lErro = Comando_BuscarProximo(alComando(37))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203331
        
        Loop
            
            
        
        '********************** fim dos registros C300 *******************************
    
    End If
    
    '********************** registro C400 *******************************
    
    sModeloECF = String(STRING_MODELOECF_NOME, 0)
    sNumSerieECF = String(STRING_IMPRESSORAECF_NUMSERIE, 0)
    
    lErro = Comando_Executar(alComando(29), "SELECT E.Codigo, E.Caixa, M.Nome, I.Numserie  FROM ECF AS E, ImpressoraECF AS I, ModeloECF AS M WHERE E.FilialEmpresa = ? AND E.ImpressoraECF = I.Codigo AND E.FilialEmpresa = I.FilialEmpresa AND I.CodModelo = M.Codigo", _
    iCodigoECF, iCaixa, sModeloECF, sNumSerieECF, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203330

    lErro = Comando_BuscarPrimeiro(alComando(29))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203331
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_Executar(alComando(30), "SELECT Data, CRO, CRZ, COOFim, GrandeTotal, VendaBruta FROM LojaArqFisMestre WHERE CodECF = ? And FilialEmpresa = ? And Data >= ? AND Data <= ?", _
        tArqMestre.dtData, tArqMestre.lCRO, tArqMestre.lCRZ, tArqMestre.lCOOFim, tArqMestre.dGrandeTotal, tArqMestre.dVendaBruta, iCodigoECF, iFilialEmpresa, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203332
    
        lErro = Comando_BuscarPrimeiro(alComando(30))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203333
    
        iAchouRedz = 0
    
        Do While lErro = AD_SQL_SUCESSO
        
            If iAchouRedz = 0 Then
        
                iAchouRedz = 1
        
                objSpedFiscal.sTipo = "C400"
                objSpedRegistroC400.sTipo = objSpedFiscal.sTipo
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistroC400.iCount = objSpedRegistroC400.iCount + 1
        
                objSpedFiscal.sRegistro = "|C400|2D|" & sModeloECF & "|" & sNumSerieECF & "|" & iCaixa & "|"
                
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203335
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
                If lErro <> AD_SQL_SUCESSO Then gError 203336
                
                lSeqECF = objSpedFiscal.lSeq
                
            End If
    
            '********************** registro C405 *******************************
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            lSeqReducaoZ = objSpedFiscal.lSeq

            '********************** registro C420 *******************************
            lErro = SpedFiscal_C420(objSpedFiscal, objSpedRegistroC420, lSeqReducaoZ, tArqMestre, iFilialEmpresa, iCodigoECF, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, objSpedRegistroC425, dVendaBruta, objSpedAux)
            If lErro <> SUCESSO Then gError 203338
    
            lErro = SpedFiscal_C405(objSpedFiscal, objSpedRegistroC405, lSeqECF, tArqMestre, lSeqReducaoZ, dVendaBruta)
            If lErro <> SUCESSO Then gError 203337
    
            Set colSpedC490 = New Collection
    
            '******************* registro C460 *******************
    
            lErro = SpedFiscal_C460(iCodigoECF, iFilialEmpresa, tArqMestre, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, objSpedRegistroC460, objSpedRegistroC470, lSeqReducaoZ, colSpedC490, dValDebICMS_objSpedC490, objSpedAux)
            If lErro <> SUCESSO Then gError 203339
    
            '******************* registro C490 *******************
    
            lErro = SpedFiscal_C490(objSpedFiscal, objSpedRegistroC490, lSeqReducaoZ, colSpedC490)
            If lErro <> SUCESSO Then gError 203340
    
            lErro = Comando_BuscarProximo(alComando(30))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203341
    
        Loop
    
        If Not objTela Is Nothing Then
    
        lErro = objTela.Processa_Item
        If lErro <> SUCESSO Then gError 204134
    
        End If
        lErro = Comando_BuscarProximo(alComando(29))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203342
    
    Loop

    '******************* registro C500 *******************
    
    lErro = SpedFiscal_C500(objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, objSpedRegistroC500, objSpedRegistroC590, dValCredICMS_objSpedC590, dtDataIni, dtDataFim, colSpedC590UF, objFilialEmpresa)
    If lErro <> SUCESSO Then gError 203343
    
    '******************* registro C990 *******************
    
    objSpedFiscal.sTipo = "C990"
    objSpedRegistroC990.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroC990.iCount = objSpedRegistroC990.iCount + 1

    lTotal = objSpedRegistroC001.iCount + objSpedRegistroC100.iCount + objSpedRegistroC110.iCount + objSpedRegistroC111.iCount + objSpedRegistroC112.iCount + objSpedRegistroC113.iCount
    lTotal = lTotal + objSpedRegistroC114.iCount + objSpedRegistroC115.iCount + objSpedRegistroC120.iCount + objSpedRegistroC130.iCount + objSpedRegistroC101.iCount
    lTotal = lTotal + objSpedRegistroC140.iCount + objSpedRegistroC141.iCount + objSpedRegistroC160.iCount + objSpedRegistroC170.iCount
    lTotal = lTotal + objSpedRegistroC300.iCount + objSpedRegistroC320.iCount + objSpedRegistroC321.iCount
    lTotal = lTotal + objSpedRegistroC172.iCount + objSpedRegistroC190.iCount + objSpedRegistroC195.iCount + objSpedRegistroC400.iCount
    lTotal = lTotal + objSpedRegistroC405.iCount + objSpedRegistroC420.iCount + objSpedRegistroC425.iCount + objSpedRegistroC460.iCount
    lTotal = lTotal + objSpedRegistroC470.iCount + objSpedRegistroC490.iCount + objSpedRegistroC500.iCount + objSpedRegistroC590.iCount + objSpedRegistroC990.iCount
    
    If giVersao > 12 Then lTotal = lTotal + objSpedRegistroC191.iCount
    
    objSpedFiscal.sRegistro = "|C990|" & lTotal & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203344
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203345
    
    
    '******************* registro D001 *******************
    lErro = SpedFiscal_D001(objSpedFiscal, objSpedRegistroD001, dtDataIni, dtDataFim)
    If lErro <> SUCESSO Then gError 203346
    
    '*************** registro D100 **********************
    lErro = SpedFiscal_D100(objSpedFiscal, objSpedRegistroD100, objSpedRegistroD190, objSpedRegistro0150, objSpedRegistro0175, objSpedRegistro0450, dtDataIni, dtDataFim, colMensagem, dValCredICMS_objSpedD190)
    If lErro <> SUCESSO Then gError 203347
    
    '*************** registro D500 **********************
    lErro = SpedFiscal_D500(objSpedFiscal, objSpedRegistroD500, objSpedRegistroD590, objSpedRegistro0150, objSpedRegistro0175, objSpedRegistro0450, dtDataIni, dtDataFim, colMensagem, dValCredICMS_objSpedD590, colSpedD590UF, objFilialEmpresa)
    If lErro <> SUCESSO Then gError 203348
    
    '******************* registro D990 *******************
    lErro = SpedFiscal_D990(objSpedFiscal, objSpedRegistroD990, objSpedRegistroD001, objSpedRegistroD100, objSpedRegistroD190, objSpedRegistroD500, objSpedRegistroD590)
    If lErro <> SUCESSO Then gError 203349
    
    '*************** registro E001 **********************
    lErro = SpedFiscal_E001(objSpedFiscal, objSpedRegistroE001)
    If lErro <> SUCESSO Then gError 203350
    
    '*************** registro E100 **********************
    lErro = SpedFiscal_E100(objSpedFiscal, objSpedRegistroE100, dtDataIni, dtDataFim, lSeqE100)
    If lErro <> SUCESSO Then gError 203351
    
    '******************* registro E110 *******************
    lErro = SpedFiscal_E110(objSpedFiscal, objSpedRegistroE110, objSpedRegistroE111, objFilialEmpresa, dtDataIni, dtDataFim, dValDebICMS_objSpedC190, dValDebICMS_objSpedC490, dValCredICMS_objSpedC190, dValCredICMS_objSpedC590, dValCredICMS_objSpedD190, dValCredICMS_objSpedD590, dValorICMSRecolher, lSeqPai, lSeqE100)
    If lErro <> SUCESSO Then gError 203352

    '*************** registro E116 **********************
    lErro = SpedFiscal_E116(objSpedFiscal, objSpedRegistroE116, objFilialEmpresa, dtDataIni, dtDataFim, lSeqPai, dValorICMSRecolher)
    If lErro <> SUCESSO Then gError 203353
    
    '*************** registro E200 **********************
    lErro = SpedFiscal_E200(objFilialEmpresa, objSpedFiscal, objSpedRegistroE200, objSpedRegistroE210, objSpedRegistroE220, objSpedRegistroE250, colSpedC190UF, colSpedC590UF, colSpedD590UF, dtDataIni, dtDataFim)
    If lErro <> SUCESSO Then gError 204953
    
    lErro = SpedFiscal_E300(objSpedAux)
    If lErro <> SUCESSO Then gError 203358
    
    'se é contribuinte do IPI registro E500 e filhos, caso contrario nao
    If objFilialEmpresa.iContribuinteIPI <> NAO_CONTRIBUINTE_IPI Then
    
        '*************** registro E500 **********************
        lErro = SpedFiscal_E500(objSpedFiscal, objSpedRegistroE500, dtDataIni, dtDataFim, lSeqE500)
        If lErro <> SUCESSO Then gError 203354
        
        '*************** registro E510 **********************
        lErro = SpedFiscal_E510(objSpedFiscal, objSpedRegistroE510, colSpedE510, dValDebIPI, dValCredIPI, lSeqE500)
        If lErro <> SUCESSO Then gError 203355
        
        '*************** registro E520 **********************
        lErro = SpedFiscal_E520(objSpedFiscal, objSpedRegistroE520, objSpedRegistroE530, objFilialEmpresa, dtDataIni, dtDataFim, dValDebIPI, dValCredIPI, lSeqE500)
        If lErro <> SUCESSO Then gError 203356
    
    End If
    
    '*************** registro E990 **********************
    lErro = SpedFiscal_E990(objSpedFiscal, objSpedRegistroE990, objSpedRegistroE001, objSpedRegistroE100, objSpedRegistroE110, objSpedRegistroE111, objSpedRegistroE116, objSpedRegistroE200, objSpedRegistroE210, objSpedRegistroE220, objSpedRegistroE250, objSpedRegistroE500, objSpedRegistroE510, objSpedRegistroE520, objSpedRegistroE530, objSpedAux)
    If lErro <> SUCESSO Then gError 203357
     
    '*************** registro G001 **********************
    With objSpedAux
    
        Set .objSpedFiscal = objSpedFiscal
        
        .dtDataIni = dtDataIni
        .dtDataFim = dtDataFim
        
        Set .objSpedRegistroG001 = objSpedRegistroG001
        Set .objSpedRegistroG110 = objSpedRegistroG110
        Set .objSpedRegistroG125 = objSpedRegistroG125
        Set .objSpedRegistroG126 = objSpedRegistroG126
        Set .objSpedRegistroG130 = objSpedRegistroG130
        Set .objSpedRegistroG140 = objSpedRegistroG140
        Set .objSpedRegistroG990 = objSpedRegistroG990
        
        Set .objSpedRegistro0150 = objSpedRegistro0150
        Set .objSpedRegistro0175 = objSpedRegistro0175
        Set .objSpedRegistro0190 = objSpedRegistro0190
        Set .objSpedRegistro0200 = objSpedRegistro0200
        Set .objSpedRegistro0205 = objSpedRegistro0205
        Set .objSpedRegistro0220 = objSpedRegistro0220
        Set .objSpedRegistro0300 = objSpedRegistro0300
        Set .objSpedRegistro0305 = objSpedRegistro0305
        Set .objSpedRegistro0500 = objSpedRegistro0500
        Set .objSpedRegistro0600 = objSpedRegistro0600
        
        Set .objSpedRegistro1001 = objSpedRegistro1001
        Set .objSpedRegistro1010 = objSpedRegistro1010
        Set .objSpedRegistro1100 = objSpedRegistro1100
        Set .objSpedRegistro1105 = objSpedRegistro1105
        Set .objSpedRegistro1400 = objSpedRegistro1400
        
        .iFilialEmpresa = iFilialEmpresa
        
        Set .colUM = colUM
    
    End With
    
    lErro = SpedFiscal_G001(objSpedAux)
    If lErro <> SUCESSO Then gError 203358
        
    '*************** registro G990 **********************
    lErro = SpedFiscal_G990(objSpedAux)
    If lErro <> SUCESSO Then gError 203361
    
    'Set objSpedAux = Nothing
    
    '*************** registro H001 **********************
    lErro = SpedFiscal_H001(objSpedFiscal, objSpedRegistroH001, dtDataIni, iIncluiRegInv)
    If lErro <> SUCESSO Then gError 203358
    
    'If Month(dtDataIni) = 2 Then
    If iIncluiRegInv = MARCADO Then
    
        '*************** registro H010 **********************
        lErro = SpedFiscal_H010(iFilialEmpresa, dtDataIni, dtDataFim, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, objSpedRegistro0150, objSpedRegistro0175, objSpedRegistroH010, objSpedRegistro0190, objSpedRegistro0220, colUM, dValorTotalEst, objTela, dtDataInv, iFiltroNatureza, objSpedAux)
        If lErro <> SUCESSO Then gError 203359
    
        '*************** registro H005 **********************
    
        lErro = SpedFiscal_H005(objSpedFiscal, objSpedRegistroH005, dtDataIni, dValorTotalEst, iMotivoRegInv, dtDataInv)
        If lErro <> SUCESSO Then gError 203360
    
    End If
    
    '*************** registro H990 **********************
    lErro = SpedFiscal_H990(objSpedFiscal, objSpedRegistroH990, objSpedRegistroH001, objSpedRegistroH005, objSpedRegistroH010)
    If lErro <> SUCESSO Then gError 203361
    
'    If dtDataIni >= #1/1/2016# Then
'
'        '*************** registro K001 **********************
'        lErro = SpedFiscal_K001(objSpedFiscal, objSpedRegistroK001)
'        If lErro <> SUCESSO Then gError 203358
'
'        '*************** registro K990 **********************
'        lErro = SpedFiscal_K990(objSpedFiscal, objSpedRegistroK990, objSpedRegistroK001)
'        If lErro <> SUCESSO Then gError 203361
'
'    End If

    '*******************BLOCO K****************
    If Year(objSpedAux.dtDataIni) >= 2016 Then
    
        '*************** registro K001 **********************
        lErro = SpedFiscal_K001(objSpedAux)
        If lErro <> SUCESSO Then gError 210771
        
        If iIncluiRCPE = MARCADO Then
    
            '*************** registro K100 **********************
            lErro = SpedFiscal_K100(objSpedAux)
            If lErro <> SUCESSO Then gError 210771
            
            '*************** registro K200 **********************
            lErro = SpedFiscal_K200(objSpedAux)
            If lErro <> SUCESSO Then gError 210771
            
            '*************** registro K220 **********************
            'lErro = SpedFiscal_K220(objSpedAux)
            'If lErro <> SUCESSO Then gError 210771
            
            '*************** registro K230 **********************
            '*************** registro K235 **********************
            lErro = SpedFiscal_K230(objSpedAux)
            If lErro <> SUCESSO Then gError 210771
            
            '*************** registro K250 **********************
            '*************** registro K255 **********************
            lErro = SpedFiscal_K250(objSpedAux)
            If lErro <> SUCESSO Then gError 210771
            
        End If
               
        '*************** registro K990 **********************
        lErro = SpedFiscal_K990(objSpedAux)
        If lErro <> SUCESSO Then gError 210771
        
    End If
    '*******************BLOCO K****************
    
    '*******************BLOCO B****************
    
    'Reaproveita o código do EFD Contribuições
    Dim objEFD As New ClassEFDPisCofinsSel
    objEFD.dtDataIni = objSpedAux.dtDataIni
    objEFD.dtDataFim = objSpedAux.dtDataFim
    objEFD.iFilialEmpresa = objSpedAux.iFilialEmpresa
    
    If giVersao > 12 Then
        
        lErro = SpedFiscal_Obtem_DadosInicias(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Processa_Bloco_B(objEFD, objSpedAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
           
    End If
    '*******************BLOCO B****************
    
    '*************** registro 1001 **********************
    lErro = SpedFiscal_1001(objSpedFiscal, objSpedRegistro1001)
    If lErro <> SUCESSO Then gError 203362
    
    '*************** registro 1010 **********************
    lErro = SpedFiscal_1010(objSpedFiscal, iFilialEmpresa, dtDataIni, dtDataFim, objSpedRegistro1010)
    If lErro <> SUCESSO Then gError 210771
    
    '*************** registro 1100 **********************
    '*************** registro 1105 **********************
    lErro = SpedFiscal_1100(objSpedAux)
    If lErro <> SUCESSO Then gError 210771
    
    '*************** registro 1400 **********************
    lErro = SpedFiscal_1400(objSpedAux)
    If lErro <> SUCESSO Then gError 210771
    
    '*************** registro 1990 **********************
    lErro = SpedFiscal_1990(objSpedFiscal, objSpedRegistro1990, objSpedAux)
    If lErro <> SUCESSO Then gError 203363
    
    Set objSpedAux = Nothing
    
    '*************** registro 9001 **********************
    lErro = SpedFiscal_9001(objSpedFiscal, objSpedRegistro9001)
    If lErro <> SUCESSO Then gError 203364
    
    '******************* registro 0990 *******************
    
    objSpedFiscal.sTipo = "0990"
    objSpedRegistro0990.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistro0990.iCount = objSpedRegistro0990.iCount + 1
    
    lTotal = 0
    
    For Each objSpedRegistro In colSpedRegistro
        lTotal = lTotal + objSpedRegistro.iCount
    Next

    lTotal = lTotal + objSpedRegistro0150.iCount + objSpedRegistro0175.iCount + objSpedRegistro0190.iCount
    lTotal = lTotal + objSpedRegistro0200.iCount + objSpedRegistro0205.iCount + objSpedRegistro0220.iCount
    lTotal = lTotal + objSpedRegistro0300.iCount + objSpedRegistro0305.iCount + objSpedRegistro0450.iCount + objSpedRegistro0500.iCount + objSpedRegistro0600.iCount + objSpedRegistro0210.iCount

    lTotal = lTotal + 1 'que é o proprio 0990
    
    objSpedFiscal.sRegistro = "|0990|" & lTotal & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204142
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204143
    
    
    If objSpedRegistro0150.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0150
    End If
    
    If objSpedRegistro0175.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0175
    End If
    
    If objSpedRegistro0190.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0190
    End If
    
    If objSpedRegistro0200.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0200
    End If
    
    If objSpedRegistro0205.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0205
    End If
    
    If objSpedRegistro0210.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0210
    End If
    
    If objSpedRegistro0220.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0220
    End If
    
    If objSpedRegistro0300.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0300
    End If
    
    If objSpedRegistro0305.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0305
    End If
    
    If objSpedRegistro0450.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0450
    End If
    
    If objSpedRegistro0500.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0500
    End If
    
    If objSpedRegistro0600.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro0600
    End If
    
    colSpedRegistro.Add objSpedRegistro0990
    
    If objSpedRegistroC001.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC001
    End If
    
    If objSpedRegistroC100.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC100
    End If
    
    If objSpedRegistroC101.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC101
    End If
    
    If objSpedRegistroC110.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC110
    End If
    
    If objSpedRegistroC111.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC111
    End If

    If objSpedRegistroC112.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC112
    End If
    
    If objSpedRegistroC113.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC113
    End If
    
    If objSpedRegistroC114.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC114
    End If
    
    If objSpedRegistroC115.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC115
    End If
    
    If objSpedRegistroC120.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC120
    End If
    
    If objSpedRegistroC130.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC130
    End If
    
    If objSpedRegistroC140.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC140
    End If
    
    If objSpedRegistroC141.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC141
    End If
    
    If objSpedRegistroC160.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC160
    End If

    If objSpedRegistroC170.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC170
    End If

    If objSpedRegistroC172.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC172
    End If

    If objSpedRegistroC190.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC190
    End If
    
    If giVersao > 12 Then
        colSpedRegistro.Add objSpedRegistroC191
    End If

    If objSpedRegistroC195.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC195
    End If

    If objSpedRegistroC300.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC300
    End If

    If objSpedRegistroC320.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC320
    End If

    If objSpedRegistroC321.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC321
    End If

    If objSpedRegistroC400.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC400
    End If

    If objSpedRegistroC405.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC405
    End If

    If objSpedRegistroC420.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC420
    End If

    If objSpedRegistroC425.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC425
    End If

    If objSpedRegistroC460.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC460
    End If

    If objSpedRegistroC470.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC470
    End If

    If objSpedRegistroC490.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC490
    End If

    If objSpedRegistroC500.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC500
    End If

    If objSpedRegistroC590.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC590
    End If

    If objSpedRegistroC990.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroC990
    End If

    If objSpedRegistroD001.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroD001
    End If

    If objSpedRegistroD100.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroD100
    End If

    If objSpedRegistroD190.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroD190
    End If

    If objSpedRegistroD500.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroD500
    End If

    If objSpedRegistroD590.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroD590
    End If
    If objSpedRegistroD990.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroD990
    End If

    If objSpedRegistroE001.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE001
    End If

    If objSpedRegistroE100.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE100
    End If

    If objSpedRegistroE110.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE110
    End If

    If objSpedRegistroE111.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE111
    End If

    If objSpedRegistroE116.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE116
    End If

    If objSpedRegistroE200.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE200
    End If

    If objSpedRegistroE210.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE210
    End If

    If objSpedRegistroE220.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE220
    End If

    If objSpedRegistroE250.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE250
    End If
    

    If objSpedRegistroE300.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE300
    End If
    
    If objSpedRegistroE310.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE310
    End If
    
    If objSpedRegistroE311.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE311
    End If
    
    If objSpedRegistroE312.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE312
    End If

    If objSpedRegistroE313.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE313
    End If
    
    If objSpedRegistroE316.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE316
    End If
    
    If objSpedRegistroE500.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE500
    End If

    If objSpedRegistroE510.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE510
    End If

    If objSpedRegistroE520.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE520
    End If

    If objSpedRegistroE530.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE530
    End If
    
    If objSpedRegistroE990.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroE990
    End If
    
    If objSpedRegistroG001.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroG001
    End If
    
    If objSpedRegistroG110.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroG110
    End If
    
    If objSpedRegistroG125.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroG125
    End If
    
    If objSpedRegistroG126.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroG126
    End If
    
    If objSpedRegistroG130.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroG130
    End If
    
    If objSpedRegistroG140.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroG140
    End If
    
    If objSpedRegistroG990.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroG990
    End If
    
    If objSpedRegistroH001.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroH001
    End If
    
    If objSpedRegistroH005.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroH005
    End If
    
    If objSpedRegistroH010.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroH010
    End If
    
    If objSpedRegistroH990.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroH990
    End If
    
    If objSpedRegistroK001.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroK001
    End If
    
    If objSpedRegistroK100.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroK100
    End If
    
    If objSpedRegistroK200.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroK200
    End If
    
    If objSpedRegistroK230.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroK230
    End If
    
    If objSpedRegistroK235.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroK235
    End If
    
    If objSpedRegistroK250.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroK250
    End If
    
    If objSpedRegistroK255.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroK255
    End If
        
    If objSpedRegistroK990.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistroK990
    End If
    
    If objSpedRegistro1001.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro1001
    End If

    If objSpedRegistro1010.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro1010
    End If

    If objSpedRegistro1100.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro1100
    End If

    If objSpedRegistro1105.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro1105
    End If

    If objSpedRegistro1400.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro1400
    End If

    If objSpedRegistro1990.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro1990
    End If
    
    If objSpedRegistro9001.iCount > 0 Then
        colSpedRegistro.Add objSpedRegistro9001
    End If
    
    'Bloco B e outros registros novos
    For Each objSpedRegistroAux In objEFD.colSpedRegistro
        objSpedRegistroAux.iCount = objSpedRegistroAux.lCount
        colSpedRegistro.Add objSpedRegistroAux
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + objSpedRegistroAux.lCount
    Next
    
    objSpedRegistro9900.sTipo = "9900"
    objSpedRegistro9900.iCount = colSpedRegistro.Count + 3

    colSpedRegistro.Add objSpedRegistro9900
    
    objSpedRegistro9990.sTipo = "9990"
    objSpedRegistro9990.iCount = 1
    
    colSpedRegistro.Add objSpedRegistro9990
    
    objSpedRegistro9999.sTipo = "9999"
    objSpedRegistro9999.iCount = 1
    
    colSpedRegistro.Add objSpedRegistro9999
    
    '*************** registro 9900 **********************
    lErro = SpedFiscal_9900(objSpedFiscal, colSpedRegistro) ', objEFD)
    If lErro <> SUCESSO Then gError 203365

    '*************** registro 9990 **********************
    lErro = SpedFiscal_9990(objSpedFiscal, objSpedRegistro9990, colSpedRegistro)
    If lErro <> SUCESSO Then gError 203366

    '*************** registro 9999 **********************
    lErro = SpedFiscal_9999(objSpedFiscal)
    If lErro <> SUCESSO Then gError 203367


    lErro = SpedFiscal_Gravar_Arq(sDiretorio, objSpedFiscal, objEFD)
    If lErro <> SUCESSO Then gError 203368

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Gera_Sped_Fiscal_Trans = SUCESSO
    
    Exit Function
    
Erro_Gera_Sped_Fiscal_Trans:

    Gera_Sped_Fiscal_Trans = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 203288
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 203289
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 203290, 203294, 203296 To 203300, 203305, 203306, 203316 To 203328, 203337 To 203340, 203343, 203346 To 203368, 204953, 210771

        Case 203291, 203292
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISVERSAOLEIAUTE", gErr)

        Case 203293
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISVERSAOLEIAUTE_NAO_CADASTRADO", gErr)
        
        Case 203295
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, iFilialEmpresa)
        
        Case 203301, 203302, 203329, 204129, 204130, 204915, 204916, 204918
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case 203303, 203304
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFEFEDPROTNFE", gErr)
        
        Case 203307, 203308
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFEFEDRETCANCNFE", gErr)
        
        Case 203309, 203310
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAGTODOS", gErr)
        
        Case 203311, 203312
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTODOS", gErr)
        
        Case 203313, 203314
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAONF", gErr)
        
        Case 203315, 203336, 203345
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case 203330, 203331, 203342, 204131, 204132
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ECF1", gErr)
        
        Case 203332, 203333, 203341
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJAARQFISMESTRE", gErr)
        
        Case 203334, 203335, 203344
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
        
        Case 203369
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 204127, 204128
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST1", gErr, iAno, iFilialEmpresa)
    
        Case 204133, 204134, 204141, 204914, 204917
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203370)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function
    
End Function


Private Function SpedFiscal_0150_Cliente(ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro0150 As ClassSpedRegistro, ByVal objSpedRegistro0175 As ClassSpedRegistro, sCodParticipante As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date, Optional ByRef iIEIsento As Integer = 0, Optional ByRef iIENaoContrib As Integer = 0) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 8) As Long
Dim tCliente As typeCliente
Dim tFilialCli As typeFilialCliente
Dim tEndereco As typeEndereco
Dim objPais As New ClassPais
Dim objCidades As New ClassCidades
Dim lSeqPai As Long
Dim dtDataAtual As Date
Dim sPais As String
Dim tClienteAlt As typeCliente
Dim tFilialCliAlt As typeFilialCliente
Dim tEnderecoAlt As typeEndereco
Dim objPaisAlt As New ClassPais
Dim objCidadesAlt As New ClassCidades
Dim lSeq As Long
Dim iDuplicado As Integer
Dim sIE As String


On Error GoTo Erro_SpedFiscal_0150_Cliente

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203371
    Next

    'verifica se o cliente ja esta cadastrado nos registros de participantes (0150)
    lErro = Comando_Executar(alComando(6), "SELECT Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0150' AND SUBSTRING(Registro, 7,16) = ?", lSeq, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, "C" & Format(lCliente, "0000000000") & Format(iFilialCli, "00000"))
    If lErro <> AD_SQL_SUCESSO Then gError 203372

    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203373
    
    If lErro = AD_SQL_SUCESSO Then

        sCodParticipante = "C" & Format(lCliente, "0000000000") & Format(iFilialCli, "00000")

    'se nao estiver cadastrado, cadastra agora
    ElseIf lErro = AD_SQL_SEM_DADOS Then
    
        '******************* registro 0150 **********************************
    
        tCliente.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
        tFilialCli.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
        tFilialCli.sCgc = String(STRING_CLIENTE_CGC, 0)
        tFilialCli.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        tFilialCli.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
        tEndereco.sEndereco = String(STRING_ENDERECO, 0)
        tEndereco.sBairro = String(STRING_BAIRRO, 0)
        tEndereco.sCidade = String(STRING_CIDADE, 0)
    
        'pegar os dados do cliente
        lErro = Comando_Executar(alComando(7), "SELECT RazaoSocial, Nome, CGC, InscricaoEstadual, InscricaoSuframa, E.Endereco, E.Bairro, E.Cidade, CodigoPais, IEIsento, IENaoContrib FROM Clientes AS C, FiliaisClientes AS F, Enderecos AS E WHERE C.Codigo = CodCliente AND E.Codigo = F.Endereco AND CodCliente = ? AND CodFilial = ?", _
        tCliente.sRazaoSocial, tFilialCli.sNome, tFilialCli.sCgc, tFilialCli.sInscricaoEstadual, tFilialCli.sInscricaoSuframa, tEndereco.sEndereco, tEndereco.sBairro, tEndereco.sCidade, tEndereco.iCodigoPais, iIEIsento, iIENaoContrib, lCliente, iFilialCli)
        If lErro <> AD_SQL_SUCESSO Then gError 203374
    
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203375
    
        objCidades.sDescricao = tEndereco.sCidade
        
        lErro = CF("Cidade_Le_Nome", objCidades)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203376
    
    
        objPais.iCodigo = tEndereco.iCodigoPais
        
        lErro = CF("Paises_Le", objPais)
        If lErro <> SUCESSO And lErro <> 47876 Then gError 203377
    
        objSpedFiscal.sTipo = "0150"
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        lSeqPai = objSpedFiscal.lSeq
        
        Call Formata_String_Numero2(tFilialCli.sInscricaoEstadual, sIE)
        
        
        sCodParticipante = "C" & Format(lCliente, "0000000000") & Format(iFilialCli, "00000")
        
        objSpedFiscal.sRegistro = "|0150|" & "C" & Format(lCliente, "0000000000") & Format(iFilialCli, "00000") & "|"
'        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Trim(tCliente.sRazaoSocial) & "-" & Trim(left(tFilialCli.sNome, 19)) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(tCliente.sRazaoSocial, 100) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & objPais.iCodBacen & "|" & IIf(Len(tFilialCli.sCgc) = 14, tFilialCli.sCgc, "") & "|" & IIf(Len(tFilialCli.sCgc) = 11, tFilialCli.sCgc, "") & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(sIE) & "|" & objCidades.sCodIBGE & "|" & SFF_Texto(tFilialCli.sInscricaoSuframa) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(tEndereco.sEndereco, 60) & "|||" & Desacentua_SpedFiscal(tEndereco.sBairro) & "|"
        
        objSpedRegistro0150.iCount = objSpedRegistro0150.iCount + 1
        
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203378
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203379
    
        '******************* registro 0175 **********************************
    
        tClienteAlt.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
        tFilialCliAlt.sCgc = String(STRING_CLIENTE_CGC, 0)
        tFilialCliAlt.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        tFilialCliAlt.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
        tEnderecoAlt.sEndereco = String(STRING_ENDERECO, 0)
        tEnderecoAlt.sBairro = String(STRING_BAIRRO, 0)
        tEnderecoAlt.sCidade = String(STRING_CIDADE, 0)
        sPais = String(STRING_PAISES_NOME, 0)
    
        'pegar os dados do clientehistorico
        lErro = Comando_Executar(alComando(8), "SELECT DataAtualizacao, RazaoSocial, Pais, CPF_CGC, InscricaoEstadual, Cidade, InscricaoSuframa, Endereco, Bairro FROM ClienteHistorico WHERE Codigo = ? AND Filial = ? AND DataAtualizacao >= ? And DataAtualizacao <= ? ORDER BY DataAtualizacao", _
        dtDataAtual, tClienteAlt.sRazaoSocial, sPais, tFilialCliAlt.sCgc, tFilialCliAlt.sInscricaoEstadual, tEnderecoAlt.sCidade, tFilialCliAlt.sInscricaoSuframa, tEnderecoAlt.sEndereco, tEnderecoAlt.sBairro, lCliente, iFilialCli, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203380
    
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203381
    
        Do While lErro = AD_SQL_SUCESSO
    
            objSpedFiscal.sTipo = "0175"
            
            If Len(tClienteAlt.sRazaoSocial) > 0 Then
            
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 3, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204154
                
                If iDuplicado = 0 Then
                
'                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|03|" & tClienteAlt.sRazaoSocial & "-" & left(tFilialCli.sNome, 19) & "|"
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|03|" & SFF_Texto(tClienteAlt.sRazaoSocial) & "|"
                    
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203382
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203383
            
            
                End If
            
            End If
            
            If Len(sPais) > 0 Then
            
                objPaisAlt.sNome = sPais
                
                lErro = CF("Pais_Le_Nome", objPaisAlt)
                If lErro <> SUCESSO And lErro <> 203267 Then gError 203384
        
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 4, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204155
                
                If iDuplicado = 0 Then
        
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|04|" & objPais.iCodBacen & "|"
                    
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203385
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203386
    
                End If
    
            End If
    
            If Len(tFilialCliAlt.sCgc) > 0 Then
                
                If Len(tFilialCliAlt.sCgc) = 11 Then
                
                    lErro = Checa_Duplicidade_0175(objSpedFiscal, 6, dtDataAtual, iDuplicado, lSeqPai)
                    If lErro <> SUCESSO Then gError 204156
                    
                    If iDuplicado = 0 Then
                
                
                        objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|06|" & tFilialCliAlt.sCgc & "|"
                        
                        
                        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                        objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                        
                        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203387
                        
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                        If lErro <> AD_SQL_SUCESSO Then gError 203388
                        
                    End If
                    
                Else
            
                    lErro = Checa_Duplicidade_0175(objSpedFiscal, 5, dtDataAtual, iDuplicado, lSeqPai)
                    If lErro <> SUCESSO Then gError 204157
                    
                    If iDuplicado = 0 Then
            
            
                        objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|05|" & tFilialCliAlt.sCgc & "|"
                        
                        
                        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                        objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                        
                        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203389
                        
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                        If lErro <> AD_SQL_SUCESSO Then gError 203390
                        
                    End If
                    
                End If
        
        
            End If
    
' ***********  inscricao estadual nao gera mais registro 0175 *************
'            If Len(tFilialCliAlt.sInscricaoEstadual) > 0 Then
'
'                lErro = Checa_Duplicidade_0175(objSpedFiscal, 7, dtDataAtual, iDuplicado, lSeqPai)
'                If lErro <> SUCESSO Then gError 204158
'
'                If iDuplicado = 0 Then
'
'                    Call Formata_String_Numero2(tFilialCliAlt.sInscricaoEstadual, sIE)
'
'                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|07|" & sIE & "|"
'
'
'                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
'                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
'
'                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203391
'
'                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
'                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
'                    If lErro <> AD_SQL_SUCESSO Then gError 203392
'
'                End If
'
'            End If
    
    
            If Len(tEnderecoAlt.sCidade) > 0 Then
            
                objCidadesAlt.sDescricao = tEnderecoAlt.sCidade
                
                lErro = CF("Cidade_Le_Nome", objCidadesAlt)
                If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203393
    
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 8, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204159
                
                If iDuplicado = 0 And Len(Trim(objCidadesAlt.sCodIBGE)) > 0 Then
    
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|08|" & objCidadesAlt.sCodIBGE & "|"
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203394
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203395
                    
                End If
    
            End If
    
    
            If Len(tFilialCliAlt.sInscricaoSuframa) > 0 Then
            
                
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 9, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204160
                
                If iDuplicado = 0 Then
                
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|09|" & SFF_Texto(tFilialCliAlt.sInscricaoSuframa) & "|"
                    
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203396
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203397
                    
                End If
    
            End If
    
            If Len(tEnderecoAlt.sEndereco) > 0 Then
            
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 10, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204161
                
                If iDuplicado = 0 Then
            
            
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|10|" & Desacentua_SpedFiscal(tEnderecoAlt.sEndereco) & "|"
                    
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203398
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203399
                    
                End If
    
            End If
    
            If Len(tEnderecoAlt.sBairro) > 0 Then
            
            
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 13, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204162
                
                If iDuplicado = 0 Then
            
            
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|13|" & Desacentua_SpedFiscal(tEnderecoAlt.sBairro) & "|"
                    
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203400
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203401
                    
                End If
    
            End If
    
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203402
    
        Loop
    
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_0150_Cliente = SUCESSO

    Exit Function
    
    
Erro_SpedFiscal_0150_Cliente:

    SpedFiscal_0150_Cliente = gErr
    
    Select Case gErr
    
        Case 203371
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203372, 203373
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISCAL", gErr)
    
        Case 203374, 203375
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
    
        Case 203376, 203377, 203384, 203393, 204154 To 204162
    
        Case 203378, 203382, 203385, 203387, 203389, 203391, 203394, 203396, 203398, 203400
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203379, 203383, 203386, 203388, 203390, 203392, 203395, 203397, 203399, 203401
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case 203380, 203381, 203402
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTEHISTORICO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203403)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_0150_Fornecedor(ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro0150 As ClassSpedRegistro, ByVal objSpedRegistro0175 As ClassSpedRegistro, sCodParticipante As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date, Optional ByRef iIEIsento As Integer = 0, Optional ByRef iIENaoContrib As Integer = 0) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 8) As Long
Dim tFornecedor As typeFornecedor
Dim tFilialForn As typeFilialFornecedor
Dim tEndereco As typeEndereco
Dim objPais As New ClassPais
Dim objCidades As New ClassCidades
Dim lSeqPai As Long
Dim dtDataAtual As Date
Dim sPais As String
Dim tFornecedorAlt As typeFornecedor
Dim tFilialFornAlt As typeFilialFornecedor
Dim tEnderecoAlt As typeEndereco
Dim objPaisAlt As New ClassPais
Dim objCidadesAlt As New ClassCidades
Dim lSeq As Long
Dim iDuplicado As Integer
Dim sIE As String


On Error GoTo Erro_SpedFiscal_0150_Fornecedor

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203404
    Next

    'verifica se o fornecedor ja esta cadastrado nos registros de participantes (0150)
    lErro = Comando_Executar(alComando(6), "SELECT Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0150' AND SUBSTRING(Registro, 7,16) = ?", lSeq, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, "F" & Format(lFornecedor, "0000000000") & Format(iFilialForn, "00000"))
    If lErro <> AD_SQL_SUCESSO Then gError 203405

    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203406


    If lErro = AD_SQL_SUCESSO Then
    
        sCodParticipante = "F" & Format(lFornecedor, "0000000000") & Format(iFilialForn, "00000")
    

    'se nao estiver cadastrado, cadastra agora
    ElseIf lErro = AD_SQL_SEM_DADOS Then
    
    
        '******************* registro 0150 **********************************
    
        tFornecedor.sRazaoSocial = String(STRING_FORNECEDOR_RAZAO_SOC, 0)
        tFilialForn.sNome = String(STRING_FILIAL_FORNECEDOR_NOME, 0)
        tFilialForn.sCgc = String(STRING_CGC, 0)
        tFilialForn.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        tFilialForn.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
        tEndereco.sEndereco = String(STRING_ENDERECO, 0)
        tEndereco.sBairro = String(STRING_BAIRRO, 0)
        tEndereco.sCidade = String(STRING_CIDADE, 0)
    
    
        'pegar os dados do fornecedor
        lErro = Comando_Executar(alComando(7), "SELECT RazaoSocial, Nome, CGC, InscricaoEstadual, InscricaoSuframa, E.Endereco, E.Bairro, E.Cidade, CodigoPais,IEIsento, IENaoContrib FROM Fornecedores AS F, FiliaisFornecedores AS FF, Enderecos AS E WHERE F.Codigo = CodFornecedor AND E.Codigo = FF.Endereco AND CodFornecedor = ? AND CodFilial = ?", _
        tFornecedor.sRazaoSocial, tFilialForn.sNome, tFilialForn.sCgc, tFilialForn.sInscricaoEstadual, tFilialForn.sInscricaoSuframa, tEndereco.sEndereco, tEndereco.sBairro, tEndereco.sCidade, tEndereco.iCodigoPais, iIEIsento, iIENaoContrib, lFornecedor, iFilialForn)
        If lErro <> AD_SQL_SUCESSO Then gError 203407
    
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203408
    
        objCidades.sDescricao = tEndereco.sCidade
        
        lErro = CF("Cidade_Le_Nome", objCidades)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203409
    
    
        objPais.iCodigo = tEndereco.iCodigoPais
        
        lErro = CF("Paises_Le", objPais)
        If lErro <> SUCESSO And lErro <> 47876 Then gError 203410
    
    
        objSpedFiscal.sTipo = "0150"
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        lSeqPai = objSpedFiscal.lSeq
        
        sCodParticipante = "F" & Format(lFornecedor, "0000000000") & Format(iFilialForn, "00000")
        
        Call Formata_String_Numero2(tFilialForn.sInscricaoEstadual, sIE)
        
        objSpedFiscal.sRegistro = "|0150|" & "F" & Format(lFornecedor, "0000000000") & Format(iFilialForn, "00000") & "|"
'        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Trim(tFornecedor.sRazaoSocial) & "-" & Trim(left(tFilialForn.sNome, 19)) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(tFornecedor.sRazaoSocial, 100) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & objPais.iCodBacen & "|" & IIf(Len(tFilialForn.sCgc) = 14, tFilialForn.sCgc, "") & "|" & IIf(Len(tFilialForn.sCgc) = 11, tFilialForn.sCgc, "") & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(sIE) & "|" & objCidades.sCodIBGE & "|" & SFF_Texto(tFilialForn.sInscricaoSuframa) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(tEndereco.sEndereco, 60) & "|||" & Desacentua_SpedFiscal(tEndereco.sBairro) & "|"
        
        objSpedRegistro0150.iCount = objSpedRegistro0150.iCount + 1
        
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203411
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203412
    
        '******************* registro 0175 **********************************
    
        tFornecedorAlt.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
        tFilialFornAlt.sCgc = String(STRING_CLIENTE_CGC, 0)
        tFilialFornAlt.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        tFilialFornAlt.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
        tEnderecoAlt.sEndereco = String(STRING_ENDERECO, 0)
        tEnderecoAlt.sBairro = String(STRING_BAIRRO, 0)
        tEnderecoAlt.sCidade = String(STRING_CIDADE, 0)
        sPais = String(STRING_PAISES_NOME, 0)
    
        'pegar os dados do fornecedorhistorico
        lErro = Comando_Executar(alComando(8), "SELECT DataAtualizacao, RazaoSocial, Pais, CPF_CGC, InscricaoEstadual, Cidade, InscricaoSuframa, Endereco, Bairro FROM FornecedorHistorico WHERE Codigo = ? AND Filial = ? AND DataAtualizacao >= ? And DataAtualizacao <= ? ORDER BY DataAtualizacao", _
        dtDataAtual, tFornecedorAlt.sRazaoSocial, sPais, tFilialFornAlt.sCgc, tFilialFornAlt.sInscricaoEstadual, tEnderecoAlt.sCidade, tFilialFornAlt.sInscricaoSuframa, tEnderecoAlt.sEndereco, tEnderecoAlt.sBairro, lFornecedor, iFilialForn, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203413
    
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203414
    
    
        Do While lErro = AD_SQL_SUCESSO
    
        
            objSpedFiscal.sTipo = "0175"
            
            If Len(tFornecedorAlt.sRazaoSocial) > 0 Then
                
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 3, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204163
                
                If iDuplicado = 0 Then
                
                
'                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|03|" & tFornecedorAlt.sRazaoSocial & "-" & left(tFilialForn.sNome, 19) & "|"
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|03|" & SFF_Texto(tFornecedorAlt.sRazaoSocial) & "|"
                    
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203415
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203416
                    
                End If
            
            End If
            
            If Len(sPais) > 0 Then
            
                objPaisAlt.sNome = sPais
                
                lErro = CF("Pais_Le_Nome", objPaisAlt)
                If lErro <> SUCESSO And lErro <> 203267 Then gError 203417
            
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 4, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204164
                
                If iDuplicado = 0 Then
            
            
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|04|" & objPais.iCodBacen & "|"
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203418
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203419
                    
                End If
    
            End If
    
            If Len(tFilialFornAlt.sCgc) > 0 Then
                
                If Len(tFilialFornAlt.sCgc) = 11 Then
                
                    lErro = Checa_Duplicidade_0175(objSpedFiscal, 6, dtDataAtual, iDuplicado, lSeqPai)
                    If lErro <> SUCESSO Then gError 204165
                    
                    If iDuplicado = 0 Then
                
                        objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|06|" & tFilialFornAlt.sCgc & "|"
                        
                        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                        objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                        
                        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203420
                        
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                        If lErro <> AD_SQL_SUCESSO Then gError 203421
                        
                    End If
                    
                Else
            
                    lErro = Checa_Duplicidade_0175(objSpedFiscal, 5, dtDataAtual, iDuplicado, lSeqPai)
                    If lErro <> SUCESSO Then gError 204166
                    
                    If iDuplicado = 0 Then
            
                        objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|05|" & tFilialFornAlt.sCgc & "|"
                        
                        
                        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                        objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                        
                        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203422
                        
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                        If lErro <> AD_SQL_SUCESSO Then gError 203423
                        
                    End If
                    
                End If
        
        
            End If
    
'***** inscricao estadual nao gera + registro 0175 *********
'            If Len(tFilialFornAlt.sInscricaoEstadual) > 0 Then
'
'                lErro = Checa_Duplicidade_0175(objSpedFiscal, 7, dtDataAtual, iDuplicado, lSeqPai)
'                If lErro <> SUCESSO Then gError 204167
'
'                If iDuplicado = 0 Then
'
'                   Call Formata_String_Numero2(tFilialForn.sInscricaoEstadual, sIE)
'
'
'                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|07|" & sIE & "|"
'
'                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
'                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
'
'                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203424
'
'                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
'                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
'                    If lErro <> AD_SQL_SUCESSO Then gError 203425
'
'                End If
'
'            End If
    
    
            If Len(tEnderecoAlt.sCidade) > 0 Then
            
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 8, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204168
                
                If iDuplicado = 0 And Len(Trim(objCidadesAlt.sCodIBGE)) > 0 Then
            
            
                    objCidadesAlt.sDescricao = tEnderecoAlt.sCidade
                    
                    lErro = CF("Cidade_Le_Nome", objCidadesAlt)
                    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203426
        
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|08|" & objCidadesAlt.sCodIBGE & "|"
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203427
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203428
                    
                End If
    
            End If
    
    
            If Len(tFilialFornAlt.sInscricaoSuframa) > 0 Then
            
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 9, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204169
                
                If iDuplicado = 0 Then
            
            
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|09|" & SFF_Texto(tFilialFornAlt.sInscricaoSuframa) & "|"
                    
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203429
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203430
                    
                End If
    
            End If
    
            If Len(tEnderecoAlt.sEndereco) > 0 Then
            
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 10, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204170
                
                If iDuplicado = 0 Then
                
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|10|" & Desacentua_SpedFiscal(tEnderecoAlt.sEndereco) & "|"
                    
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203431
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203432
                    
                End If
    
            End If
    
            If Len(tEnderecoAlt.sBairro) > 0 Then
            
                lErro = Checa_Duplicidade_0175(objSpedFiscal, 13, dtDataAtual, iDuplicado, lSeqPai)
                If lErro <> SUCESSO Then gError 204171
                
                If iDuplicado = 0 Then
            
                    objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|13|" & Desacentua_SpedFiscal(tEnderecoAlt.sBairro) & "|"
                    
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                    
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203433
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                    If lErro <> AD_SQL_SUCESSO Then gError 203434
                    
                End If
    
            End If
    
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203435
    
        Loop
    
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_0150_Fornecedor = SUCESSO

    Exit Function
    
    
Erro_SpedFiscal_0150_Fornecedor:

    SpedFiscal_0150_Fornecedor = gErr

   Select Case gErr
    
        Case 203404
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203405, 203406
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISCAL", gErr)
    
        Case 203407, 203408
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORES", gErr)
    
        Case 203409, 203410, 203417, 203426, 204163 To 204171
    
        Case 203411, 203415, 203418, 203420, 203422, 203424, 203427, 203429, 203431, 203433
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203412, 203416, 203419, 203421, 203423, 203425, 203428, 203430, 203432, 203434
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case 203413, 203414, 203435
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORHISTORICO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203436)
     
    End Select
 
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_0150_Transportadora(ByVal iTransportadora As Integer, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro0150 As ClassSpedRegistro, ByVal objSpedRegistro0175 As ClassSpedRegistro, sCodParticipante As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 8) As Long
Dim tTransp As typeTransportadora
Dim tEndereco As typeEndereco
Dim objPais As New ClassPais
Dim objCidades As New ClassCidades
Dim lSeqPai As Long
Dim dtDataAtual As Date
Dim sPais As String
Dim tTranspAlt As typeTransportadora
Dim tEnderecoAlt As typeEndereco
Dim objPaisAlt As New ClassPais
Dim objCidadesAlt As New ClassCidades
Dim lSeq As Long
Dim iDuplicado As Integer
Dim sIE As String


On Error GoTo Erro_SpedFiscal_0150_Transportadora

    If iTransportadora <> 0 Then

        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 203437
        Next
    
        'verifica se o cliente ja esta cadastrado nos registros de participantes (0150)
        lErro = Comando_Executar(alComando(6), "SELECT Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0150' AND SUBSTRING(Registro, 7,6) = ?", lSeq, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, "T" & Format(iTransportadora, "00000"))
        If lErro <> AD_SQL_SUCESSO Then gError 203438
    
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203439
        
        If lErro = AD_SQL_SUCESSO Then
    
            sCodParticipante = "T" & Format(iTransportadora, "00000")
    
        'se nao estiver cadastrado, cadastra agora
        ElseIf lErro = AD_SQL_SEM_DADOS Then
        
        
            '******************* registro 0150 **********************************
        
            tTransp.sNome = String(STRING_TRANSPORTADORA_NOME, 0)
            tTransp.sCgc = String(STRING_CGC, 0)
            tTransp.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
            tEndereco.sEndereco = String(STRING_ENDERECO, 0)
            tEndereco.sBairro = String(STRING_BAIRRO, 0)
            tEndereco.sCidade = String(STRING_CIDADE, 0)
        
        
            'pegar os dados da transportadora
            lErro = Comando_Executar(alComando(7), "SELECT Nome, CGC, InscricaoEstadual, E.Endereco, E.Bairro, E.Cidade, CodigoPais FROM Transportadoras AS T, Enderecos AS E WHERE E.Codigo = T.Endereco AND T.Codigo = ? ", _
            tTransp.sNome, tTransp.sCgc, tTransp.sInscricaoEstadual, tEndereco.sEndereco, tEndereco.sBairro, tEndereco.sCidade, tEndereco.iCodigoPais, iTransportadora)
            If lErro <> AD_SQL_SUCESSO Then gError 203440
        
            lErro = Comando_BuscarPrimeiro(alComando(7))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203441
        
            objCidades.sDescricao = tEndereco.sCidade
            
            lErro = CF("Cidade_Le_Nome", objCidades)
            If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203442
        
        
            objPais.iCodigo = tEndereco.iCodigoPais
            
            lErro = CF("Paises_Le", objPais)
            If lErro <> SUCESSO And lErro <> 47876 Then gError 203443
        
        
            objSpedFiscal.sTipo = "0150"
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            lSeqPai = objSpedFiscal.lSeq
            
            sCodParticipante = "T" & Format(iTransportadora, "00000")
            
            Call Formata_String_Numero2(tTransp.sInscricaoEstadual, sIE)
            
            objSpedFiscal.sRegistro = "|0150|" & sCodParticipante & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(tTransp.sNome, 100) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & objPais.iCodBacen & "|" & IIf(Len(tTransp.sCgc) = 14, tTransp.sCgc, "") & "|" & IIf(Len(tTransp.sCgc) = 11, tTransp.sCgc, "") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(sIE) & "|" & objCidades.sCodIBGE & "||"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(tEndereco.sEndereco, 60) & "|||" & Desacentua_SpedFiscal(tEndereco.sBairro) & "|"
            
            objSpedRegistro0150.iCount = objSpedRegistro0150.iCount + 1
            
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203444
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
            If lErro <> AD_SQL_SUCESSO Then gError 203445
        
            '******************* registro 0175 **********************************
        
            tTranspAlt.sNome = String(STRING_TRANSPORTADORA_NOME, 0)
            tTranspAlt.sCgc = String(STRING_CGC, 0)
            tTranspAlt.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
            tEnderecoAlt.sEndereco = String(STRING_ENDERECO, 0)
            tEnderecoAlt.sBairro = String(STRING_BAIRRO, 0)
            tEnderecoAlt.sCidade = String(STRING_CIDADE, 0)
            sPais = String(STRING_PAISES_NOME, 0)
        
        
            'pegar os dados do transportadorahistorico
            lErro = Comando_Executar(alComando(8), "SELECT DataAtualizacao, RazaoSocial, Pais, CGC, InscricaoEstadual, Cidade, Endereco, Bairro FROM TransportadoraHistorico WHERE CodTransp = ? AND DataAtualizacao >= ? And DataAtualizacao <= ? ORDER BY DataAtualizacao", _
            dtDataAtual, tTranspAlt.sNome, sPais, tTranspAlt.sCgc, tTranspAlt.sInscricaoEstadual, tEnderecoAlt.sCidade, tEnderecoAlt.sEndereco, tEnderecoAlt.sBairro, iTransportadora, dtDataIni, dtDataFim)
            If lErro <> AD_SQL_SUCESSO Then gError 203446
        
            lErro = Comando_BuscarPrimeiro(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203447
        
            Do While lErro = AD_SQL_SUCESSO
        
                objSpedFiscal.sTipo = "0175"
                
                If Len(tTranspAlt.sNome) > 0 Then
                    
                    lErro = Checa_Duplicidade_0175(objSpedFiscal, 3, dtDataAtual, iDuplicado, lSeqPai)
                    If lErro <> SUCESSO Then gError 204172
                    
                    If iDuplicado = 0 Then
                    
                    
                        objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|03|" & SFF_Texto(tTranspAlt.sNome) & "|"
                        
                        
                        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                        objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                        
                        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203448
                        
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                        If lErro <> AD_SQL_SUCESSO Then gError 203449
                
                    End If
                
                
                End If
                
                If Len(sPais) > 0 Then
                
                    objPaisAlt.sNome = sPais
                    
                    lErro = CF("Pais_Le_Nome", objPaisAlt)
                    If lErro <> SUCESSO And lErro <> 203267 Then gError 203450
                
                    lErro = Checa_Duplicidade_0175(objSpedFiscal, 4, dtDataAtual, iDuplicado, lSeqPai)
                    If lErro <> SUCESSO Then gError 204173
                    
                    If iDuplicado = 0 Then
                
                
                        objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|04|" & objPais.iCodBacen & "|"
                        
                        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                        objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                        
                        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203451
                        
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                        If lErro <> AD_SQL_SUCESSO Then gError 203452
                        
                    End If
        
                End If
        
                If Len(tTranspAlt.sCgc) > 0 Then
                    
                    If Len(tTranspAlt.sCgc) = 11 Then
                    
                        lErro = Checa_Duplicidade_0175(objSpedFiscal, 6, dtDataAtual, iDuplicado, lSeqPai)
                        If lErro <> SUCESSO Then gError 204174
                        
                        If iDuplicado = 0 Then
                    
                    
                            objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|06|" & tTranspAlt.sCgc & "|"
                            
                            
                            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                            objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                            
                            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203453
                            
                            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                            If lErro <> AD_SQL_SUCESSO Then gError 203454
                            
                        End If
                        
                    Else
                
                        lErro = Checa_Duplicidade_0175(objSpedFiscal, 5, dtDataAtual, iDuplicado, lSeqPai)
                        If lErro <> SUCESSO Then gError 204175
                        
                        If iDuplicado = 0 Then
                
                            objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|05|" & tTranspAlt.sCgc & "|"
                            
                            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                            objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                            
                            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203455
                            
                            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                            If lErro <> AD_SQL_SUCESSO Then gError 203456
                        
                        End If
                        
                        
                    End If
            
            
                End If
        
        
'******* inscricao estadual nao gera + registro 0175 *******
'                If Len(tTranspAlt.sInscricaoEstadual) > 0 Then
'
'                    lErro = Checa_Duplicidade_0175(objSpedFiscal, 7, dtDataAtual, iDuplicado, lSeqPai)
'                    If lErro <> SUCESSO Then gError 204176
'
'                    If iDuplicado = 0 Then
'
'                        Call Formata_String_Numero2(tTransp.sInscricaoEstadual, sIE)
'
'
'                        objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|07|" & sIE & "|"
'
'
'                        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
'                        objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
'
'                        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203457
'
'                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
'                        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
'                        If lErro <> AD_SQL_SUCESSO Then gError 203458
'
'                    End If
'
'                End If
        
        
                If Len(tEnderecoAlt.sCidade) > 0 Then
                
                    objCidadesAlt.sDescricao = tEnderecoAlt.sCidade
                    
                    lErro = CF("Cidade_Le_Nome", objCidadesAlt)
                    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203459
        
                    lErro = Checa_Duplicidade_0175(objSpedFiscal, 8, dtDataAtual, iDuplicado, lSeqPai)
                    If lErro <> SUCESSO Then gError 204177
                    
                    If iDuplicado = 0 Then
        
                        objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|08|" & objCidadesAlt.sCodIBGE & "|"
                        
                        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                        objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                        
                        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203460
                        
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                        If lErro <> AD_SQL_SUCESSO Then gError 203461
        
                    End If
        
                End If
        
        
                If Len(tEnderecoAlt.sEndereco) > 0 Then
                
                    lErro = Checa_Duplicidade_0175(objSpedFiscal, 10, dtDataAtual, iDuplicado, lSeqPai)
                    If lErro <> SUCESSO Then gError 204178
                    
                    If iDuplicado = 0 Then
                
                        objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|10|" & Desacentua_SpedFiscal(tEnderecoAlt.sEndereco) & "|"
                        
                        
                        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                        objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                        
                        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203462
                        
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                        If lErro <> AD_SQL_SUCESSO Then gError 203463
        
                    End If
        
                End If
        
                If Len(tEnderecoAlt.sBairro) > 0 Then
                
                    lErro = Checa_Duplicidade_0175(objSpedFiscal, 13, dtDataAtual, iDuplicado, lSeqPai)
                    If lErro <> SUCESSO Then gError 204179
                    
                    If iDuplicado = 0 Then
                
                        objSpedFiscal.sRegistro = "|0175|" & Format(dtDataAtual, "ddmmyyyy") & "|13|" & Desacentua_SpedFiscal(tEnderecoAlt.sBairro) & "|"
                        
                        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                        objSpedRegistro0175.iCount = objSpedRegistro0175.iCount + 1
                        
                        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203464
                        
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                        If lErro <> AD_SQL_SUCESSO Then gError 203465
                        
                    End If
        
                End If
        
                lErro = Comando_BuscarProximo(alComando(8))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203466
        
            Loop
        
        End If
    
        For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
        Next

    End If

    SpedFiscal_0150_Transportadora = SUCESSO

    Exit Function
    
Erro_SpedFiscal_0150_Transportadora:

    SpedFiscal_0150_Transportadora = gErr
    
    Select Case gErr
    
        Case 203437
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203438, 203439
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISCAL", gErr)
    
        Case 203440, 203441
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORA2", gErr)
    
        Case 203442, 203443, 203450, 203459, 204172 To 204179
    
        Case 203444, 203448, 203451, 203453, 203455, 203457, 203460, 203462, 203464
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203445, 203449, 203452, 203454, 203456, 203458, 203461, 203463, 203465
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case 203446, 203447, 203466
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORAHISTORICO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203467)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_0200(ByVal sProduto As String, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro0200 As ClassSpedRegistro, ByVal objSpedRegistro0205 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal objSpedRegistro0190 As ClassSpedRegistro, ByVal objSpedRegistro0220 As ClassSpedRegistro, ByVal colUM As Collection, lSeqPai As Long, ByVal objFisSpedAux As ClassFISSpedAux) As Long


Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 8) As Long
Dim sRegistro As String
Dim tProduto As typeProduto
Dim iNatureza As Integer
Dim tProdutoAlt As typeProduto
Dim dtDataIniAlt As Date
Dim dtDataAtual As Date
Dim sUMResult As String


On Error GoTo Erro_SpedFiscal_0200

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203468
    Next

    sRegistro = String(500, 0)

    'verifica se o produto ja esta cadastrado nos registros  (0200)
    lErro = Comando_Executar(alComando(6), "SELECT Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0200' AND SUBSTRING(Registro, 7,60) LIKE ?", lSeqPai, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, Trim(sProduto) & "|%")
    If lErro <> AD_SQL_SUCESSO Then gError 203469

    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203470
    
    'se nao estiver cadastrado, cadastra agora
    If lErro = AD_SQL_SEM_DADOS Then
    
    
        '******************* registro 0200 **********************************
    
        tProduto.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        tProduto.sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        tProduto.sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        tProduto.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        tProduto.sGenero = String(STRING_PRODUTO_GENERO, 2)
        tProduto.sCEST = String(STRING_CEST, 0)
    
    
        'pegar os dados do produto
        lErro = Comando_Executar(alComando(7), "SELECT Descricao, CodigoBarras, SiglaUMEstoque, Natureza, IPICodigo, Genero, TipoServico, ClasseUM, CEST  FROM Produtos WHERE Codigo = ? ", _
        tProduto.sDescricao, tProduto.sCodigoBarras, tProduto.sSiglaUMEstoque, tProduto.iNatureza, tProduto.sIPICodigo, tProduto.sGenero, tProduto.iTipoServico, tProduto.iClasseUM, tProduto.sCEST, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 203471
    
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203472
    
        Select Case tProduto.iNatureza
        
            Case 1 'Materia Prima
                iNatureza = 1
                
            Case 2 'Produto Intermediario
                iNatureza = 6
                
            Case 3 'Embalagem
                iNatureza = 2
                
            Case 4 'Produto Acabado
                iNatureza = 4
                
            Case 5 'Revenda
                iNatureza = 0
                
            Case 6, 7 'Manutencao, Outros
                iNatureza = 99
                
            Case 8 'Serviços
                iNatureza = 9
                
            Case 9 'Em Processo
                iNatureza = 3
                
            Case 10 'SubProduto
                iNatureza = 5
                
            Case 11 'Uso e Consumo
                iNatureza = 7
                
            Case 12 'Ativo
                iNatureza = 8
                
            Case 13 'Outros Insumos
                iNatureza = 10
    
            Case Else
                gError 99999
    
        End Select
    
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        lSeqPai = objSpedFiscal.lSeq
        
        lErro = SpedFiscal_0190(tProduto.sSiglaUMEstoque, tProduto.iClasseUM, objSpedFiscal, objSpedRegistro0190, objSpedRegistro0220, colUM, sUMResult, sProduto, lSeqPai)
        If lErro <> SUCESSO Then gError 203473
        
        objSpedFiscal.sTipo = "0200"
        
        objSpedFiscal.sRegistro = "|0200|" & SFF_Texto(sProduto) & "|" & SFF_Texto(tProduto.sDescricao) & "|" & tProduto.sCodigoBarras & "||" & sUMResult & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(iNatureza, "00") & "|" & SFF_Texto(tProduto.sIPICodigo, 8) & "||" & IIf(tProduto.sGenero = "0", "", Trim(tProduto.sGenero)) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tProduto.sISSQN & "||" 'tProduto.iTipoServico & "||"
        
         If giVersao >= 11 Then objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Trim(tProduto.sCEST) & "|"
        
        objSpedRegistro0200.iCount = objSpedRegistro0200.iCount + 1
        
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203474
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, lSeqPai, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203475
    
    
        '******************* registro 0205 **********************************
    
        tProdutoAlt.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    
        'dtDataIniAlt = dtDataIni
        dtDataIniAlt = DateAdd("d", -1, dtDataFim)
    
        'pegar os dados do produtohistorico
        'lErro = Comando_Executar(alComando(8), "SELECT DataAtualizacao, DescProduto FROM ProdutoHistorico WHERE CodProduto = ? AND DataAtualizacao >= ? And DataAtualizacao <= ? ORDER BY DataAtualizacao",
        lErro = Comando_Executar(alComando(8), "SELECT DataAtualizacao, DescProduto FROM ProdutoHistorico WHERE CodProduto = ? AND DataAtualizacao <= ? ORDER BY DataAtualizacao DESC, NumIntDoc DESC ", _
        dtDataAtual, tProdutoAlt.sDescricao, sProduto, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 203476
    
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203477
    
        Do While lErro = AD_SQL_SUCESSO
    
            If dtDataIni > dtDataAtual Then dtDataAtual = dtDataIni
            
            'Se o mesmo produto foi alterado diversas vezes no mesmo dia só considera
            ' a última mexida
            If dtDataIniAlt >= dtDataAtual Then
        
                objSpedFiscal.sTipo = "0205"
                'objSpedFiscal.sRegistro = "|0205|" & Trim(tProdutoAlt.sDescricao) & "|" & Format(dtDataIniAlt, "ddmmyyyy") & "|" & Format(dtDataAtual, "ddmmyyyy") & "||"
                objSpedFiscal.sRegistro = "|0205|" & SFF_Texto(tProdutoAlt.sDescricao) & "|" & Format(dtDataAtual, "ddmmyyyy") & "|" & Format(dtDataIniAlt, "ddmmyyyy") & "||"
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistro0205.iCount = objSpedRegistro0205.iCount + 1
                
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203478
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                If lErro <> AD_SQL_SUCESSO Then gError 203479
                
            End If
            
            'dtDataIniAlt = dtDataAtual
            
            If dtDataIni = dtDataAtual Then Exit Do
            
            dtDataIniAlt = DateAdd("d", -1, dtDataAtual) 'Não pode ficar com a data final igual a data inicial de outro
            
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203480
    
        Loop
    
        If Year(objFisSpedAux.dtDataIni) >= 2016 And objFisSpedAux.iIncluiRCPE = MARCADO Then
            If iNatureza = 3 Or iNatureza = 4 Then
                lErro = SpedFiscal_0210(objFisSpedAux, sProduto, lSeqPai)
                If lErro <> SUCESSO Then gError 203473
            End If
            
        End If
        
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_0200 = SUCESSO

    Exit Function
    
    
Erro_SpedFiscal_0200:

    SpedFiscal_0200 = gErr
    
    Select Case gErr
    
        Case 203468
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203469, 203470
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISCAL", gErr)
    
        Case 203471, 203472
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
    
        Case 203473
    
        Case 203474, 203478
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203475, 203479
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case 203476, 203477, 203480
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOHISTORICO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203481)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_0210(ByVal objFisSpedAux As ClassFISSpedAux, ByVal sProduto As String, ByVal lSeqPai As Long) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long, lSeq0200 As Long
Dim iIndice As Integer, objProduto As ClassProduto
Dim lTotal As Long, objSpedFiscal As ClassSpedFiscal
Dim objKit As New ClassKit, objProdutoKit As ClassProdutoKit
Dim dFatorConversao As Double, dQuantidadePai As Double, dQuantidadeFilho As Double
Dim objPK As ClassProdutoKit, objPKAux As ClassProdutoKit, colPK As New Collection, bAchou As Boolean

On Error GoTo Erro_SpedFiscal_0210
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201226
    Next
    
    objKit.sProdutoRaiz = sProduto

    lErro = CF("Kit_Le_Padrao", objKit)
    If lErro <> SUCESSO And lErro <> 106304 Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then

        'tem que ler o proprio cara e os seus filhos
        lErro = CF("ProdutoKit_Le_PrimeiroNivel_Versao", objKit)
        If lErro <> SUCESSO And lErro <> 103203 Then gError ERRO_SEM_MENSAGEM

        'percorre os elementos componentes do Kit Lido
        For Each objProdutoKit In objKit.colComponentes
        
            Set objProduto = New ClassProduto
        
            objProduto.sCodigo = objProdutoKit.sProduto
                        
            'Lê o produto para descobrir as unidades de medidas associadas
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            If objProduto.iNatureza = NATUREZA_PROD_PRODUTO_REVENDA Or objProduto.iNatureza = NATUREZA_PROD_MATERIA_PRIMA Or _
                objProduto.iNatureza = NATUREZA_PROD_EMBALAGENS Or objProduto.iNatureza = NATUREZA_PROD_PRODUTO_EM_PROCESSO Or _
                objProduto.iNatureza = NATUREZA_PROD_PRODUTO_INTERMEDIARIO Or objProduto.iNatureza = NATUREZA_PROD_OUTROS_INSUMOS Or _
                objProduto.iNatureza = NATUREZA_PROD_PRODUTO_ACABADO Or objProduto.iNatureza = NATUREZA_PROD_SUBPRODUTO Then
                    
                'Descobre o fator de conversao da UM recebida p/UM de estoque do produto
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProduto.sSiglaUMEstoque, dFatorConversao)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                    
                    dQuantidadePai = (1 / objProdutoKit.dQuantidade * dFatorConversao) / (1 - objProdutoKit.dPercentualPerda)
                    
                Else
                    
                    dQuantidadeFilho = dQuantidadePai * (objProdutoKit.dQuantidade / dFatorConversao)
                    
                    bAchou = False
                    For Each objPKAux In colPK
                        If objPKAux.sProduto = objProdutoKit.sProduto Then
                            bAchou = True
                            objPKAux.dQuantidade = objPKAux.dQuantidade + dQuantidadeFilho
                        End If
                    Next
                    If Not bAchou Then
                        Set objPK = New ClassProdutoKit
                        objPK.sProduto = objProdutoKit.sProduto
                        objPK.dQuantidade = dQuantidadeFilho
                        objPK.dPercentualPerda = objProdutoKit.dPercentualPerda
                        colPK.Add objPK
                    End If
                    
'                    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
'
'                    objSpedFiscal.sTipo = "0210"
'                    objFisSpedAux.objSpedRegistro0210.sTipo = objSpedFiscal.sTipo
'                    objFisSpedAux.objSpedRegistro0210.iCount = objFisSpedAux.objSpedRegistro0210.iCount + 1
'
'                    '01 REG Texto fixo contendo "0210" C 4 - O
'                    '02 COD_ITEM_COMP Código do item componente/insumo (campo 02 do Registro 0200)
'                    '03 QTD_COMP Quantidade do item componente/insumo para se produzir uma unidade do item composto/resultante
'                    '04 PERDA Perda/quebra normal percentual do insumo/componente para se produzir uma unidade do item composto/resultante
'                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
'                    objSpedFiscal.sRegistro = "|0210|" '01
'                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(objProdutoKit.sProduto) & "|" '02
'                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dQuantidadeFilho, "##0.000000") & "|" '03
'                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(objProdutoKit.dPercentualPerda, "##0.0000") & "|" '04
'
'                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 201227
'
'                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
'                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
'                    If lErro <> AD_SQL_SUCESSO Then gError 201228
'
'                    lErro = SpedFiscal_0200(objProdutoKit.sProduto, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, lSeq0200, objFisSpedAux)
'                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                End If
                
            End If
            
        Next
        
        For Each objPK In colPK
        
             Set objSpedFiscal = objFisSpedAux.objSpedFiscal
             
             objSpedFiscal.sTipo = "0210"
             objFisSpedAux.objSpedRegistro0210.sTipo = objSpedFiscal.sTipo
             objFisSpedAux.objSpedRegistro0210.iCount = objFisSpedAux.objSpedRegistro0210.iCount + 1
             
             '01 REG Texto fixo contendo "0210" C 4 - O
             '02 COD_ITEM_COMP Código do item componente/insumo (campo 02 do Registro 0200)
             '03 QTD_COMP Quantidade do item componente/insumo para se produzir uma unidade do item composto/resultante
             '04 PERDA Perda/quebra normal percentual do insumo/componente para se produzir uma unidade do item composto/resultante
             objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
             objSpedFiscal.sRegistro = "|0210|" '01
             objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(objPK.sProduto) & "|" '02
             objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(objPK.dQuantidade, "##0.000000") & "|" '03
             objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(objPK.dPercentualPerda, "##0.0000") & "|" '04
        
             If Len(objSpedFiscal.sRegistro) > 500 Then gError 201227
             
             lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
             objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
             If lErro <> AD_SQL_SUCESSO Then gError 201228
             
             lErro = SpedFiscal_0200(objPK.sProduto, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, lSeq0200, objFisSpedAux)
             If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
        Next
        
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_0210 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_0210:

   SpedFiscal_0210 = gErr

   Select Case gErr
    
        Case 201226
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201227
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201228
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204086)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function PIS_CST(ByRef iCST As Integer, ByRef tTribItemNF As typeTribItemNF) As Long
    If tTribItemNF.dPISCredito > 0 Then
        iCST = 1
    Else
        iCST = 4
    End If
    PIS_CST = SUCESSO
End Function

Function PIS_Aliquota(ByRef dAliquota As Double, ByVal objFilialEmpresa As AdmFiliais) As Long
    If objFilialEmpresa.iPISNaoCumulativo = 1 Then
        dAliquota = 0.0165
    Else
        dAliquota = 0.0065
    End If
    PIS_Aliquota = SUCESSO
End Function

Private Function COFINS_CST(ByRef iCST As Integer, ByRef tTribItemNF As typeTribItemNF) As Long
    If tTribItemNF.dCOFINSCredito > 0 Then
        iCST = 1
    Else
        iCST = 4
    End If
    COFINS_CST = SUCESSO
End Function

Function COFINS_Aliquota(ByRef dAliquota As Double, ByVal objFilialEmpresa As AdmFiliais) As Long
    If objFilialEmpresa.iCOFINSNaoCumulativo = 1 Then
        dAliquota = 0.076
    Else
        dAliquota = 0.03
    End If
    COFINS_Aliquota = SUCESSO
End Function

Function SpedFiscal_0190(ByVal sUM As String, ByVal iClasseUM As Integer, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro0190 As ClassSpedRegistro, ByVal objSpedRegistro0220 As ClassSpedRegistro, ByVal colUM As Collection, sUMResult As String, ByVal sProduto As String, ByVal lSeqPai As Long) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 6) As Long
Dim objUM As ClassUnidadeDeMedida
Dim iSufixo As Integer
Dim objUM1 As New ClassUnidadeDeMedida
Dim tProduto As typeProduto
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lSeq As Long
Dim objUM2 As New ClassUnidadeDeMedida

On Error GoTo Erro_SpedFiscal_0190

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203482
    Next

    sUMResult = ""
    iSufixo = -1

    For Each objUM In colUM
        If UCase(objUM.sSiglaUMBase) = UCase(sUM) Then
        
            If objUM.iClasse = iClasseUM Then
                sUMResult = objUM.sSigla
'                Set objUM1 = objUM
                Exit For
            Else
                If objUM.iEditavel > iSufixo Then
                    iSufixo = objUM.iEditavel
                End If
            End If
            
        End If
            
    Next
    
    If sUMResult = "" Then
    
        objUM1.sSigla = sUM
        objUM1.iClasse = iClasseUM
    
'        MsgBox ("sSigla = " & sUM)
'        MsgBox ("iClasse = " & iClasseUM)
    
    
        lErro = CF("UM_Le", objUM1)
        If lErro <> SUCESSO Then gError 203483
     
        Set objUM = New ClassUnidadeDeMedida
        
        objUM.iEditavel = iSufixo + 1
        objUM.sSiglaUMBase = sUM
        objUM.sSigla = sUM & IIf(objUM.iEditavel > 0, "_" & CStr(objUM.iEditavel), "")
        objUM.sNome = objUM1.sNome
        objUM.iClasse = iClasseUM
        objUM.dQuantidade = objUM1.dQuantidade
'        MsgBox ("objUM1.dQuantidade = " & objUM1.dQuantidade)
        sUMResult = objUM.sSigla
        
        colUM.Add objUM
        
        objSpedFiscal.sTipo = "0190"
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistro0190.iCount = objSpedRegistro0190.iCount + 1
        
        If UCase(objUM.sNome) = UCase(sUMResult) Then objUM.sNome = "Descricao " & objUM.sNome
        
        objSpedFiscal.sRegistro = "|0190|" & sUMResult & "|" & objUM.sNome & "|"
        
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203484
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203485
        
        
        'verifica se  o fator de conversao da unidade ja esta cadastrado para o produto em questao
        lErro = Comando_Executar(alComando(2), "SELECT Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0220' AND SUBSTRING(Registro, 7,7) LIKE ?", lSeq, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, sUMResult & "|%")
        If lErro <> AD_SQL_SUCESSO Then gError 203486
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203487
            
        If lErro = AD_SQL_SEM_DADOS Then
            
            tProduto.sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
            
            'le  o produto
            lErro = Comando_Executar(alComando(6), "SELECT SiglaUMEstoque, ClasseUM FROM Produtos WHERE Codigo = ?", tProduto.sSiglaUMEstoque, tProduto.iClasseUM, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 203488
        
            lErro = Comando_BuscarPrimeiro(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203489
        
    '        MsgBox ("lErro = " & lErro)
        
            dFator = 0
        
            'se a sigla de estoque for diferente da sigla em questao ==> tem que gravar o registro 0220 com o fator de conversao
'            MsgBox ("tProduto.sSiglaUMEstoque = " & tProduto.sSiglaUMEstoque)
'            MsgBox ("sUM = " & sUM)
            
            If tProduto.sSiglaUMEstoque <> sUM Then
            
                    
                objUM2.sSigla = tProduto.sSiglaUMEstoque
                objUM2.iClasse = tProduto.iClasseUM
            
                lErro = CF("UM_Le", objUM2)
                If lErro <> SUCESSO Then gError 203483
            
'                MsgBox ("dFator = 0 objUM1.dQuantidade = " & objUM1.dQuantidade)
'                MsgBox ("dFator = 0 objUM2.dQuantidade = " & objUM2.dQuantidade)
            
                dFator = objUM1.dQuantidade / objUM2.dQuantidade
                
'                MsgBox ("dFator = 0 dFator = " & dFator)
                
                
                objSpedFiscal.sTipo = "0220"
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistro0220.iCount = objSpedRegistro0220.iCount + 1
                
                objSpedFiscal.sRegistro = "|0220|" & sUMResult & "|" & Format(dFator, "########0.0#####") & "|"
                
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203490
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                If lErro <> AD_SQL_SUCESSO Then gError 203491
                
                
            End If
            
        End If
            
    End If
    
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next


    SpedFiscal_0190 = SUCESSO

    Exit Function
    
    
Erro_SpedFiscal_0190:

    SpedFiscal_0190 = gErr

    Select Case gErr
    
        Case 203482
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203483
            Call Rotina_Erro(vbOKOnly, "ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA", gErr, iClasseUM, sUM)
    
        Case 203484, 203490
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203485, 203491
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case 203486, 203487
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISCAL", gErr)
    
        Case 203488, 203489
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203492)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

'Function SpedFiscal_0460(ByVal sObs As String, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro0460 As ClassSpedRegistro, sCodObs As String) As Long
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim alComando(1 To 6) As Long
'
'On Error GoTo Erro_SpedFiscal_0460
'
'    'Abertura de Comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 203493
'    Next
'
'    sCodObs = String(6, 0)
'
'    'verifica se a obs ja esta cadastrado nos registros  (0460)
'    lErro = Comando_Executar(alComando(6), "SELECT sCodObs FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0460' AND SUBSTRING(Registro,14, 500) = ?", sCodObs, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, sObs)
'    If lErro <> AD_SQL_SUCESSO Then gError 203494
'
'    lErro = Comando_BuscarPrimeiro(alComando(6))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203495
'
'    If lErro = AD_SQL_SEM_DADOS Then
'
'        sCodObs = String(6, 0)
'
'        lErro = Comando_Executar(alComando(6), "SELECT Substring(Registro,7,6) FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0460' ORDER BY SUBSTRING(Registro,7,6)", sCodObs, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, sObs)
'        If lErro <> AD_SQL_SUCESSO Then gError 203496
'
'        lErro = Comando_BuscarPrimeiro(alComando(6))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203497
'
'        If lErro = AD_SQL_SEM_DADOS Then
'
'            sCodObs = "000001"
'
'        Else
'
'            sCodObs = Format(CLng(sCodObs) + 1, "000000")
'
'        End If
'
'        objSpedFiscal.sTipo = "0460"
'
'        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
'        objSpedRegistro0460.iCount = objSpedRegistro0460.iCount + 1
'
'        objSpedFiscal.sRegistro = "|0460|" & sCodObs & "|" & sObs & "|"
'
'        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203498
'
'        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
'        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
'        If lErro <> AD_SQL_SUCESSO Then gError 203499
'
'    End If
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'         Call Comando_Fechar(alComando(iIndice))
'    Next
'
'
'    SpedFiscal_0460 = SUCESSO
'
'    Exit Function
'
'
'Erro_SpedFiscal_0460:
'
'    SpedFiscal_0460 = gErr
'
'   Select Case gErr
'
'        Case 203493
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 203494 To 203497
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISCAL", gErr)
'
'        Case 203498
'            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
'
'        Case 203499
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203500)
'
'    End Select
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'         Call Comando_Fechar(alComando(iIndice))
'    Next
'
'End Function

Private Function SpedFiscal_C425(ByVal iCodigoECF As Integer, ByVal iFilialEmpresa As Integer, ByRef tArqMestre As typeLojaArqFisMestre, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro0200 As ClassSpedRegistro, ByVal objSpedRegistro0205 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal objSpedRegistro0190 As ClassSpedRegistro, ByVal objSpedRegistro0220 As ClassSpedRegistro, ByVal colUM As Collection, ByVal objSpedRegistroC425 As ClassSpedRegistro, ByVal sSituacaoTrib As String, ByRef dValorTotal As Double, ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim tItemCFiscal As typeItemCupomFiscal
Dim dValor As Double
Dim tProduto As typeProduto
Dim lSeqC420 As Long
Dim lSeq0200 As Long
Dim sUMResult As String
Dim dValorProduto As Double
Dim dValorDesconto As Double
Dim dValorAcrescimo As Double
Dim sProduto As String
Dim sUM As String
Dim iClasseUM As Integer
Dim dValorC425 As Double
Dim dQuantC425 As Double

On Error GoTo Erro_SpedFiscal_C425


    If gsPerfil <> "A" And gsPerfil <> "C" Then

        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 203501
        Next
    
    
        lSeqC420 = objSpedFiscal.lSeq
    
        tItemCFiscal.sUnidadeMed = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        tItemCFiscal.sProduto = String(STRING_PRODUTO, 0)
        
        
'        lErro = Comando_Executar(alComando(2), "SELECT Produto, SUM(Quantidade) , MAX(UnidadeMed), SUM((PrecoUnitario * Quantidade) - I.ValorDesconto), MAX(ClasseUM) FROM ItensCupomFiscal AS I, CupomFiscal AS C, Produtos As P WHERE P.Codigo = I.Produto AND I.NumIntCupom = C.NumIntDoc AND ECF = ? And C.FilialEmpresa = ? And DataEmissao = ? AND SituacaoTrib = ? GROUP BY Produto", _
'        tItemCFiscal.sProduto, tItemCFiscal.dQuantidade, tItemCFiscal.sUnidadeMed, dValor, tProduto.iClasseUM, iCodigoECF, iFilialEmpresa, tArqMestre.dtData, sSituacaoTrib)
'        If lErro <> AD_SQL_SUCESSO Then gError 203502
        
        
        lErro = Comando_Executar(alComando(2), "SELECT Produto, Quantidade , UnidadeMed, (PrecoUnitario * Quantidade) - I.ValorDesconto, ClasseUM, C.ValorProdutos, C.ValorDesconto, C.ValorAcrescimo  FROM ItensCupomFiscal AS I, CupomFiscal AS C, Produtos As P WHERE C.NFeChaveAcesso = '' AND P.Codigo = I.Produto AND I.NumIntCupom = C.NumIntDoc AND ECF = ? And C.FilialEmpresa = ? And DataEmissao = ? AND SituacaoTrib = ? AND C.Status <> 7 AND I.Status <> 7 AND C.Tipo = 1 ORDER  BY Produto", _
        tItemCFiscal.sProduto, tItemCFiscal.dQuantidade, tItemCFiscal.sUnidadeMed, dValor, tProduto.iClasseUM, dValorProduto, dValorDesconto, dValorAcrescimo, iCodigoECF, iFilialEmpresa, tArqMestre.dtData, sSituacaoTrib)
        If lErro <> AD_SQL_SUCESSO Then gError 203502
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203503
        
        If lErro = AD_SQL_SUCESSO Then
        
            sProduto = tItemCFiscal.sProduto
            sUM = tItemCFiscal.sUnidadeMed
            iClasseUM = tProduto.iClasseUM
            
        Else
        
            sProduto = ""
            sUM = ""
            iClasseUM = 0
            
        End If
        
        Do While lErro = AD_SQL_SUCESSO
        
            If sProduto <> tItemCFiscal.sProduto Then
                
    '            lErro = SpedFiscal_0200(tItemCFiscal.sProduto, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, lSeq0200)
    '            If lErro <> SUCESSO Then gError 203504
    
                lErro = SpedFiscal_0200(sProduto, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, lSeq0200, objFisSpedAux)
                If lErro <> SUCESSO Then gError 203504
    
    
    '            lErro = SpedFiscal_0190(tItemCFiscal.sUnidadeMed, tProduto.iClasseUM, objSpedFiscal, objSpedRegistro0190, objSpedRegistro0220, colUM, sUMResult, tItemCFiscal.sProduto, lSeq0200)
    '            If lErro <> SUCESSO Then gError 203505
            
                lErro = SpedFiscal_0190(sUM, iClasseUM, objSpedFiscal, objSpedRegistro0190, objSpedRegistro0220, colUM, sUMResult, sProduto, lSeq0200)
                If lErro <> SUCESSO Then gError 203505
        
                objSpedFiscal.sTipo = "C425"
                objSpedRegistroC425.sTipo = objSpedFiscal.sTipo
            
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistroC425.iCount = objSpedRegistroC425.iCount + 1
    
                objSpedFiscal.sRegistro = "|C425|" & sProduto & "|" & dQuantC425 & "|" & sUMResult & "|" & IIf(dValorC425 = 0, 0, Format(dValorC425, "fixed")) & "|||"
    
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203506
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqC420)
                If lErro <> AD_SQL_SUCESSO Then gError 203507
        
                dValorTotal = dValorTotal + dValorC425
    
                sProduto = tItemCFiscal.sProduto
                sUM = tItemCFiscal.sUnidadeMed
                iClasseUM = tProduto.iClasseUM
                dValorC425 = 0
                dQuantC425 = 0
    
            End If
        
        
            dValorC425 = dValorC425 + dValor + Round(Round((dValorAcrescimo - dValorDesconto) / dValorProduto, 8) * dValor, 2)
            dQuantC425 = dQuantC425 + tItemCFiscal.dQuantidade
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203508
    
        Loop
    
        If sProduto <> "" Then
            
            lErro = SpedFiscal_0200(sProduto, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, lSeq0200, objFisSpedAux)
            If lErro <> SUCESSO Then gError 210935


            lErro = SpedFiscal_0190(sUM, iClasseUM, objSpedFiscal, objSpedRegistro0190, objSpedRegistro0220, colUM, sUMResult, sProduto, lSeq0200)
            If lErro <> SUCESSO Then gError 210936
    
            objSpedFiscal.sTipo = "C425"
            objSpedRegistroC425.sTipo = objSpedFiscal.sTipo
        
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            objSpedRegistroC425.iCount = objSpedRegistroC425.iCount + 1

            objSpedFiscal.sRegistro = "|C425|" & sProduto & "|" & dQuantC425 & "|" & sUMResult & "|" & IIf(dValorC425 = 0, 0, Format(dValorC425, "fixed")) & "|||"

            If Len(objSpedFiscal.sRegistro) > 500 Then gError 210937
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqC420)
            If lErro <> AD_SQL_SUCESSO Then gError 210938
    
            dValorTotal = dValorTotal + dValorC425

        End If
    
    
        For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
        Next

    End If


    SpedFiscal_C425 = SUCESSO

    Exit Function
    
    
Erro_SpedFiscal_C425:

    SpedFiscal_C425 = gErr

   Select Case gErr
    
        Case 203501
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203502, 203503, 203508
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCUPOMFISCAL", gErr)
    
        Case 203504, 203505, 210935, 210936
    
        Case 203506, 210937
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203507, 210938
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203509)
     
    End Select
 
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_C460(ByVal iCodigoECF As Integer, ByVal iFilialEmpresa As Integer, ByRef tArqMestre As typeLojaArqFisMestre, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro0200 As ClassSpedRegistro, ByVal objSpedRegistro0205 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal objSpedRegistro0190 As ClassSpedRegistro, ByVal objSpedRegistro0220 As ClassSpedRegistro, ByVal colUM As Collection, ByVal objSpedRegistroC460 As ClassSpedRegistro, ByVal objSpedRegistroC470 As ClassSpedRegistro, lSeqReducaoZ As Long, colSpedC490 As Collection, dValDebICMS_objSpedC490 As Double, ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 3) As Long
Dim tProduto As typeProduto
Dim lSeqC460 As Long
Dim tCupomFiscal As typeCupomFiscal
Dim tItemCFiscal As typeItemCupomFiscal
Dim iTipoTribCST As Integer
Dim sUMResult As String
Dim sTipoTribCST As String
Dim sCFOP As String
Dim objSpedC490 As ClassSpedC490
Dim iAchou As Integer
Dim lSeq0200 As Long
Dim dValor As Double
Dim dDescAcresc As Double


On Error GoTo Erro_SpedFiscal_C460



    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203510
    Next

    tCupomFiscal.sCPFCGC = String(STRING_CGC, 0)
    
    lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Status, Numero, DataEmissao, ValorTotal, CPFCGC, ValorProdutos, ValorDesconto, ValorAcrescimo  FROM CupomFiscal  WHERE NFeChaveAcesso = '' AND ECF = ? And FilialEmpresa = ? And DataEmissao = ? AND Tipo = 1 ORDER BY Numero", _
        tCupomFiscal.lNumIntDoc, tCupomFiscal.iStatus, tCupomFiscal.lNumero, tCupomFiscal.dtDataEmissao, tCupomFiscal.dValorTotal, tCupomFiscal.sCPFCGC, tCupomFiscal.dValorProdutos, tCupomFiscal.dValorDesconto, tCupomFiscal.dValorAcrescimo, iCodigoECF, iFilialEmpresa, tArqMestre.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 203511
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203512
    
    Do While lErro = AD_SQL_SUCESSO
    
         If gsPerfil <> "B" And gsPerfil <> "C" Then
   
            objSpedFiscal.sTipo = "C460"
            objSpedRegistroC460.sTipo = objSpedFiscal.sTipo
        
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            objSpedRegistroC460.iCount = objSpedRegistroC460.iCount + 1

            If tCupomFiscal.iStatus <> STATUS_CANCELADO Then
                objSpedFiscal.sRegistro = "|C460|2D|00|" & tCupomFiscal.lNumero & "|" & Format(tCupomFiscal.dtDataEmissao, "ddmmyyyy") & "|" & IIf(tCupomFiscal.dValorTotal = 0, 0, Format(tCupomFiscal.dValorTotal, "fixed")) & "|||"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tCupomFiscal.sCPFCGC & "||"
            Else
                objSpedFiscal.sRegistro = "|C460|2D|02|" & tCupomFiscal.lNumero & "|||||||"
            End If

            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203513
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
            If lErro <> AD_SQL_SUCESSO Then gError 203514
    
            lSeqC460 = objSpedFiscal.lSeq
    
        End If
    
        '********** registro C470 **************************************************************
        
        If tCupomFiscal.iStatus <> STATUS_CANCELADO Then
        
            tItemCFiscal.sProduto = String(STRING_PRODUTO, 0)
            tItemCFiscal.sUnidadeMed = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
            tItemCFiscal.sSituacaoTrib = String(STRING_ITEMCFISCAL_SITUACAOTRIB, 0)
        
            lErro = Comando_Executar(alComando(3), "SELECT Produto, Quantidade, UnidadeMed, PrecoUnitario, TipoTribCST, AliquotaICMS, ClasseUM, ValorDesconto, Compras, Natureza, OrigemMercadoria, SituacaoTrib  FROM ItensCupomFiscal AS I INNER JOIN Produtos AS P ON I.Produto = P.Codigo  LEFT OUTER JOIN TiposTribICMS AS T ON Left(I.SituacaoTrib,1) = T.SituacaoTribECF WHERE  NumIntCupom = ? AND I.Status <> 7 ORDER BY Item", _
            tItemCFiscal.sProduto, tItemCFiscal.dQuantidade, tItemCFiscal.sUnidadeMed, tItemCFiscal.dPrecoUnitario, iTipoTribCST, tItemCFiscal.dAliquotaICMS, tProduto.iClasseUM, tItemCFiscal.dValorDesconto, tProduto.iCompras, tProduto.iNatureza, tProduto.iOrigemMercadoria, tItemCFiscal.sSituacaoTrib, tCupomFiscal.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 203515
            
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203516
            
            Do While lErro = AD_SQL_SUCESSO
                
'                lErro = SpedFiscal_0200(tItemCFiscal.sProduto, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, lSeq0200)
'                If lErro <> SUCESSO Then gError 203517
'
'                lErro = SpedFiscal_0190(tItemCFiscal.sUnidadeMed, tProduto.iClasseUM, objSpedFiscal, objSpedRegistro0190, objSpedRegistro0220, colUM, sUMResult, tItemCFiscal.sProduto, lSeq0200)
'                If lErro <> SUCESSO Then gError 203518
            
                If gsPerfil <> "B" And gsPerfil <> "C" Then
        
            
                    objSpedFiscal.sTipo = "C470"
                    objSpedRegistroC470.sTipo = objSpedFiscal.sTipo
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistroC470.iCount = objSpedRegistroC470.iCount + 1
                    
                End If
            
            
                If left(tItemCFiscal.sSituacaoTrib, 1) = "S" Then
                    sCFOP = "5949"
                    '****** os servicos vao estrar com CST = nao tributado ******
                    iTipoTribCST = 41
                ElseIf tProduto.iCompras = PRODUTO_PRODUZIVEL Then
                    sCFOP = "5101"
                Else
                    sCFOP = "5102"
                End If
        
                sTipoTribCST = tProduto.iOrigemMercadoria & iTipoTribCST
            
            
                dValor = (tItemCFiscal.dQuantidade * tItemCFiscal.dPrecoUnitario) - tItemCFiscal.dValorDesconto
                dDescAcresc = Round(Round((tCupomFiscal.dValorAcrescimo - tCupomFiscal.dValorDesconto) / tCupomFiscal.dValorProdutos, 8) * dValor, 2)
                
                dValor = dValor + dDescAcresc
            
                If gsPerfil <> "B" And gsPerfil <> "C" Then
            
            
                    If tCupomFiscal.iStatus <> STATUS_CANCELADO Then
                    
                    
                        lErro = SpedFiscal_0200(tItemCFiscal.sProduto, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, lSeq0200, objFisSpedAux)
                        If lErro <> SUCESSO Then gError 203517
        
                        lErro = SpedFiscal_0190(tItemCFiscal.sUnidadeMed, tProduto.iClasseUM, objSpedFiscal, objSpedRegistro0190, objSpedRegistro0220, colUM, sUMResult, tItemCFiscal.sProduto, lSeq0200)
                        If lErro <> SUCESSO Then gError 203518
                    
                        objSpedFiscal.sRegistro = "|C470|" & SFF_Texto(tItemCFiscal.sProduto) & "|" & tItemCFiscal.dQuantidade & "||" & sUMResult & "|" & IIf(dValor = 0, 0, Format(dValor, "fixed")) & "|"
                        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & sTipoTribCST & "|" & sCFOP & "|" & IIf(tProduto.iNatureza <> NATUREZA_PROD_SERVICO, Format(tItemCFiscal.dAliquotaICMS * 100, "fixed"), "") & "|||"
                    End If
            
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203519
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqC460)
                    If lErro <> AD_SQL_SUCESSO Then gError 203520
        
                End If
        
                iAchou = 0
        
                For Each objSpedC490 In colSpedC490
                
                    If objSpedC490.iCSTICMS = StrParaInt(sTipoTribCST) And objSpedC490.sCFOP = sCFOP And _
                         objSpedC490.dAliquotaICMS = IIf(iTipoTribCST = 30 Or iTipoTribCST = 40 Or iTipoTribCST = 41 Or iTipoTribCST = 50 Or iTipoTribCST = 60, 0, tItemCFiscal.dAliquotaICMS) Then
                        iAchou = 1
                        
                        If iTipoTribCST <> 30 And iTipoTribCST <> 40 And iTipoTribCST <> 41 And iTipoTribCST <> 50 And iTipoTribCST <> 60 Then
                            objSpedC490.dValorBCICMS = objSpedC490.dValorBCICMS + dValor
                            objSpedC490.dValorICMS = objSpedC490.dValorICMS + Round(dValor * tItemCFiscal.dAliquotaICMS, 2)
                            dValDebICMS_objSpedC490 = dValDebICMS_objSpedC490 + Round(dValor * tItemCFiscal.dAliquotaICMS, 2)
                        End If
                        
                        objSpedC490.dValorOper = objSpedC490.dValorOper + dValor

                        Exit For
                    End If
                
                
                Next
        
                If iAchou = 0 Then
                    Set objSpedC490 = New ClassSpedC490
                    
                    objSpedC490.iCSTICMS = StrParaInt(sTipoTribCST)
                    objSpedC490.sCFOP = sCFOP
                    
                    If iTipoTribCST = 30 Or iTipoTribCST = 40 Or iTipoTribCST = 41 Or iTipoTribCST = 50 Or iTipoTribCST = 60 Then
                        objSpedC490.dAliquotaICMS = 0
                        objSpedC490.dValorBCICMS = 0
                        objSpedC490.dValorICMS = 0
                    Else
                        objSpedC490.dAliquotaICMS = tItemCFiscal.dAliquotaICMS
                        objSpedC490.dValorBCICMS = dValor
                        objSpedC490.dValorICMS = Round(dValor * tItemCFiscal.dAliquotaICMS, 2)
                        dValDebICMS_objSpedC490 = dValDebICMS_objSpedC490 + Round(dValor * tItemCFiscal.dAliquotaICMS, 2)
                        
                    End If
                    objSpedC490.dValorOper = dValor
                    
                    colSpedC490.Add objSpedC490
                    
                End If
        
                'soma os valores de ICMS dos registros D405 que estiverem com data de reducao dentro do periodo para poder colocar
                'posteriormente no registro E110
        
                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203521
        
            Loop
        
        End If
        
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203522
        
    Loop
        
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C460 = SUCESSO

    Exit Function
    
    
Erro_SpedFiscal_C460:

    SpedFiscal_C460 = gErr

   Select Case gErr
    
        Case 203510
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203511, 203512, 203522
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUPOMFISCAL", gErr)
    
        Case 203513, 203519
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203514, 203520
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case 203515, 203516, 203521
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCUPOMFISCAL", gErr)
        
        
        Case 203517, 203518
    
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203523)
     
    End Select
 
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_C490(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC490 As ClassSpedRegistro, lSeqReducaoZ As Long, colSpedC490 As Collection) As Long

Dim lErro As Long
Dim objSpedC490 As ClassSpedC490
Dim iIndice As Integer
Dim alComando(1 To 1) As Long


On Error GoTo Erro_SpedFiscal_C490

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203524
    Next

    For Each objSpedC490 In colSpedC490
    
        objSpedFiscal.sTipo = "C490"
        objSpedRegistroC490.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC490.iCount = objSpedRegistroC490.iCount + 1

        objSpedFiscal.sRegistro = "|C490|" & Format(objSpedC490.iCSTICMS, "000") & "|" & objSpedC490.sCFOP & "|" & IIf(objSpedC490.dAliquotaICMS = 0, 0, Format(objSpedC490.dAliquotaICMS * 100, "fixed")) & "|" & IIf(objSpedC490.dValorOper = 0, 0, Format(objSpedC490.dValorOper, "fixed")) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC490.dValorBCICMS = 0, 0, Format(objSpedC490.dValorBCICMS, "fixed")) & "|" & IIf(objSpedC490.dValorICMS = 0, 0, Format(objSpedC490.dValorICMS, "fixed")) & "||"
        
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203525
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 203526
    
    Next

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C490 = SUCESSO

    Exit Function
    
    
Erro_SpedFiscal_C490:

   SpedFiscal_C490 = gErr

   Select Case gErr
    
        Case 203524
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203525
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203526
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203527)
     
    End Select
 
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
 

End Function

Function SpedFiscal_C500(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro0150 As ClassSpedRegistro, ByVal objSpedRegistro0175 As ClassSpedRegistro, ByVal objSpedRegistroC500 As ClassSpedRegistro, ByVal objSpedRegistroC590 As ClassSpedRegistro, dValCredICMS_objSpedC590 As Double, ByVal dtDataIni As Date, ByVal dtDataFim As Date, colSpedC590UF As Collection, ByVal objFilialEmpresa As AdmFiliais) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer
Dim alComando(1 To 27) As Long
Dim tNFiscal As typeNFiscal
Dim tTipoDocInfo As typeTipoDocInfo
Dim tTributacaoNF As typeTributacaoNF
Dim lSeqPaiNF As Long
Dim tItemNF As typeItemNF
Dim objSpedC590 As ClassSpedC590
Dim colSpedC590 As Collection
Dim iCodSit As Integer
Dim sData As String
Dim sCodParticipante As String
Dim tTribItemNF As typeTribItemNF
Dim iTipoTribCST_ICMS As Integer
Dim tProduto As typeProduto
'Dim sData1 As String
Dim iTipoRegistro As Integer
Dim dTotalRateioNF As Double
Dim objSpedC590UF As ClassSpedC590UF
Dim iTipoTribCST_ICMSAux As Integer
Dim iIENaoContrib As Integer, iIEIsento As Integer, iIndDest As Integer

On Error GoTo Erro_SpedFiscal_C500

    'sData1 = left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 4)

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203528
    Next

    tNFiscal.sSerie = String(STRING_SERIE, 0)
    tNFiscal.sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
    tNFiscal.sPlaca = String(STRING_NFISCAL_PLACA, 0)
    tNFiscal.sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
    tTipoDocInfo.sNomeTelaNFiscal = String(STRING_NOME, 0)

    'se for uma nota fiscal de entrada de energia ou gas
    lErro = Comando_Executar(alComando(4), "SELECT NumIntDoc, ModeloArqICMS, Tipo, Cliente, FilialCli, Fornecedor, FilialForn, Status, Complementar, Serie, NumNotaFiscal, DataEmissao, DataEntrada, DataSaida, ValorTotal, ClasseDocCPR, NumIntDocCPR, ValorDesconto, " & _
    "ValorProdutos, ValorFrete, FreteRespons, ValorSeguro, ValorOutrasDespesas, MensagemNota, FilialEntrega, CodTransportadora, Placa, PlacaUF, VolumeQuant, PesoLiq, PesoBruto, T.TipoMovtoEstoque, NomeTelaNFiscal From NFiscal AS N, TiposDocInfo AS T WHERE TipoNFiscal = Codigo AND (Tipo = 1 Or Tipo =3) AND DataEntrada >= ? AND DataEntrada <= ? AND (ModeloArqICMS = 06 OR ModeloArqICMS = 28) AND FilialEmpresa = ?", _
    tNFiscal.lNumIntDoc, tTipoDocInfo.iModeloArqICMS, tTipoDocInfo.iTipo, tNFiscal.lCliente, tNFiscal.iFilialCli, tNFiscal.lFornecedor, tNFiscal.iFilialForn, tNFiscal.iStatus, tTipoDocInfo.iComplementar, tNFiscal.sSerie, tNFiscal.lNumNotaFiscal, tNFiscal.dtDataEmissao, tNFiscal.dtDataEntrada, tNFiscal.dtDataSaida, tNFiscal.dValorTotal, _
    tNFiscal.iClasseDocCPR, tNFiscal.lNumIntDocCPR, tNFiscal.dValorDesconto, tNFiscal.dValorProdutos, tNFiscal.dValorFrete, tNFiscal.iFreteRespons, tNFiscal.dValorSeguro, tNFiscal.dValorOutrasDespesas, tNFiscal.sMensagemNota, tNFiscal.iFilialEntrega, tNFiscal.iCodTransportadora, tNFiscal.sPlaca, tNFiscal.sPlacaUF, _
    tNFiscal.lVolumeQuant, tNFiscal.dPesoLiq, tNFiscal.dPesoBruto, tTipoDocInfo.iTipoMovtoEstoque, tTipoDocInfo.sNomeTelaNFiscal, dtDataIni, dtDataFim, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203529
    
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203530
    
    Do While lErro = AD_SQL_SUCESSO
    
        dTotalRateioNF = 0
    
        'se é uma nota de cliente
        If tNFiscal.lCliente > 0 Then
        
            lErro = SpedFiscal_0150_Cliente(tNFiscal.lCliente, tNFiscal.iFilialCli, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim, iIEIsento, iIENaoContrib)
            If lErro <> SUCESSO Then gError 203531
        
        Else
    
    
            lErro = SpedFiscal_0150_Fornecedor(tNFiscal.lFornecedor, tNFiscal.iFilialForn, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim, iIEIsento, iIENaoContrib)
            If lErro <> SUCESSO Then gError 203532
    
        End If
    
        lErro = NF_Obter_Situacao(alComando(27), tNFiscal, tTipoDocInfo, iCodSit)
        If lErro <> SUCESSO Then gError 203532
            
'        If tTipoDocInfo.iComplementar = 1 Then
'
'            iCodSit = 6
'
'        Else
'
'            iCodSit = 0
'
'        End If
        
        lErro = Retorna_UF_Subst_Tributaria(objFilialEmpresa, tNFiscal, tTipoDocInfo)
        If lErro <> SUCESSO Then gError 204931
        
        'If tNFiscal.iStatus = STATUS_CANCELADO Then iCodSit = 2

        sData = Format(tNFiscal.dtDataEntrada, "ddmmyyyy")
                        
        'pesquisa a tributacaoNF
        lErro = Comando_Executar(alComando(12), "SELECT ICMSBase, ICMSValor, ICMSSubstBase, ICMSSubstValor, IPIValor, PISCredito, COFINSCredito, PISRetido, COFINSRetido, ISSAliquota, ICMSCredito FROM TributacaoNF WHERE NumIntDoc = ?", _
        tTributacaoNF.dICMSBase, tTributacaoNF.dICMSValor, tTributacaoNF.dICMSSubstBase, tTributacaoNF.dICMSSubstValor, tTributacaoNF.dIPIValor, tTributacaoNF.dPISCredito, tTributacaoNF.dCOFINSCredito, tTributacaoNF.dPISRetido, tTributacaoNF.dCOFINSRetido, tTributacaoNF.dISSAliquota, tTributacaoNF.dICMSCredito, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203533
    
        lErro = Comando_BuscarPrimeiro(alComando(12))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203534

        objSpedFiscal.sTipo = "C500"
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        lSeqPaiNF = objSpedFiscal.lSeq
        
        objSpedRegistroC500.sTipo = objSpedFiscal.sTipo
        objSpedRegistroC500.iCount = objSpedRegistroC500.iCount + 1
        
        'se a nota nao for cancelada
        If iCodSit <> 2 And iCodSit <> 4 Then
        
            objSpedFiscal.sRegistro = "|C500|" & IIf(tTipoDocInfo.iTipo = 2, 1, 0) & "|" & IIf(tTipoDocInfo.iTipo = 3, 1, 0) & "|" & sCodParticipante & "|" & Format(tTipoDocInfo.iModeloArqICMS, "00") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(iCodSit, "00") & "|" & Ajusta_Serie(Desconverte_Serie_Eletronica(tNFiscal.sSerie)) & "|||" & tNFiscal.lNumNotaFiscal & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tNFiscal.dtDataEmissao, "ddmmyyyy") & "|" & sData & "|" & Format(tNFiscal.dValorTotal, "Fixed") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tNFiscal.dValorDesconto = 0, 0, Format(tNFiscal.dValorDesconto, "Fixed")) & "|" & Format(tNFiscal.dValorProdutos, "Fixed") & "|||"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tNFiscal.dValorOutrasDespesas = 0, 0, Format(tNFiscal.dValorOutrasDespesas, "Fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSBase = 0, 0, Format(tTributacaoNF.dICMSBase, "Fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSCredito = 0, 0, Format(tTributacaoNF.dICMSCredito, "Fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSSubstBase = 0, 0, Format(tTributacaoNF.dICMSSubstBase, "Fixed")) & "|||"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSSubstValor = 0, 0, Format(tTributacaoNF.dICMSSubstValor, "Fixed")) & "||||"

            If giVersao >= 14 Then
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tNFiscal.sChvNFe & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTipoDocInfo.iModeloArqICMS = 66, "1", "") & "|" '1 – Normal
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "|" 'CHV_DOCe_REF

                If tTipoDocInfo.iTipo = 2 Then 'Saída
                    If iIENaoContrib = MARCADO Then
                        iIndDest = 9 '9 – Não Contribuinte.
                    ElseIf iIEIsento = MARCADO Then
                        iIndDest = 2 '2 – Contribuinte Isento de Inscrição no Cadastro de Contribuintes do ICMS;
                    Else
                        iIndDest = 1 '1 – Contribuinte do ICMS;
                    End If
                Else
                    If Len(Trim(objFilialEmpresa.sInscricaoEstadual)) = 0 Then
                        iIndDest = 9 '9 – Não Contribuinte.
                    Else
                        iIndDest = 1 '1 – Contribuinte do ICMS;
                    End If
                End If
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & CStr(iIndDest) & "|" 'Indicador do Destinatário/Acessante:
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & gsCodIBGE & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "|"


            End If

        Else
        
            objSpedFiscal.sRegistro = "|C500|" & IIf(tTipoDocInfo.iTipo = 2, 1, 0) & "|" & IIf(tTipoDocInfo.iTipo = 3, 1, 0) & "||" & Format(tTipoDocInfo.iModeloArqICMS, "00") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(iCodSit, "00") & "|" & Ajusta_Serie(Desconverte_Serie_Eletronica(tNFiscal.sSerie)) & "|||" & tNFiscal.lNumNotaFiscal & "|" & Format(tNFiscal.dtDataEmissao, "ddmmyyyy") & "|||||||||||||||||"
        
            If giVersao >= 14 Then objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "||||||"

        End If
        
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203535
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203536
        
        

        '**************** registro C590 ********************************
        
        Set colSpedC590 = New Collection

        tItemNF.sProduto = String(STRING_PRODUTO, 0)
        tItemNF.sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
        tItemNF.sUnidadeMed = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
    
        lErro = Comando_Executar(alComando(25), "SELECT Item, Produto, DescricaoItem, Quantidade, UnidadeMed, ClasseUM, PrecoUnitario, ValorDesconto FROM ItensNFiscal, Produtos  WHERE Produto = Codigo AND NumIntNF  = ?", _
        tItemNF.iItem, tItemNF.sProduto, tItemNF.sDescricaoItem, tItemNF.dQuantidade, tItemNF.sUnidadeMed, tProduto.iClasseUM, tItemNF.dPrecoUnitario, tItemNF.dValorDesconto, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203537
    
        lErro = Comando_BuscarPrimeiro(alComando(25))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203538
    
    
        Do While lErro = AD_SQL_SUCESSO
    
            tTribItemNF.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    
            lErro = Comando_Executar(alComando(26), "SELECT OrigemMercadoria, TipoTribCST, NaturezaOP, ICMSBase, ICMSAliquota, ICMSValor, ICMSCredito, ICMSSubstBase, ICMSSubstAliquota, ICMSSubstValor, " & _
            "TipoTributacao, ICMSPercRedBase FROM TributacaoDocItem, tiposTribICMS AS ICMS, TiposTribIPI AS IPI  WHERE TipoDoc = 0 AND NumIntDoc  = ? AND Item = ? AND ICMS.Tipo = ICMSTipo AND IPI.Tipo = IPITipo", _
            tTribItemNF.iOrigemMercadoria, iTipoTribCST_ICMSAux, tTribItemNF.sNaturezaOp, tTribItemNF.dICMSBase, tTribItemNF.dICMSAliquota, tTribItemNF.dICMSValor, tTribItemNF.dICMSCredito, tTribItemNF.dICMSSubstBase, tTribItemNF.dICMSSubstAliquota, tTribItemNF.dICMSSubstValor, _
             tTribItemNF.iTipoTributacao, tTribItemNF.dICMSPercRedBase, tNFiscal.lNumIntDoc, tItemNF.iItem)
            If lErro <> AD_SQL_SUCESSO Then gError 203539
        
            lErro = Comando_BuscarPrimeiro(alComando(26))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203540
            
            iTipoTribCST_ICMS = iTipoTribCST_ICMSAux + (100 * tTribItemNF.iOrigemMercadoria)
    
            lErro = Rateia_Compl_ItemNF(tNFiscal, tItemNF, tTribItemNF)
            If lErro <> SUCESSO Then gError 204207
    
            iAchou = 0
    
            For Each objSpedC590 In colSpedC590
            
                If objSpedC590.iCSTICMS = iTipoTribCST_ICMS And objSpedC590.sCFOP = tTribItemNF.sNaturezaOp And objSpedC590.dAliquotaICMS = tTribItemNF.dICMSAliquota Then
                    objSpedC590.dValorOper = objSpedC590.dValorOper + (tItemNF.dQuantidade * tItemNF.dPrecoUnitario)
                    If tNFiscal.dValorProdutos > 0 Then
                        objSpedC590.dValorOper = objSpedC590.dValorOper + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
                        dTotalRateioNF = dTotalRateioNF + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'                        objSpedC590.dValorOper = objSpedC590.dValorOper - Round(((tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)

                    End If
                    objSpedC590.dValorBCICMS = objSpedC590.dValorBCICMS + tTribItemNF.dICMSBase
'mario                    objSpedC590.dValorICMS = objSpedC590.dValorICMS + tTribItemNF.dICMSValor
                    objSpedC590.dValorICMS = objSpedC590.dValorICMS + tTribItemNF.dICMSCredito
                    objSpedC590.dValorBCICMSST = objSpedC590.dValorBCICMSST + tTribItemNF.dICMSSubstBase
                    objSpedC590.dValorICMSST = objSpedC590.dValorICMSST + tTribItemNF.dICMSSubstValor
                    objSpedC590.dValorRedBC = objSpedC590.dValorRedBC + (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
                    objSpedC590.dValorBCICMS = objSpedC590.dValorBCICMS - objSpedC590.dValorRedBC
                    iAchou = 1
                    Exit For
                End If
                
            Next
    
            If iAchou = 0 Then
            
                Set objSpedC590 = New ClassSpedC590
                
                objSpedC590.iCSTICMS = iTipoTribCST_ICMS
                objSpedC590.sCFOP = tTribItemNF.sNaturezaOp
                objSpedC590.dAliquotaICMS = tTribItemNF.dICMSAliquota
                objSpedC590.dValorOper = (tItemNF.dQuantidade * tItemNF.dPrecoUnitario)
                If tNFiscal.dValorProdutos > 0 Then
                    objSpedC590.dValorOper = objSpedC590.dValorOper + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
                    dTotalRateioNF = dTotalRateioNF + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
                    
'                    objSpedC590.dValorOper = objSpedC590.dValorOper - Round(((tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)

                End If
                objSpedC590.dValorBCICMS = tTribItemNF.dICMSBase
'mario                objSpedC590.dValorICMS = tTribItemNF.dICMSValor
                objSpedC590.dValorICMS = tTribItemNF.dICMSCredito
                objSpedC590.dValorBCICMSST = tTribItemNF.dICMSSubstBase
                objSpedC590.dValorICMSST = tTribItemNF.dICMSSubstValor
                objSpedC590.dValorRedBC = (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
                objSpedC590.dValorBCICMS = objSpedC590.dValorBCICMS - objSpedC590.dValorRedBC
                
                colSpedC590.Add objSpedC590
                
            End If
            
            
            iAchou = 0
    
            For Each objSpedC590UF In colSpedC590UF
            
                If objSpedC590UF.sUF = tNFiscal.sUFSubstTrib And objSpedC590UF.sCFOP = tTribItemNF.sNaturezaOp Then
                    objSpedC590UF.dValorBCICMSST = objSpedC590UF.dValorBCICMSST + tTribItemNF.dICMSSubstBase
                    objSpedC590UF.dValorICMSST = objSpedC590UF.dValorICMSST + tTribItemNF.dICMSSubstValor
                    
                    iAchou = 1
                    Exit For
                End If
                
            Next
    
            If iAchou = 0 Then
            
                Set objSpedC590UF = New ClassSpedC590UF
                
                objSpedC590UF.sUF = tNFiscal.sUFSubstTrib
                objSpedC590UF.sCFOP = tTribItemNF.sNaturezaOp
                
                objSpedC590UF.dValorBCICMSST = tTribItemNF.dICMSSubstBase
                objSpedC590UF.dValorICMSST = tTribItemNF.dICMSSubstValor
                
                colSpedC590UF.Add objSpedC590UF
                
            End If
            
            
            
            'soma os valores de ICMS dos registros C500 que estiverem com data de saida/entrada dentro do periodo para poder colocar
            'posteriormente no registro E110
            'If CDate(sData1) >= dtDataIni And CDate(sData1) <= dtDataFim Then
            
            'Retirado IF porque:
            '1 - sData1 viria sempre com "//" e daria erro 13
            '2 - Já foi filtrado no select pela data de entrada (é só aquisição de energia/gás/água)
            
                dValCredICMS_objSpedC590 = dValCredICMS_objSpedC590 + tTribItemNF.dICMSCredito
            'End If
    
            lErro = Comando_BuscarProximo(alComando(25))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203541

        Loop
                
            
'        'complemento de NF
'        tTribItemNF.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'        lErro = Comando_Executar(alComando(5), "SELECT T.Tipo, TipoTribCST, NaturezaOP, ICMSBase, ICMSAliquota, ICMSValor, ICMSSubstBase, ICMSSubstValor, " & _
'        "IPIValor, ICMSPercRedBase FROM TributacaoComplNF AS T, tiposTribICMS AS ICMS  WHERE T.NumIntDoc  = ? AND ICMS.Tipo = ICMSTipo AND (T.Tipo = 1 Or T.Tipo=2 Or T.Tipo = 4)", _
'        iTipoRegistro, iTipoTribCST_ICMS, tTribItemNF.sNaturezaOp, tTribItemNF.dICMSBase, tTribItemNF.dICMSAliquota, tTribItemNF.dICMSValor, tTribItemNF.dICMSSubstBase, tTribItemNF.dICMSSubstValor, _
'        tTribItemNF.dIPIValor, tTribItemNF.dICMSPercRedBase, tNFiscal.lNumIntDoc)
'        If lErro <> AD_SQL_SUCESSO Then gError 204193
'
'        lErro = Comando_BuscarPrimeiro(alComando(5))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204194
'
'        Do While lErro = AD_SQL_SUCESSO
'
'
'            iAchou = 0
'
'            For Each objSpedC590 In colSpedC590
'
'                If objSpedC590.iCSTICMS = iTipoTribCST_ICMS And objSpedC590.sCFOP = tTribItemNF.sNaturezaOp And objSpedC590.dAliquotaICMS = tTribItemNF.dICMSAliquota Then
'                    If iTipoRegistro = ITEM_TIPO_FRETE Then
'                        objSpedC590.dValorOper = objSpedC590.dValorOper + tNFiscal.dValorFrete
'                    ElseIf iTipoRegistro = ITEM_TIPO_SEGURO Then
'                        objSpedC590.dValorOper = objSpedC590.dValorOper + tNFiscal.dValorSeguro
'                    ElseIf iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                        objSpedC590.dValorOper = objSpedC590.dValorOper + tNFiscal.dValorOutrasDespesas
''                    ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''                        objSpedC590.dValorOper = objSpedC590.dValorOper - tNFiscal.dValorDesconto
'                    End If
'                    If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'
'                        objSpedC590.dValorOper = objSpedC590.dValorOper + tTribItemNF.dICMSSubstValor + tTribItemNF.dIPIValor
'                        objSpedC590.dValorBCICMS = objSpedC590.dValorBCICMS + tTribItemNF.dICMSBase
'                        objSpedC590.dValorICMS = objSpedC590.dValorICMS + tTribItemNF.dICMSValor
'                        objSpedC590.dValorBCICMSST = objSpedC590.dValorBCICMSST + tTribItemNF.dICMSSubstBase
'                        objSpedC590.dValorICMSST = objSpedC590.dValorICMSST + tTribItemNF.dICMSSubstValor
'                        objSpedC590.dValorRedBC = objSpedC590.dValorRedBC + (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
''                    ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''                        objSpedC590.dValorOper = objSpedC590.dValorOper - tTribItemNF.dICMSSubstValor
''
''                        objSpedC590.dValorBCICMS = objSpedC590.dValorBCICMS - tTribItemNF.dICMSBase
''                        objSpedC590.dValorICMS = objSpedC590.dValorICMS - tTribItemNF.dICMSValor
''                        objSpedC590.dValorBCICMSST = objSpedC590.dValorBCICMSST - tTribItemNF.dICMSSubstBase
''                        objSpedC590.dValorICMSST = objSpedC590.dValorICMSST - tTribItemNF.dICMSSubstValor
''                        objSpedC590.dValorRedBC = objSpedC590.dValorRedBC - (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
'                    End If
'
'
'                    iAchou = 1
'                    Exit For
'                End If
'
'            Next
'
'            If iAchou = 0 Then
'
'                Set objSpedC590 = New ClassSpedC590
'
'                objSpedC590.iCSTICMS = iTipoTribCST_ICMS
'                objSpedC590.sCFOP = tTribItemNF.sNaturezaOp
'                objSpedC590.dAliquotaICMS = tTribItemNF.dICMSAliquota
'
'                If iTipoRegistro = ITEM_TIPO_FRETE Then
'                    objSpedC590.dValorOper = tNFiscal.dValorFrete
'                ElseIf iTipoRegistro = ITEM_TIPO_SEGURO Then
'                    objSpedC590.dValorOper = tNFiscal.dValorSeguro
'                ElseIf iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                    objSpedC590.dValorOper = tNFiscal.dValorOutrasDespesas
''                ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''                    objSpedC590.dValorOper = objSpedC590.dValorOper - tNFiscal.dValorDesconto
'                End If
'
'
'                If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                    objSpedC590.dValorOper = objSpedC590.dValorOper + tTribItemNF.dICMSSubstValor
'
'                    objSpedC590.dValorBCICMS = tTribItemNF.dICMSBase
'                    objSpedC590.dValorICMS = tTribItemNF.dICMSValor
'                    objSpedC590.dValorBCICMSST = tTribItemNF.dICMSSubstBase
'                    objSpedC590.dValorICMSST = tTribItemNF.dICMSSubstValor
'                    objSpedC590.dValorRedBC = (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
''                ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''                    objSpedC590.dValorOper = objSpedC590.dValorOper - tTribItemNF.dICMSSubstValor
''
''                    objSpedC590.dValorBCICMS = -tTribItemNF.dICMSBase
''                    objSpedC590.dValorICMS = -tTribItemNF.dICMSValor
''                    objSpedC590.dValorBCICMSST = -tTribItemNF.dICMSSubstBase
''                    objSpedC590.dValorICMSST = -tTribItemNF.dICMSSubstValor
''                    objSpedC590.dValorRedBC = -(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
'                End If
'
'                colSpedC590.Add objSpedC590
'
'            End If
'
'            If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'
'                'soma os valores de ICMS dos registros C100 de entrada que estiverem com data de entrada dentro do periodo para poder colocar
'                'posteriormente no registro E110
'                If CDate(sData1) >= dtDataIni And CDate(sData1) <= dtDataFim Then
'                    dValCredICMS_objSpedC590 = dValCredICMS_objSpedC590 + tTribItemNF.dICMSValor
'                End If
'
'
'
''            ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''                'soma os valores de ICMS dos registros C100 de entrada que estiverem com data de entrada dentro do periodo para poder colocar
''                'posteriormente no registro E110
''                If CDate(sData) >= dtDataIni And CDate(sData) <= dtDataFim Then
''                    dValCredICMS_objSpedC590 = dValCredICMS_objSpedC590 - tTribItemNF.dICMSValor
''                End If
'
'            End If
'
'            lErro = Comando_BuscarProximo(alComando(5))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204195
'
'        Loop
            
        If tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto <> dTotalRateioNF Then
    
            Set objSpedC590 = colSpedC590.Item(1)
    
            objSpedC590.dValorOper = objSpedC590.dValorOper + Round(tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto - dTotalRateioNF, 2)
    
        End If
            
            
        For Each objSpedC590 In colSpedC590
        
            If objSpedC590.dValorOper > 0 Or objSpedC590.dValorBCICMS > 0 Or objSpedC590.dValorICMS > 0 Or _
               objSpedC590.dValorBCICMSST > 0 Or objSpedC590.dValorICMSST > 0 Or objSpedC590.dValorRedBC > 0 Then
        
                objSpedFiscal.sTipo = "C590"
                objSpedRegistroC590.sTipo = objSpedFiscal.sTipo
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistroC590.iCount = objSpedRegistroC590.iCount + 1
        
                objSpedFiscal.sRegistro = "|C590|" & Format(objSpedC590.iCSTICMS, "000") & "|" & Format(objSpedC590.sCFOP, "0000") & "|" & Format(objSpedC590.dAliquotaICMS * 100, "fixed") & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC590.dValorOper <= 0, 0, Format(objSpedC590.dValorOper, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC590.dValorBCICMS <= 0, 0, Format(objSpedC590.dValorBCICMS, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC590.dValorICMS <= 0, 0, Format(objSpedC590.dValorICMS, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC590.dValorBCICMSST <= 0, 0, Format(objSpedC590.dValorBCICMSST, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC590.dValorICMSST <= 0, 0, Format(objSpedC590.dValorICMSST, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC590.dValorRedBC <= 0, 0, Format(objSpedC590.dValorRedBC, "fixed")) & "||"
                
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203542
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
                If lErro <> AD_SQL_SUCESSO Then gError 203543
            
            End If
            
        Next
            
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203544
        
    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C500 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C500:

   SpedFiscal_C500 = gErr

   Select Case gErr
    
        Case 203528
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203529, 203530, 203544
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
    
        Case 203531, 203532, 204207, 204931
    
        Case 203533, 203534
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAONF", gErr)
    
        Case 203535, 203542
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203536, 203543
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case 203537, 203538, 203541
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL2", gErr)
    
        Case 203539, 203540
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOITEMNF", gErr)
    
        Case 204193 To 204195
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLNF", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203545)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_0450(objMensagem As ClassMensagem, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro0450 As ClassSpedRegistro, ByRef tNFiscal As typeNFiscal, lSeqPai As Long, colMensagem As Collection) As Long

Dim lErro As Long
Dim iAchou As Integer
Dim iIndice As Integer
Dim alComando(1 To 1) As Long
Dim sTexto As String

On Error GoTo Erro_SpedFiscal_0450

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203546
    Next

    iAchou = 0
    sTexto = SFF_Texto(tNFiscal.sMensagemCorpoNota)
    If Len(Trim(sTexto)) > 0 And Len(Trim(tNFiscal.sMensagemNota)) > 0 Then sTexto = sTexto & " "
    sTexto = sTexto & SFF_Texto(tNFiscal.sMensagemNota)
    
    For Each objMensagem In colMensagem
        If objMensagem.sDescricao = sTexto Then
            iAchou = 1
            Exit For
        End If
    Next
    
    
    If iAchou = 0 And Len(Trim(sTexto)) > 0 Then
    
        '************ registro 0450 ********************************

        Set objMensagem = New ClassMensagem
        objMensagem.iCodigo = colMensagem.Count + 1
        objMensagem.sDescricao = sTexto
        colMensagem.Add objMensagem
        
        objSpedFiscal.sTipo = "0450"
        objSpedRegistro0450.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistro0450.iCount = objSpedRegistro0450.iCount + 1
        
        objSpedFiscal.sRegistro = "|0450|" & objMensagem.iCodigo & "|" & Desacentua_SpedFiscal(objMensagem.sDescricao) & "|"
        
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203547
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203548
        
        
    End If
        
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
        
    SpedFiscal_0450 = SUCESSO
    
    Exit Function
            
Erro_SpedFiscal_0450:

   SpedFiscal_0450 = gErr

   Select Case gErr
    
        Case 203546
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203547
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203548
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203549)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_D190(ByRef tNFiscal As typeNFiscal, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroD190 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date, dValCredICMS_objSpedD190 As Double, ByVal lSeqPai As Long) As Long

Dim tItemNF As typeItemNF
Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim colSpedC190 As New Collection 'é C190 mesmo pois estamos aproveitando a classe ClassSpedC190
Dim objSpedC190 As ClassSpedC190
Dim iAchou As Integer
Dim iIndice As Integer
Dim iTipoTribCST_ICMS As Integer
Dim iIPICSTEntrada As Integer
Dim iIPICSTSaida As Integer
Dim tProduto As typeProduto
Dim tTribItemNF As typeTribItemNF
Dim dTotalRateioNF As Double
Dim iTipoRegistro As Integer
Dim objTribItem As ClassTributacaoDocItem
Dim iTipoTribCST_ICMSAux As Integer
Dim dValorFCP As Double
Dim dValorFCPRet As Double
Dim dValorFCPST As Double

On Error GoTo Erro_SpedFiscal_D190

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203550
    Next
            
    dTotalRateioNF = 0

    tItemNF.sProduto = String(STRING_PRODUTO, 0)
    tItemNF.sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
    tItemNF.sUnidadeMed = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)

    lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Item, Produto, DescricaoItem, Quantidade, UnidadeMed, ClasseUM, PrecoUnitario, ValorDesconto FROM ItensNFiscal, Produtos  WHERE Produto = Codigo AND NumIntNF  = ?", _
    tItemNF.lNumIntDoc, tItemNF.iItem, tItemNF.sProduto, tItemNF.sDescricaoItem, tItemNF.dQuantidade, tItemNF.sUnidadeMed, tProduto.iClasseUM, tItemNF.dPrecoUnitario, tItemNF.dValorDesconto, tNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 203551

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203552

    Do While lErro = AD_SQL_SUCESSO

        tTribItemNF.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)

        lErro = Comando_Executar(alComando(3), "SELECT OrigemMercadoria, TipoTribCST, NaturezaOP, ICMSBase, ICMSAliquota, ICMSValor, ICMSCredito, ICMSSubstBase, ICMSSubstAliquota, ICMSSubstValor, IPI.CSTEntrada, IPI.CSTSaida, " & _
        "IPIBaseCalculo, IPIAliquota, IPIValor, PISCredito, TipoTributacao, ICMSPercRedBase, ICMSvFCP, ICMSvFCPST, ICMSvFCPSTRet FROM TributacaoDocItem, tiposTribICMS AS ICMS, TiposTribIPI AS IPI  WHERE TipoDoc = 0 AND NumIntDoc  = ? AND Item = ? AND ICMS.Tipo = ICMSTipo AND IPI.Tipo = IPITipo", _
        tTribItemNF.iOrigemMercadoria, iTipoTribCST_ICMSAux, tTribItemNF.sNaturezaOp, tTribItemNF.dICMSBase, tTribItemNF.dICMSAliquota, tTribItemNF.dICMSValor, tTribItemNF.dICMSCredito, tTribItemNF.dICMSSubstBase, tTribItemNF.dICMSSubstAliquota, tTribItemNF.dICMSSubstValor, _
        iIPICSTEntrada, iIPICSTSaida, tTribItemNF.dIPIBaseCalculo, tTribItemNF.dIPIAliquota, tTribItemNF.dIPIValor, tTribItemNF.dPISCredito, tTribItemNF.iTipoTributacao, tTribItemNF.dICMSPercRedBase, dValorFCP, dValorFCPST, dValorFCPRet, tNFiscal.lNumIntDoc, tItemNF.iItem)
        If lErro <> AD_SQL_SUCESSO Then gError 203553
    
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203554
        
        iTipoTribCST_ICMS = iTipoTribCST_ICMSAux
        
        Set objTribItem = New ClassTributacaoDocItem
        objTribItem.lNumIntDocItem = tItemNF.lNumIntDoc
        objTribItem.iTipoDoc = TIPODOC_TRIB_NF
        lErro = ItemNF_Le_BasePropria(objTribItem, iTipoTribCST_ICMS, iIPICSTEntrada)
        If lErro <> SUCESSO Then gError 204208
        
        iTipoTribCST_ICMS = iTipoTribCST_ICMS + (100 * tTribItemNF.iOrigemMercadoria)
        
        tTribItemNF.dICMSBase = objTribItem.dICMSBase
        tTribItemNF.dICMSAliquota = objTribItem.dICMSAliquota
        tTribItemNF.dICMSValor = objTribItem.dICMSValor
        tTribItemNF.dICMSCredito = objTribItem.dICMSCredito
        
        tTribItemNF.dICMSSubstBase = objTribItem.dICMSSubstBase
        tTribItemNF.dICMSSubstAliquota = objTribItem.dICMSSubstAliquota
        tTribItemNF.dICMSSubstValor = objTribItem.dICMSSubstValor
        tTribItemNF.dICMSSubstPercMVA = objTribItem.dICMSSubstPercMVA
        tTribItemNF.dICMSSubstPercRedBase = objTribItem.dICMSSubstPercRedBase

        tTribItemNF.dIPIBaseCalculo = objTribItem.dIPIBaseCalculo
        tTribItemNF.dIPIAliquota = objTribItem.dIPIAliquota
        tTribItemNF.dIPIValor = objTribItem.dIPIValor
        tTribItemNF.dIPICredito = objTribItem.dIPICredito

        lErro = Rateia_Compl_ItemNF(tNFiscal, tItemNF, tTribItemNF)
        If lErro <> SUCESSO Then gError 204208

        iAchou = 0

        For Each objSpedC190 In colSpedC190
        
            If objSpedC190.iCSTICMS = iTipoTribCST_ICMS And objSpedC190.sCFOP = tTribItemNF.sNaturezaOp And objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota Then
            
                objSpedC190.dValorOper = objSpedC190.dValorOper + (tItemNF.dQuantidade * tItemNF.dPrecoUnitario)
                If tNFiscal.dValorProdutos > 0 Then
                    objSpedC190.dValorOper = objSpedC190.dValorOper + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
                    dTotalRateioNF = dTotalRateioNF + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
                    'objSpedC190.dValorOper = objSpedC190.dValorOper - Round(((tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
                End If
                objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS + tTribItemNF.dICMSBase
'                objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dICMSValor
                objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dICMSCredito
                objSpedC190.dValorRedBC = objSpedC190.dValorRedBC + Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
                objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS - Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)  'objSpedC190.dValorRedBC
                objSpedC190.dValorFCP = objSpedC190.dValorFCP + dValorFCP
                objSpedC190.dValorFCPST = objSpedC190.dValorFCPST + dValorFCPST
                objSpedC190.dValorFCPRet = objSpedC190.dValorFCPRet + dValorFCPRet
                iAchou = 1
                Exit For
            End If
            
        Next

        If iAchou = 0 Then
        
            Set objSpedC190 = New ClassSpedC190
            
            objSpedC190.iCSTICMS = iTipoTribCST_ICMS
            objSpedC190.sCFOP = tTribItemNF.sNaturezaOp
            objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota
            objSpedC190.dValorOper = (tItemNF.dQuantidade * tItemNF.dPrecoUnitario)
            If tNFiscal.dValorProdutos > 0 Then
                objSpedC190.dValorOper = objSpedC190.dValorOper + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
                dTotalRateioNF = dTotalRateioNF + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'                objSpedC190.dValorOper = objSpedC190.dValorOper - Round(((tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)

            End If
            objSpedC190.dValorBCICMS = tTribItemNF.dICMSBase
'mario            objSpedC190.dValorICMS = tTribItemNF.dICMSValor
            objSpedC190.dValorICMS = tTribItemNF.dICMSCredito
            objSpedC190.dValorRedBC = Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
            objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS - Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase) 'objSpedC190.dValorRedBC
            objSpedC190.dValorFCP = dValorFCP
            objSpedC190.dValorFCPST = dValorFCPST
            objSpedC190.dValorFCPRet = dValorFCPRet
            
            colSpedC190.Add objSpedC190
            
        End If

        'soma os valores de ICMS dos registros D100 que estiverem com data de saida/entrada dentro do periodo para poder colocar
        'posteriormente no registro E110
        If tNFiscal.dtDataEntrada >= dtDataIni And tNFiscal.dtDataEntrada <= dtDataFim Then
'mario            dValCredICMS_objSpedD190 = dValCredICMS_objSpedD190 + tTribItemNF.dICMSValor
            dValCredICMS_objSpedD190 = dValCredICMS_objSpedD190 + tTribItemNF.dICMSCredito
        End If


        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203555

    Loop
    
    '********************** registro D190 *******************************
    If tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto <> dTotalRateioNF Then

        Set objSpedC190 = colSpedC190.Item(1)

        objSpedC190.dValorOper = objSpedC190.dValorOper + Round(tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto - dTotalRateioNF, 2)

    End If
    
'        'trata os complementos das notas fiscais
'
'        tTribItemNF.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'        lErro = Comando_Executar(alComando(5), "SELECT T.Tipo, TipoTribCST, NaturezaOP, ICMSBase, ICMSAliquota, ICMSValor, ICMSSubstBase, ICMSSubstValor, IPI.CSTEntrada, IPI.CSTSaida, " & _
'        "IPIValor, ICMSPercRedBase FROM TributacaoComplNF AS T, tiposTribICMS AS ICMS, TiposTribIPI AS IPI  WHERE T.NumIntDoc  = ? AND ICMS.Tipo = ICMSTipo AND IPI.Tipo = IPITipo AND (T.Tipo = 1 Or T.Tipo=2 Or T.Tipo = 4)", _
'        iTipoRegistro, iTipoTribCST_ICMS, tTribItemNF.sNaturezaOp, tTribItemNF.dICMSBase, tTribItemNF.dICMSAliquota, tTribItemNF.dICMSValor, tTribItemNF.dICMSSubstBase, tTribItemNF.dICMSSubstValor, _
'        iIPICSTEntrada, iIPICSTSaida, tTribItemNF.dIPIValor, tTribItemNF.dICMSPercRedBase, tNFiscal.lNumIntDoc)
'        If lErro <> AD_SQL_SUCESSO Then gError 204196
'
'        lErro = Comando_BuscarPrimeiro(alComando(5))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204197
'
'        Do While lErro = AD_SQL_SUCESSO
'
'            iAchou = 0
'
'            For Each objSpedC190 In colSpedC190
'
'                If objSpedC190.iCSTICMS = iTipoTribCST_ICMS And objSpedC190.sCFOP = tTribItemNF.sNaturezaOp And objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota Then
'                    If iTipoRegistro = ITEM_TIPO_FRETE Then
'                        objSpedC190.dValorOper = objSpedC190.dValorOper + tNFiscal.dValorFrete
'                    ElseIf iTipoRegistro = ITEM_TIPO_SEGURO Then
'                        objSpedC190.dValorOper = objSpedC190.dValorOper + tNFiscal.dValorSeguro
'                    ElseIf iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                        objSpedC190.dValorOper = objSpedC190.dValorOper + tNFiscal.dValorOutrasDespesas
''                    ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''                        objSpedC190.dValorOper = objSpedC190.dValorOper - tNFiscal.dValorDesconto
'                    End If
'                    If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                        objSpedC190.dValorOper = objSpedC190.dValorOper + tTribItemNF.dICMSSubstValor
'
'                        objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS + tTribItemNF.dICMSBase
'                        objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dICMSValor
'                        objSpedC190.dValorBCICMSST = objSpedC190.dValorBCICMSST + tTribItemNF.dICMSSubstBase
'                        objSpedC190.dValorICMSST = objSpedC190.dValorICMSST + tTribItemNF.dICMSSubstValor
'                        objSpedC190.dValorRedBC = objSpedC190.dValorRedBC + (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
''                    ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''                        objSpedC190.dValorOper = objSpedC190.dValorOper - tTribItemNF.dICMSSubstValor
''
''                        objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS - tTribItemNF.dICMSBase
''                        objSpedC190.dValorICMS = objSpedC190.dValorICMS - tTribItemNF.dICMSValor
''                        objSpedC190.dValorBCICMSST = objSpedC190.dValorBCICMSST - tTribItemNF.dICMSSubstBase
''                        objSpedC190.dValorICMSST = objSpedC190.dValorICMSST - tTribItemNF.dICMSSubstValor
''                        objSpedC190.dValorRedBC = objSpedC190.dValorRedBC - (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
'                    End If
'
'
'                    iAchou = 1
'                    Exit For
'                End If
'
'            Next
'
'            If iAchou = 0 Then
'
'                Set objSpedC190 = New ClassSpedC190
'
'                objSpedC190.iCSTICMS = iTipoTribCST_ICMS
'                objSpedC190.sCFOP = tTribItemNF.sNaturezaOp
'                objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota
'
'                If iTipoRegistro = ITEM_TIPO_FRETE Then
'                    objSpedC190.dValorOper = tNFiscal.dValorFrete
'                ElseIf iTipoRegistro = ITEM_TIPO_SEGURO Then
'                    objSpedC190.dValorOper = tNFiscal.dValorSeguro
'                ElseIf iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                    objSpedC190.dValorOper = tNFiscal.dValorOutrasDespesas
''                ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''                    objSpedC190.dValorOper = objSpedC190.dValorOper - tNFiscal.dValorDesconto
'                End If
'
'
'                If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                    objSpedC190.dValorOper = objSpedC190.dValorOper + tTribItemNF.dICMSSubstValor
'
'                    objSpedC190.dValorBCICMS = tTribItemNF.dICMSBase
'                    objSpedC190.dValorICMS = tTribItemNF.dICMSValor
'                    objSpedC190.dValorBCICMSST = tTribItemNF.dICMSSubstBase
'                    objSpedC190.dValorICMSST = tTribItemNF.dICMSSubstValor
'                    objSpedC190.dValorRedBC = (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
''                ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''                    objSpedC190.dValorOper = objSpedC190.dValorOper - tTribItemNF.dICMSSubstValor
''
''                    objSpedC190.dValorBCICMS = -tTribItemNF.dICMSBase
''                    objSpedC190.dValorICMS = -tTribItemNF.dICMSValor
''                    objSpedC190.dValorBCICMSST = -tTribItemNF.dICMSSubstBase
''                    objSpedC190.dValorICMSST = -tTribItemNF.dICMSSubstValor
''                    objSpedC190.dValorRedBC = -(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
'                End If
'
'                colSpedC190.Add objSpedC190
'
'            End If
'
'            If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'
'                If tNFiscal.dtDataEntrada >= dtDataIni And tNFiscal.dtDataEntrada <= dtDataFim Then
'                    dValCredICMS_objSpedD190 = dValCredICMS_objSpedD190 + tTribItemNF.dICMSValor
'                End If
'
'
''            ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''                If tNFiscal.dtDataEntrada >= dtDataIni And tNFiscal.dtDataEntrada <= dtDataFim Then
''                    dValCredICMS_objSpedD190 = dValCredICMS_objSpedD190 - tTribItemNF.dICMSValor
''                End If
'
'
'            End If
'
'            lErro = Comando_BuscarProximo(alComando(5))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204198
'
'        Loop
    
    For Each objSpedC190 In colSpedC190
    
        If objSpedC190.dValorOper > 0 Or objSpedC190.dValorBCICMS > 0 Or objSpedC190.dValorICMS > 0 Or _
           objSpedC190.dValorBCICMSST > 0 Or objSpedC190.dValorICMSST > 0 Or objSpedC190.dValorRedBC > 0 Then
    
            objSpedFiscal.sTipo = "D190"
            objSpedRegistroD190.sTipo = objSpedFiscal.sTipo
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            objSpedRegistroD190.iCount = objSpedRegistroD190.iCount + 1
    
            objSpedFiscal.sRegistro = "|D190|" & Format(objSpedC190.iCSTICMS, "000") & "|" & Format(objSpedC190.sCFOP, "0000") & "|" & Format(objSpedC190.dAliquotaICMS * 100, "fixed") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorOper <= 0, 0, Format(objSpedC190.dValorOper, "fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorBCICMS <= 0, 0, Format(objSpedC190.dValorBCICMS, "fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorICMS <= 0, 0, Format(objSpedC190.dValorICMS, "fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorRedBC <= 0, 0, Format(objSpedC190.dValorRedBC, "fixed")) & "||"
            
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203556
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
            If lErro <> AD_SQL_SUCESSO Then gError 203557
        
        End If
        
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_D190 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_D190:

   SpedFiscal_D190 = gErr

   Select Case gErr
    
        Case 203550
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203551, 203552, 203555
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL2", gErr)
    
        Case 203553, 203554
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOITEMNF", gErr)
    
        Case 203556
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203557
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case 204196 To 204198
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLNF", gErr)
    
        Case 204208
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203558)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
End Function

Function SpedFiscal_D100(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroD100 As ClassSpedRegistro, ByVal objSpedRegistroD190 As ClassSpedRegistro, ByVal objSpedRegistro0150 As ClassSpedRegistro, ByVal objSpedRegistro0175 As ClassSpedRegistro, ByVal objSpedRegistro0450 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal colMensagem As Collection, dValCredICMS_objSpedD190 As Double) As Long

Dim tNFiscal As typeNFiscal
Dim lSeqPai As Long
Dim lErro As Long
Dim alComando(1 To 9) As Long
Dim iAchou As Integer
Dim iIndice As Integer
Dim iCodSit As Integer
Dim tTipoDocInfo As typeTipoDocInfo
Dim iFreteRespons As Integer
Dim sCodParticipante As String
Dim tTributacaoNF As typeTributacaoNF
Dim lSeqPaiNF As Long
Dim objMensagem As ClassMensagem

On Error GoTo Erro_SpedFiscal_D100

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203559
    Next

    tNFiscal.sSerie = String(STRING_SERIE, 0)
    tNFiscal.sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
    tNFiscal.sPlaca = String(STRING_NFISCAL_PLACA, 0)
    tNFiscal.sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
    tNFiscal.sChvNFe = String(STRING_NFE_CHNFE, 0)
    tTipoDocInfo.sNomeTelaNFiscal = String(STRING_NOME, 0)

    'se for uma nota externa de entrada de conhecimento de frete
    lErro = Comando_Executar(alComando(4), "SELECT NumIntDoc, ModeloArqICMS,  ChvNFe, Tipo, Cliente, FilialCli, Fornecedor, FilialForn, Status, Complementar, Serie, NumNotaFiscal, DataEmissao, DataEntrada, DataSaida, ValorTotal, ClasseDocCPR, NumIntDocCPR, ValorDesconto, " & _
    "ValorProdutos, ValorFrete, FreteRespons, ValorSeguro, ValorOutrasDespesas, MensagemNota, FilialEntrega, CodTransportadora, Placa, PlacaUF, VolumeQuant, PesoLiq, PesoBruto, T.TipoMovtoEstoque, NomeTelaNFiscal From NFiscal AS N, TiposDocInfo AS T WHERE TipoNFiscal = Codigo AND Tipo =3 AND DataEntrada >= ? AND DataEntrada <= ? AND ModeloArqICMS = 8 AND FilialEmpresa = ?", _
    tNFiscal.lNumIntDoc, tTipoDocInfo.iModeloArqICMS, tNFiscal.sChvNFe, tTipoDocInfo.iTipo, tNFiscal.lCliente, tNFiscal.iFilialCli, tNFiscal.lFornecedor, tNFiscal.iFilialForn, tNFiscal.iStatus, tTipoDocInfo.iComplementar, tNFiscal.sSerie, tNFiscal.lNumNotaFiscal, tNFiscal.dtDataEmissao, tNFiscal.dtDataEntrada, tNFiscal.dtDataSaida, tNFiscal.dValorTotal, _
    tNFiscal.iClasseDocCPR, tNFiscal.lNumIntDocCPR, tNFiscal.dValorDesconto, tNFiscal.dValorProdutos, tNFiscal.dValorFrete, tNFiscal.iFreteRespons, tNFiscal.dValorSeguro, tNFiscal.dValorOutrasDespesas, tNFiscal.sMensagemNota, tNFiscal.iFilialEntrega, tNFiscal.iCodTransportadora, tNFiscal.sPlaca, tNFiscal.sPlacaUF, _
    tNFiscal.lVolumeQuant, tNFiscal.dPesoLiq, tNFiscal.dPesoBruto, tTipoDocInfo.iTipoMovtoEstoque, tTipoDocInfo.sNomeTelaNFiscal, dtDataIni, dtDataFim, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203560
    
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203561
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = SpedFiscal_D100_Trata_ItemNF(dtDataIni, dtDataFim, objSpedFiscal, tNFiscal, objSpedRegistro0150, objSpedRegistro0175, tTipoDocInfo, alComando, objSpedRegistro0450, colMensagem, objSpedRegistroD100, objSpedRegistroD190, dValCredICMS_objSpedD190)
        If lErro <> SUCESSO Then gError 204919
    
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203570
    
    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_D100 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_D100:

   SpedFiscal_D100 = gErr

   Select Case gErr
    
        Case 203559
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203560, 203561, 203570
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
    
        Case 203562, 203563, 203566, 203569, 204919
    
        Case 203564, 203565
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAONF", gErr)
    
        Case 203567
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203568
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203571)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_D500(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroD500 As ClassSpedRegistro, ByVal objSpedRegistroD590 As ClassSpedRegistro, ByVal objSpedRegistro0150 As ClassSpedRegistro, ByVal objSpedRegistro0175 As ClassSpedRegistro, ByVal objSpedRegistro0450 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal colMensagem As Collection, dValCredICMS_objSpedD590 As Double, colSpedD590UF As Collection, ByVal objFilialEmpresa As AdmFiliais) As Long

Dim tNFiscal As typeNFiscal
Dim lSeqPai As Long
Dim lErro As Long
Dim alComando(1 To 5) As Long
Dim iAchou As Integer
Dim iIndice As Integer
Dim iCodSit As Integer
Dim tTipoDocInfo As typeTipoDocInfo
Dim iFreteRespons As Integer
Dim sCodParticipante As String
Dim tTributacaoNF As typeTributacaoNF
Dim lSeqPaiNF As Long
Dim objMensagem As ClassMensagem

On Error GoTo Erro_SpedFiscal_D500

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203572
    Next


    tNFiscal.sSerie = String(STRING_SERIE, 0)
    tNFiscal.sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
    tNFiscal.sPlaca = String(STRING_NFISCAL_PLACA, 0)
    tNFiscal.sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
    tTipoDocInfo.sNomeTelaNFiscal = String(STRING_NOME, 0)


    'se for uma nota externa de entrada de telecomunicacoes
    lErro = Comando_Executar(alComando(4), "SELECT NumIntDoc, ModeloArqICMS, Tipo, Cliente, FilialCli, Fornecedor, FilialForn, Status, Complementar, Serie, NumNotaFiscal, DataEmissao, DataEntrada, DataSaida, ValorTotal, ClasseDocCPR, NumIntDocCPR, ValorDesconto, " & _
    "ValorProdutos, ValorFrete, FreteRespons, ValorSeguro, ValorOutrasDespesas, MensagemNota, FilialEntrega, CodTransportadora, Placa, PlacaUF, VolumeQuant, PesoLiq, PesoBruto, T.TipoMovtoEstoque, NomeTelaNFiscal From NFiscal AS N, TiposDocInfo AS T WHERE TipoNFiscal = Codigo AND (Tipo =3) AND DataEntrada >= ? AND DataEntrada <= ? AND ModeloArqICMS = 22 AND FilialEmpresa = ?", _
    tNFiscal.lNumIntDoc, tTipoDocInfo.iModeloArqICMS, tTipoDocInfo.iTipo, tNFiscal.lCliente, tNFiscal.iFilialCli, tNFiscal.lFornecedor, tNFiscal.iFilialForn, tNFiscal.iStatus, tTipoDocInfo.iComplementar, tNFiscal.sSerie, tNFiscal.lNumNotaFiscal, tNFiscal.dtDataEmissao, tNFiscal.dtDataEntrada, tNFiscal.dtDataSaida, tNFiscal.dValorTotal, _
    tNFiscal.iClasseDocCPR, tNFiscal.lNumIntDocCPR, tNFiscal.dValorDesconto, tNFiscal.dValorProdutos, tNFiscal.dValorFrete, tNFiscal.iFreteRespons, tNFiscal.dValorSeguro, tNFiscal.dValorOutrasDespesas, tNFiscal.sMensagemNota, tNFiscal.iFilialEntrega, tNFiscal.iCodTransportadora, tNFiscal.sPlaca, tNFiscal.sPlacaUF, _
    tNFiscal.lVolumeQuant, tNFiscal.dPesoLiq, tNFiscal.dPesoBruto, tTipoDocInfo.iTipoMovtoEstoque, tTipoDocInfo.sNomeTelaNFiscal, dtDataIni, dtDataFim, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203573
    
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203574
    
    Do While lErro = AD_SQL_SUCESSO
    
        'se é uma nota de cliente
        If tNFiscal.lCliente > 0 Then
        
        
            lErro = SpedFiscal_0150_Cliente(tNFiscal.lCliente, tNFiscal.iFilialCli, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
            If lErro <> SUCESSO Then gError 203575
        
        
        Else
    
    
            lErro = SpedFiscal_0150_Fornecedor(tNFiscal.lFornecedor, tNFiscal.iFilialForn, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
            If lErro <> SUCESSO Then gError 203576
    
        End If
    
    
        lErro = Retorna_UF_Subst_Tributaria(objFilialEmpresa, tNFiscal, tTipoDocInfo)
        If lErro <> SUCESSO Then gError 204931
            
        '************ registro C100 ********************************

        lErro = NF_Obter_Situacao(alComando(5), tNFiscal, tTipoDocInfo, iCodSit)
        If lErro <> SUCESSO Then gError 204931

'        If tTipoDocInfo.iComplementar = 1 Then
'
'            iCodSit = 6
'
'        Else
'
'            iCodSit = 0
'
'        End If
'
'
'        If tNFiscal.iStatus = STATUS_CANCELADO Then iCodSit = 2
    
'        If tNFiscal.dValorFrete <> 0 Then
'            If tNFiscal.iFreteRespons = 0 Or tNFiscal.iFreteRespons = 1 Then
'                iFreteRespons = 1
'            Else
'                iFreteRespons = 2
'            End If
'        Else
'            iFreteRespons = 9
'        End If
        
        Select Case tNFiscal.iFreteRespons
        
            Case 0, 1 'Emitente
                If tNFiscal.dtDataEmissao < #1/1/2012# Then
                    iFreteRespons = 1
                Else
                    iFreteRespons = 0
                End If
            Case 2 'Destinatário
                If tNFiscal.dtDataEmissao < #1/1/2012# Then
                    iFreteRespons = 2
                Else
                    iFreteRespons = 1
                End If
            Case 3 'Terceiros
                If tNFiscal.dtDataEmissao < #1/1/2012# Then
                    iFreteRespons = 0
                Else
                    iFreteRespons = 2
                End If
            Case 4 'Sem Frete
                iFreteRespons = 9

        End Select
    
        'pesquisa a tributacaoNF
        lErro = Comando_Executar(alComando(2), "SELECT ICMSBase, ICMSValor, ICMSSubstBase, ICMSSubstValor, IPIValor, PISCredito, COFINSCredito, PISRetido, COFINSRetido, ISSAliquota, ICMSCredito FROM TributacaoNF WHERE NumIntDoc = ?", _
        tTributacaoNF.dICMSBase, tTributacaoNF.dICMSValor, tTributacaoNF.dICMSSubstBase, tTributacaoNF.dICMSSubstValor, tTributacaoNF.dIPIValor, tTributacaoNF.dPISCredito, tTributacaoNF.dCOFINSCredito, tTributacaoNF.dPISRetido, tTributacaoNF.dCOFINSRetido, tTributacaoNF.dISSAliquota, tTributacaoNF.dICMSCredito, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203577
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203578

        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        lSeqPaiNF = objSpedFiscal.lSeq
            
        'se a nota nao for cancelada
        If iCodSit <> 2 And iCodSit <> 4 Then
        
            Set objMensagem = New ClassMensagem
    
            If Len(Trim(tNFiscal.sMensagemNota)) <> 0 Or Len(Trim(tNFiscal.sMensagemCorpoNota)) <> 0 Then
                lErro = SpedFiscal_0450(objMensagem, objSpedFiscal, objSpedRegistro0450, tNFiscal, lSeqPaiNF, colMensagem)
                If lErro <> SUCESSO Then gError 203579
            End If
        
            objSpedFiscal.sRegistro = "|D500|0|1|" & sCodParticipante & "|" & Format(tTipoDocInfo.iModeloArqICMS, "00") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(iCodSit, "00") & "|" & Ajusta_Serie(Desconverte_Serie_Eletronica(tNFiscal.sSerie)) & "||" & tNFiscal.lNumNotaFiscal & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tNFiscal.dtDataEmissao, "ddmmyyyy") & "|" & Format(tNFiscal.dtDataEntrada, "ddmmyyyy") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tNFiscal.dValorTotal, "Fixed") & "|" & IIf(tNFiscal.dValorDesconto = 0, 0, Format(tNFiscal.dValorDesconto, "Fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tNFiscal.dValorProdutos, "Fixed") & "|||"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tNFiscal.dValorOutrasDespesas, "Fixed") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSBase = 0, 0, Format(tTributacaoNF.dICMSBase, "Fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSCredito = 0, 0, Format(tTributacaoNF.dICMSCredito, "Fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objMensagem.iCodigo = 0, "", CStr(objMensagem.iCodigo)) & "|||||"
            
        Else
        
            objSpedFiscal.sRegistro = "|D500|0|1||" & Format(tTipoDocInfo.iModeloArqICMS, "00") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(iCodSit, "00") & "|" & Ajusta_Serie(Desconverte_Serie_Eletronica(tNFiscal.sSerie)) & "||" & tNFiscal.lNumNotaFiscal & "|" & Format(tNFiscal.dtDataEmissao, "ddmmyyyy") & "|||||||||||||||"
        
        End If
        
        objSpedFiscal.sTipo = "D500"
        objSpedRegistroD500.sTipo = objSpedFiscal.sTipo
        objSpedRegistroD500.iCount = objSpedRegistroD500.iCount + 1
            
        
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203580
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, lSeqPaiNF, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203581
    
    
        'se a nota nao for cancelada
        If iCodSit <> 2 And iCodSit <> 4 Then
    
            '*****************  registro D590 ******************************************************
            lErro = SpedFiscal_D590(tNFiscal, objSpedFiscal, objSpedRegistroD590, dtDataIni, dtDataFim, dValCredICMS_objSpedD590, lSeqPaiNF, colSpedD590UF)
            If lErro <> SUCESSO Then gError 203582
    
        End If
    
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203583
    
    Loop


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_D500 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_D500:

   SpedFiscal_D500 = gErr

   Select Case gErr
    
        Case 203572
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203573, 203574, 203583
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
    
        Case 203575, 203576, 203579, 203582
    
        Case 203577, 203578
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAONF", gErr)
    
        Case 203580
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203581
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203584)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_D590(ByRef tNFiscal As typeNFiscal, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroD590 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date, dValCredICMS_objSpedD590 As Double, ByVal lSeqPai As Long, colSpedD590UF As Collection) As Long

Dim tItemNF As typeItemNF
Dim dTotalRateioNF As Double
Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim colSpedC190 As New Collection 'é C190 mesmo pois estamos aproveitando a classe ClassSpedC190
Dim objSpedC190 As ClassSpedC190
Dim iAchou As Integer
Dim iIndice As Integer
Dim iTipoTribCST_ICMS As Integer
Dim iIPICSTEntrada As Integer
Dim iIPICSTSaida As Integer
Dim tProduto As typeProduto
Dim tTribItemNF As typeTribItemNF
Dim iTipoRegistro As Integer
Dim objSpedD590UF As ClassSpedD590UF
Dim objTribItem As ClassTributacaoDocItem
Dim iTipoTribCST_ICMSAux As Integer
Dim dValorFCP As Double
Dim dValorFCPST As Double
Dim dValorFCPRet As Double

On Error GoTo Erro_SpedFiscal_D590

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203585
    Next
            
    tItemNF.sProduto = String(STRING_PRODUTO, 0)
    tItemNF.sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
    tItemNF.sUnidadeMed = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)

    lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Item, Produto, DescricaoItem, Quantidade, UnidadeMed, ClasseUM, PrecoUnitario, ValorDesconto FROM ItensNFiscal, Produtos  WHERE Produto = Codigo AND NumIntNF  = ?", _
    tItemNF.lNumIntDoc, tItemNF.iItem, tItemNF.sProduto, tItemNF.sDescricaoItem, tItemNF.dQuantidade, tItemNF.sUnidadeMed, tProduto.iClasseUM, tItemNF.dPrecoUnitario, tItemNF.dValorDesconto, tNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 203586

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203587

    Do While lErro = AD_SQL_SUCESSO

        tTribItemNF.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)

        lErro = Comando_Executar(alComando(3), "SELECT OrigemMercadoria, TipoTribCST, NaturezaOP, ICMSBase, ICMSAliquota, ICMSValor, ICMSCredito, ICMSSubstBase, ICMSSubstAliquota, ICMSSubstValor, IPI.CSTEntrada, IPI.CSTSaida, " & _
        "IPIBaseCalculo, IPIAliquota, IPIValor, PISCredito, TipoTributacao, ICMSPercRedBase, ICMSvFCP, ICMSvFCPST, ICMSvFCPSTRet FROM TributacaoDocItem, tiposTribICMS AS ICMS, TiposTribIPI AS IPI  WHERE TipoDoc = 0 AND NumIntDoc = ? AND Item = ? AND ICMS.Tipo = ICMSTipo AND IPI.Tipo = IPITipo", _
        tTribItemNF.iOrigemMercadoria, iTipoTribCST_ICMSAux, tTribItemNF.sNaturezaOp, tTribItemNF.dICMSBase, tTribItemNF.dICMSAliquota, tTribItemNF.dICMSValor, tTribItemNF.dICMSCredito, tTribItemNF.dICMSSubstBase, tTribItemNF.dICMSSubstAliquota, tTribItemNF.dICMSSubstValor, _
        iIPICSTEntrada, iIPICSTSaida, tTribItemNF.dIPIBaseCalculo, tTribItemNF.dIPIAliquota, tTribItemNF.dIPIValor, tTribItemNF.dPISCredito, tTribItemNF.iTipoTributacao, tTribItemNF.dICMSPercRedBase, dValorFCP, dValorFCPST, dValorFCPRet, tNFiscal.lNumIntDoc, tItemNF.iItem)
        If lErro <> AD_SQL_SUCESSO Then gError 203588
    
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203589
        
        iTipoTribCST_ICMS = iTipoTribCST_ICMSAux
        
        Set objTribItem = New ClassTributacaoDocItem
        objTribItem.lNumIntDocItem = tItemNF.lNumIntDoc
        objTribItem.iTipoDoc = TIPODOC_TRIB_NF
        lErro = ItemNF_Le_BasePropria(objTribItem, iTipoTribCST_ICMS, iIPICSTEntrada)
        If lErro <> SUCESSO Then gError 204209
        
        iTipoTribCST_ICMS = iTipoTribCST_ICMS + (100 * tTribItemNF.iOrigemMercadoria)
        
        tTribItemNF.dICMSBase = objTribItem.dICMSBase
        tTribItemNF.dICMSAliquota = objTribItem.dICMSAliquota
        tTribItemNF.dICMSValor = objTribItem.dICMSValor
        tTribItemNF.dICMSCredito = objTribItem.dICMSCredito
        
        tTribItemNF.dICMSSubstBase = objTribItem.dICMSSubstBase
        tTribItemNF.dICMSSubstAliquota = objTribItem.dICMSSubstAliquota
        tTribItemNF.dICMSSubstValor = objTribItem.dICMSSubstValor
        tTribItemNF.dICMSSubstPercMVA = objTribItem.dICMSSubstPercMVA
        tTribItemNF.dICMSSubstPercRedBase = objTribItem.dICMSSubstPercRedBase

        tTribItemNF.dIPIBaseCalculo = objTribItem.dIPIBaseCalculo
        tTribItemNF.dIPIAliquota = objTribItem.dIPIAliquota
        tTribItemNF.dIPIValor = objTribItem.dIPIValor
        tTribItemNF.dIPICredito = objTribItem.dIPICredito

        lErro = Rateia_Compl_ItemNF(tNFiscal, tItemNF, tTribItemNF)
        If lErro <> SUCESSO Then gError 204209

        iAchou = 0

        For Each objSpedC190 In colSpedC190
        
            If objSpedC190.iCSTICMS = iTipoTribCST_ICMS And objSpedC190.sCFOP = tTribItemNF.sNaturezaOp And objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota Then
                objSpedC190.dValorOper = objSpedC190.dValorOper + (tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
                If tNFiscal.dValorProdutos > 0 Then
                    objSpedC190.dValorOper = objSpedC190.dValorOper + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
                    dTotalRateioNF = dTotalRateioNF + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'                    objSpedC190.dValorOper = objSpedC190.dValorOper - Round(((tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)

                End If
                    
                objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS + tTribItemNF.dICMSBase
'mario                objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dICMSValor
                objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dICMSCredito

                objSpedC190.dValorBCICMSST = objSpedC190.dValorBCICMSST + tTribItemNF.dICMSSubstBase
                objSpedC190.dValorICMSST = objSpedC190.dValorICMSST + tTribItemNF.dICMSSubstValor
                objSpedC190.dValorRedBC = objSpedC190.dValorRedBC + Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
                objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS - Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase) 'objSpedC190.dValorRedBC
                objSpedC190.dValorFCP = objSpedC190.dValorFCP + dValorFCP
                objSpedC190.dValorFCPST = objSpedC190.dValorFCPST + dValorFCPST
                objSpedC190.dValorFCPRet = objSpedC190.dValorFCPRet + dValorFCPRet
                iAchou = 1
                Exit For
            End If
            
        Next

        If iAchou = 0 Then
        
            Set objSpedC190 = New ClassSpedC190
            
            objSpedC190.iCSTICMS = iTipoTribCST_ICMS
            objSpedC190.sCFOP = tTribItemNF.sNaturezaOp
            objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota
            objSpedC190.dValorOper = (tItemNF.dQuantidade * tItemNF.dPrecoUnitario)
            If tNFiscal.dValorProdutos > 0 Then
                objSpedC190.dValorOper = objSpedC190.dValorOper + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
                dTotalRateioNF = Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'                objSpedC190.dValorOper = objSpedC190.dValorOper - Round(((tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)

            End If
            objSpedC190.dValorBCICMS = tTribItemNF.dICMSBase
'mario            objSpedC190.dValorICMS = tTribItemNF.dICMSValor
            objSpedC190.dValorICMS = tTribItemNF.dICMSCredito
            objSpedC190.dValorBCICMSST = tTribItemNF.dICMSSubstBase
            objSpedC190.dValorICMSST = tTribItemNF.dICMSSubstValor
            objSpedC190.dValorRedBC = Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
            objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS - Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase) 'objSpedC190.dValorRedBC
            objSpedC190.dValorFCP = dValorFCP
            objSpedC190.dValorFCPST = dValorFCPST
            objSpedC190.dValorFCPRet = dValorFCPRet
            
            colSpedC190.Add objSpedC190
            
        End If
        
            iAchou = 0
    
            For Each objSpedD590UF In colSpedD590UF
            
                If objSpedD590UF.sUF = tNFiscal.sUFSubstTrib And objSpedD590UF.sCFOP = tTribItemNF.sNaturezaOp Then
                    objSpedD590UF.dValorBCICMSST = objSpedD590UF.dValorBCICMSST + tTribItemNF.dICMSSubstBase
                    objSpedD590UF.dValorICMSST = objSpedD590UF.dValorICMSST + tTribItemNF.dICMSSubstValor
                    
                    iAchou = 1
                    Exit For
                End If
                
            Next
    
            If iAchou = 0 Then
            
                Set objSpedD590UF = New ClassSpedD590UF
                
                objSpedD590UF.sUF = tNFiscal.sUFSubstTrib
                objSpedD590UF.sCFOP = tTribItemNF.sNaturezaOp
                
                objSpedD590UF.dValorBCICMSST = tTribItemNF.dICMSSubstBase
                objSpedD590UF.dValorICMSST = tTribItemNF.dICMSSubstValor
                
                colSpedD590UF.Add objSpedD590UF
                
            End If
        

        'soma os valores de ICMS dos registros D500 que estiverem com data de saida/entrada dentro do periodo para poder colocar
        'posteriormente no registro E110
        If tNFiscal.dtDataEntrada >= dtDataIni And tNFiscal.dtDataEntrada <= dtDataFim Then
            dValCredICMS_objSpedD590 = dValCredICMS_objSpedD590 + tTribItemNF.dICMSCredito
        End If

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203590

    Loop
    
    '********************** registro D590 *******************************
'    If tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro <> dTotalRateioNF Then
'
'        Set objSpedC190 = colSpedC190.Item(1)
'
'        objSpedC190.dValorOper = objSpedC190.dValorOper + Round(tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - dTotalRateioNF, 2)
'
'    End If
    
    
'    tTribItemNF.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'    lErro = Comando_Executar(alComando(5), "SELECT T.Tipo, TipoTribCST, NaturezaOP, ICMSBase, ICMSAliquota, ICMSValor, ICMSSubstBase, ICMSSubstValor, " & _
'    "IPIValor, ICMSPercRedBase FROM TributacaoComplNF AS T, tiposTribICMS AS ICMS  WHERE T.NumIntDoc  = ? AND ICMS.Tipo = ICMSTipo AND (T.Tipo = 1 Or T.Tipo=2 Or T.Tipo = 4)", _
'    iTipoRegistro, iTipoTribCST_ICMS, tTribItemNF.sNaturezaOp, tTribItemNF.dICMSBase, tTribItemNF.dICMSAliquota, tTribItemNF.dICMSValor, tTribItemNF.dICMSSubstBase, tTribItemNF.dICMSSubstValor, _
'    tTribItemNF.dIPIValor, tTribItemNF.dICMSPercRedBase, tNFiscal.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 204199
'
'    lErro = Comando_BuscarPrimeiro(alComando(5))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204200
'
'    Do While lErro = AD_SQL_SUCESSO
'
'
'        iAchou = 0
'
'        For Each objSpedC190 In colSpedC190
'
'            If objSpedC190.iCSTICMS = iTipoTribCST_ICMS And objSpedC190.sCFOP = tTribItemNF.sNaturezaOp And objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota Then
'                If iTipoRegistro = ITEM_TIPO_FRETE Then
'                    objSpedC190.dValorOper = objSpedC190.dValorOper + tNFiscal.dValorFrete
'                ElseIf iTipoRegistro = ITEM_TIPO_SEGURO Then
'                    objSpedC190.dValorOper = objSpedC190.dValorOper + tNFiscal.dValorSeguro
'                ElseIf iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                    objSpedC190.dValorOper = objSpedC190.dValorOper + tNFiscal.dValorOutrasDespesas
''                ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''                    objSpedC190.dValorOper = objSpedC190.dValorOper - tNFiscal.dValorDesconto
'                End If
'                If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'
'                    objSpedC190.dValorOper = objSpedC190.dValorOper + tTribItemNF.dICMSSubstValor + tTribItemNF.dIPIValor
'                    objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS + tTribItemNF.dICMSBase
'                    objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dICMSValor
'                    objSpedC190.dValorBCICMSST = objSpedC190.dValorBCICMSST + tTribItemNF.dICMSSubstBase
'                    objSpedC190.dValorICMSST = objSpedC190.dValorICMSST + tTribItemNF.dICMSSubstValor
'                    objSpedC190.dValorRedBC = objSpedC190.dValorRedBC + (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
''                ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''                    objSpedC190.dValorOper = objSpedC190.dValorOper - tTribItemNF.dICMSSubstValor
''
''                    objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS - tTribItemNF.dICMSBase
''                    objSpedC190.dValorICMS = objSpedC190.dValorICMS - tTribItemNF.dICMSValor
''                    objSpedC190.dValorBCICMSST = objSpedC190.dValorBCICMSST - tTribItemNF.dICMSSubstBase
''                    objSpedC190.dValorICMSST = objSpedC190.dValorICMSST - tTribItemNF.dICMSSubstValor
''                    objSpedC190.dValorRedBC = objSpedC190.dValorRedBC - (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
'                End If
'
'
'                iAchou = 1
'                Exit For
'            End If
'
'        Next
'
'        If iAchou = 0 Then
'
'            Set objSpedC190 = New ClassSpedC190
'
'            objSpedC190.iCSTICMS = iTipoTribCST_ICMS
'            objSpedC190.sCFOP = tTribItemNF.sNaturezaOp
'            objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota
'
'            If iTipoRegistro = ITEM_TIPO_FRETE Then
'                objSpedC190.dValorOper = tNFiscal.dValorFrete
'            ElseIf iTipoRegistro = ITEM_TIPO_SEGURO Then
'                objSpedC190.dValorOper = tNFiscal.dValorSeguro
'            ElseIf iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                objSpedC190.dValorOper = tNFiscal.dValorOutrasDespesas
''            ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''                objSpedC190.dValorOper = objSpedC190.dValorOper - tNFiscal.dValorDesconto
'            End If
'
'
'            If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                objSpedC190.dValorOper = objSpedC190.dValorOper + tTribItemNF.dICMSSubstValor
'
'                objSpedC190.dValorBCICMS = tTribItemNF.dICMSBase
'                objSpedC190.dValorICMS = tTribItemNF.dICMSValor
'                objSpedC190.dValorBCICMSST = tTribItemNF.dICMSSubstBase
'                objSpedC190.dValorICMSST = tTribItemNF.dICMSSubstValor
'                objSpedC190.dValorRedBC = (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
''            ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''                objSpedC190.dValorOper = objSpedC190.dValorOper - tTribItemNF.dICMSSubstValor
''
''                objSpedC190.dValorBCICMS = -tTribItemNF.dICMSBase
''                objSpedC190.dValorICMS = -tTribItemNF.dICMSValor
''                objSpedC190.dValorBCICMSST = -tTribItemNF.dICMSSubstBase
''                objSpedC190.dValorICMSST = -tTribItemNF.dICMSSubstValor
''                objSpedC190.dValorRedBC = -(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
'            End If
'
'            colSpedC190.Add objSpedC190
'
'        End If
'
'        If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'
'            'soma os valores de ICMS dos registros C100 de entrada que estiverem com data de entrada dentro do periodo para poder colocar
'            'posteriormente no registro E110
'            If tNFiscal.dtDataEntrada >= dtDataIni And tNFiscal.dtDataEntrada <= dtDataFim Then
'                dValCredICMS_objSpedD590 = dValCredICMS_objSpedD590 + tTribItemNF.dICMSValor
'            End If
'
'
''        ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''            'soma os valores de ICMS dos registros C100 de entrada que estiverem com data de entrada dentro do periodo para poder colocar
''            'posteriormente no registro E110
''            If tNFiscal.dtDataEntrada >= dtDataIni And tNFiscal.dtDataEntrada <= dtDataFim Then
''                dValCredICMS_objSpedD590 = dValCredICMS_objSpedD590 - tTribItemNF.dICMSValor
''            End If
'
'        End If
'
'        lErro = Comando_BuscarProximo(alComando(5))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204201
'
'    Loop
    
    '********************** registro D190 *******************************
    If tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto <> dTotalRateioNF Then

        Set objSpedC190 = colSpedC190.Item(1)

        objSpedC190.dValorOper = objSpedC190.dValorOper + Round(tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto - dTotalRateioNF, 2)

    End If
    
    For Each objSpedC190 In colSpedC190
    
        If objSpedC190.dValorOper > 0 Or objSpedC190.dValorBCICMS > 0 Or objSpedC190.dValorICMS > 0 Or _
           objSpedC190.dValorBCICMSST > 0 Or objSpedC190.dValorICMSST > 0 Or objSpedC190.dValorRedBC > 0 Then
    
            objSpedFiscal.sTipo = "D590"
            objSpedRegistroD590.sTipo = objSpedFiscal.sTipo
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            objSpedRegistroD590.iCount = objSpedRegistroD590.iCount + 1
    
            objSpedFiscal.sRegistro = "|D590|" & Format(objSpedC190.iCSTICMS, "000") & "|" & Format(objSpedC190.sCFOP, "0000") & "|" & Format(objSpedC190.dAliquotaICMS * 100, "fixed") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorOper <= 0, 0, Format(objSpedC190.dValorOper, "fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorBCICMS <= 0, 0, Format(objSpedC190.dValorBCICMS, "fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorICMS <= 0, 0, Format(objSpedC190.dValorICMS, "fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorBCICMSST <= 0, 0, Format(objSpedC190.dValorBCICMSST, "fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorICMSST <= 0, 0, Format(objSpedC190.dValorICMSST, "fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorRedBC <= 0, 0, Format(objSpedC190.dValorRedBC, "fixed")) & "||"
            
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203591
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
            If lErro <> AD_SQL_SUCESSO Then gError 203592
            
        End If
        
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_D590 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_D590:

   SpedFiscal_D590 = gErr

   Select Case gErr
    
        Case 203585
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203586, 203587, 203590
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL2", gErr)
    
        Case 203588, 203589
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOITEMNF", gErr)
    
        Case 203591
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203592
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case 204199 To 204201
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLNF", gErr)
            
        Case 204209
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203593)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
End Function

Function SpedFiscal_E001(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE001 As ClassSpedRegistro) As Long
        
Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_SpedFiscal_E001
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203594
    Next
        
    '******************* registro E001 *******************

    objSpedFiscal.sTipo = "E001"
    objSpedRegistroE001.sTipo = objSpedFiscal.sTipo

    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroE001.iCount = objSpedRegistroE001.iCount + 1
    
    objSpedFiscal.sRegistro = "|E001|0|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203595

    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203596

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E001 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E001:

   SpedFiscal_E001 = gErr

   Select Case gErr
    
        Case 203594
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203595
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203596
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203597)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_E100(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE100 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date, lSeqE100 As Long) As Long
        
Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_SpedFiscal_E100
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203598
    Next
        
    '******************* registro E100 *******************

    objSpedFiscal.sTipo = "E100"
    objSpedRegistroE100.sTipo = objSpedFiscal.sTipo

    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroE100.iCount = objSpedRegistroE100.iCount + 1
    
    objSpedFiscal.sRegistro = "|E100|" & Format(dtDataIni, "ddmmyyyy") & "|" & Format(dtDataFim, "ddmmyyyy") & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203599

    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203600

    lSeqE100 = objSpedFiscal.lSeq

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E100 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E100:

   SpedFiscal_E100 = gErr

   Select Case gErr
    
        Case 203598
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203599
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203600
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203601)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function
        
Function SpedFiscal_E111(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE111 As ClassSpedRegistro, ByVal objFilialEmpresa As AdmFiliais, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal lSeqPai As Long, dTotAjDebitos As Double, dTotEstCreditos As Double, dTotAjCreditos As Double, dTotEstDebitos As Double, dTotDeducoes As Double) As Long
        
Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim tRegApuracaoItens As typeRegApuracaoItens
Dim iRegApuracaoICMS As Integer
Dim iRegApuracaoICMSAtual As Integer
Dim iSequencial As Integer
Dim sCodAjApur As String
Dim iTipoReg As Integer

On Error GoTo Erro_SpedFiscal_E111
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203602
    Next
    
    iRegApuracaoICMSAtual = -1
        
    tRegApuracaoItens.sDescricao = String(STRING_DESCRICAO_APURACAO, 0)
        
    'Verifica se o Item de apuração ICMS está cadastrada
    lErro = Comando_Executar(alComando(2), "SELECT Valor, R.Descricao, RegApuracaoICMSSped, TipoReg FROM RegApuracaoICMSItem AS R, TiposRegApuracaoICMS WHERE TipoReg = Codigo AND Data >= ? AND Data <= ? AND FilialEmpresa = ? AND RegApuracaoICMSSped BETWEEN 0 AND 5  ORDER BY RegApuracaoICMSSped ", _
    tRegApuracaoItens.dValor, tRegApuracaoItens.sDescricao, iRegApuracaoICMS, iTipoReg, dtDataIni, dtDataFim, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 203603
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203604

    Do While lErro = AD_SQL_SUCESSO
        
        If iRegApuracaoICMSAtual <> iRegApuracaoICMS Then
            iRegApuracaoICMSAtual = iRegApuracaoICMS
            iSequencial = 1
        Else
            iSequencial = iSequencial + 1
        End If
        
'        0 – Outros débitos;
'        1 – Estorno de créditos;
'        2 – Outros créditos;
'        3 – Estorno de débitos;
'        4 – Deduções do imposto apurado;
'        5 – Débitos Especiais

        If iRegApuracaoICMS = 0 Then
            dTotAjDebitos = dTotAjDebitos + tRegApuracaoItens.dValor
        ElseIf iRegApuracaoICMS = 1 Then
            dTotEstCreditos = dTotEstCreditos + tRegApuracaoItens.dValor
        ElseIf iRegApuracaoICMS = 2 Then
            dTotAjCreditos = dTotAjCreditos + tRegApuracaoItens.dValor
        ElseIf iRegApuracaoICMS = 3 Then
            dTotEstDebitos = dTotEstDebitos + tRegApuracaoItens.dValor
        ElseIf iRegApuracaoICMS = 4 Then
            dTotDeducoes = dTotDeducoes + tRegApuracaoItens.dValor
        End If
        
'        If iRegApuracaoICMS = ICMS_OUTROS_DEBITOS Then
'            dTotAjDebitos = dTotAjDebitos + tRegApuracaoItens.dValor
'        ElseIf iRegApuracaoICMS = ICMS_ESTORNO_CREDITOS Then
'            dTotEstCreditos = dTotEstCreditos + tRegApuracaoItens.dValor
'        ElseIf iRegApuracaoICMS = ICMS_OUTROS_CREDITOS Then
'            dTotAjCreditos = dTotAjCreditos + tRegApuracaoItens.dValor
'        ElseIf iRegApuracaoICMS = ICMS_ESTORNO_DEBITOS Then
'            dTotEstDebitos = dTotEstDebitos + tRegApuracaoItens.dValor
'        ElseIf iRegApuracaoICMS = -1 And iTipoReg = 12 Then
'            dTotAjCreditos = dTotAjCreditos + tRegApuracaoItens.dValor
'        End If
        
        sCodAjApur = objFilialEmpresa.objEndereco.sSiglaEstado & "0" & iRegApuracaoICMS & Format(9999, "0000") '??? colocar codigo vindo da tabela de lctos da apuracao
        
        '******************* registro E111 *******************
     
        objSpedFiscal.sTipo = "E111"
        objSpedRegistroE111.sTipo = objSpedFiscal.sTipo
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroE111.iCount = objSpedRegistroE111.iCount + 1
        
        objSpedFiscal.sRegistro = "|E111|" & sCodAjApur & "|" & Desacentua_SpedFiscal(tRegApuracaoItens.sDescricao) & "|" & Format(tRegApuracaoItens.dValor, "fixed") & "|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203605
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
        If lErro <> AD_SQL_SUCESSO Then gError 203606

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203607

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E111 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E111:

   SpedFiscal_E111 = gErr

   Select Case gErr
    
        Case 203602
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203603, 203604, 203607
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMSITEM", gErr)
    
        Case 203605
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203606
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203608)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_E116(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE116 As ClassSpedRegistro, ByVal objFilialEmpresa As AdmFiliais, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal lSeqPai As Long, ByVal dValorICMSRecolher As Double) As Long
        
Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim tRegApuracaoItens As typeRegApuracaoItens
Dim iRegApuracaoICMS As Integer
Dim iRegApuracaoICMSAtual As Integer
Dim iSequencial As Integer
Dim sCodAjApur As String
Dim tGNRICMS As typeGNRICMS
Dim tGuiasICMS As typeGuiasICMS
Dim dTotalRecolher As Double
Dim iFilialEmpresa As Integer


On Error GoTo Erro_SpedFiscal_E116
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203609
    Next
    
    iFilialEmpresa = objFilialEmpresa.iCodFilial
    
    iRegApuracaoICMSAtual = -1
        
    tGNRICMS.sCodReceita = String(STRING_GNRICMS_CODRECEITA, 0)
    tGNRICMS.sCodObrigRecolher = String(STRING_GNRICMS_CODOBRIGRECOLHER, 0)
        
    'Verifica se tem Guias GNR cadastradas
    lErro = Comando_ExecutarPos(alComando(2), "SELECT CodReceita, CodObrigRecolher, Vencimento, Valor FROM GNRICMS WHERE DataRef >= ? And DataRef <= ? AND FilialEmpresa = ? ", 0, _
    tGNRICMS.sCodReceita, tGNRICMS.sCodObrigRecolher, tGNRICMS.dtDataVencimento, tGNRICMS.dValor, dtDataIni, dtDataFim, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203610
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203611

    Do While lErro = AD_SQL_SUCESSO
        
        dTotalRecolher = dTotalRecolher + tGNRICMS.dValor
        
        objSpedFiscal.sTipo = "E116"
        objSpedRegistroE116.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroE116.iCount = objSpedRegistroE116.iCount + 1
        
        objSpedFiscal.sRegistro = "|E116|" & tGNRICMS.sCodObrigRecolher & "|" & Format(tGNRICMS.dValor, "Fixed") & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tGNRICMS.dtDataVencimento, "ddmmyyyy") & "|" & SFF_Texto(tGNRICMS.sCodReceita) & "|||||" & IIf(dtDataIni >= #1/1/2011#, Format(dtDataIni, "mmyyyy") & "|", "")
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203612
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
        If lErro <> AD_SQL_SUCESSO Then gError 203613

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203614

    Loop

    tGuiasICMS.sCodReceita = String(STRING_GUIASICMS_CODRECEITA, 0)
    tGuiasICMS.sCodObrigRecolher = String(STRING_GUIASICMS_CODOBRIGRECOLHER, 0)

    'Verifica se tem GuiaICMS cadastrada
    lErro = Comando_ExecutarPos(alComando(3), "SELECT CodReceita, CodObrigRecolher, Vencimento, Valor FROM GuiasICMS WHERE ApuracaoDe >= ? And ApuracaoDe <= ? AND ApuracaoAte >= ? And ApuracaoAte <= ? AND FilialEmpresa = ? ", 0, _
    tGuiasICMS.sCodReceita, tGuiasICMS.sCodObrigRecolher, tGuiasICMS.dtDataVencimento, tGuiasICMS.dValor, dtDataIni, dtDataFim, dtDataIni, dtDataFim, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203615
    
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203616

    Do While lErro = AD_SQL_SUCESSO
        
        dTotalRecolher = dTotalRecolher + tGuiasICMS.dValor
        
        objSpedFiscal.sTipo = "E116"
        objSpedRegistroE116.sTipo = objSpedFiscal.sTipo
            
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroE116.iCount = objSpedRegistroE116.iCount + 1
        
        objSpedFiscal.sRegistro = "|E116|" & SFF_Texto(tGuiasICMS.sCodObrigRecolher) & "|" & Format(tGuiasICMS.dValor, "Fixed") & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tGuiasICMS.dtDataVencimento, "ddmmyyyy") & "|" & SFF_Texto(tGuiasICMS.sCodReceita) & "|||||" & IIf(dtDataIni >= #1/1/2011#, Format(dtDataIni, "mmyyyy") & "|", "")
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203617
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
        If lErro <> AD_SQL_SUCESSO Then gError 203618

        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203619

    Loop


    'os valores tem que ser iguais...senao erro...mostrar os totais na mensagem
    '??? nao vou dar erro para poder gerar arq antes da apuracao
    'If Abs(dValorICMSRecolher - dTotalRecolher) > DELTA_VALORMONETARIO Then gError 203620

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E116 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E116:

   SpedFiscal_E116 = gErr

   Select Case gErr
    
        Case 203609
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203610, 203611, 203614
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GNRICMS", gErr)
    
        Case 203612, 203617
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203613, 203618
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case 203615, 203616, 203619
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GUIAICMS", gErr)
    
        Case 203620
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ICMS_RECOLHER_DIFERE_GUIAS", gErr, dValorICMSRecolher, dTotalRecolher)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203621)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_E500(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE500 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date, lSeqE500 As Long) As Long
        
Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_SpedFiscal_E500
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203622
    Next
        
    '******************* registro E500 *******************

    objSpedFiscal.sTipo = "E500"
    objSpedRegistroE500.sTipo = objSpedFiscal.sTipo

    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroE500.iCount = objSpedRegistroE500.iCount + 1
    
    lSeqE500 = objSpedFiscal.lSeq
    
    objSpedFiscal.sRegistro = "|E500|0|" & Format(dtDataIni, "ddmmyyyy") & "|" & Format(dtDataFim, "ddmmyyyy") & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203623

    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203624

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E500 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E500:

   SpedFiscal_E500 = gErr

   Select Case gErr
    
        Case 203622
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203623
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203624
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203625)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_E510(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE510 As ClassSpedRegistro, ByVal colSpedE510 As Collection, dValDebIPI As Double, dValCredIPI As Double, ByVal lSeqE500 As Long) As Long
        
Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objSpedE510 As ClassSpedE510

On Error GoTo Erro_SpedFiscal_E510
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203626
    Next
        
    '******************* registro E510 *******************

    For Each objSpedE510 In colSpedE510

        objSpedFiscal.sTipo = "E510"
        objSpedRegistroE510.sTipo = objSpedFiscal.sTipo
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroE510.iCount = objSpedRegistroE510.iCount + 1
        
        If left(objSpedE510.sCFOP, 1) = "5" Or left(objSpedE510.sCFOP, 1) = "6" Then
            dValDebIPI = dValDebIPI + objSpedE510.dValorIPI
        ElseIf left(objSpedE510.sCFOP, 1) = "1" Or left(objSpedE510.sCFOP, 1) = "2" Or left(objSpedE510.sCFOP, 1) = "3" Then
            dValCredIPI = dValCredIPI + objSpedE510.dValorIPI
        End If
        
        objSpedFiscal.sRegistro = "|E510|" & objSpedE510.sCFOP & "|" & Format(objSpedE510.iCSTIPI, "00") & "|" & Format(objSpedE510.dValorContabilIPI, "Fixed") & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(objSpedE510.dValorBCIPI, "Fixed") & "|" & Format(objSpedE510.dValorIPI, "Fixed") & "|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203627
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqE500)
        If lErro <> AD_SQL_SUCESSO Then gError 203628

    Next

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E510 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E510:

   SpedFiscal_E510 = gErr

   Select Case gErr
    
        Case 203626
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203627
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203628
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203629)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_E530(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE530 As ClassSpedRegistro, ByVal objFilialEmpresa As AdmFiliais, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal lSeqPai As Long, dTotAjDebitos As Double, dTotAjCreditos As Double) As Long
        
Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim tRegApuracaoItens As typeRegApuracaoItens
Dim iRegApuracaoIPI As Integer
Dim iRegApuracaoIPIAtual As Integer
Dim iIndAj As Integer

On Error GoTo Erro_SpedFiscal_E530
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203630
    Next
    
    tRegApuracaoItens.sDescricao = String(STRING_DESCRICAO_APURACAO, 0)
        
    'Verifica se o Item de apuração ICMS está cadastrada
    lErro = Comando_Executar(alComando(2), "SELECT Valor, R.Descricao, RegApuracaoIPI FROM RegApuracaoIPIItem AS R, TiposRegApuracaoIPI WHERE TipoReg = Codigo AND Data >= ? And Data <= ? AND FilialEmpresa = ? AND (RegApuracaoIPI = 1 OR RegApuracaoIPI = 199 OR RegApuracaoIPI = 99 OR RegApuracaoIPI = 101)  ORDER BY RegApuracaoIPI ", _
    tRegApuracaoItens.dValor, tRegApuracaoItens.sDescricao, iRegApuracaoIPI, dtDataIni, dtDataFim, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 203631
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203632

    Do While lErro = AD_SQL_SUCESSO
        
        If iRegApuracaoIPI = IPI_OUTROS_DEBITOS Or iRegApuracaoIPI = IPI_ESTORNO_CREDITOS Then
            iIndAj = 0
            dTotAjDebitos = dTotAjDebitos + tRegApuracaoItens.dValor
        Else
            iIndAj = 1
            dTotAjCreditos = dTotAjCreditos + tRegApuracaoItens.dValor
        End If
        
        
        '******************* registro E530 *******************
    
        objSpedFiscal.sTipo = "E530"
        objSpedRegistroE530.sTipo = objSpedFiscal.sTipo
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroE530.iCount = objSpedRegistroE530.iCount + 1
        
        'IND_DOC Indicador da origem do documento vinculado ao ajuste: (passou a ser obrigatório)
'        0 - Processo Judicial;
'        1 - Processo Administrativo;
'        2 - PER/DCOMP;
'        3 – Documento Fiscal
'        9 – Outros.
        If giVersao >= 12 Then
            objSpedFiscal.sRegistro = "|E530|" & iIndAj & "|" & Format(tRegApuracaoItens.dValor, "fixed") & "|" & Format(iRegApuracaoIPI, "000") & "|9||" & SFF_Texto(tRegApuracaoItens.sDescricao) & "|"
        Else
            objSpedFiscal.sRegistro = "|E530|" & iIndAj & "|" & Format(tRegApuracaoItens.dValor, "fixed") & "|" & Format(iRegApuracaoIPI, "000") & "|||" & SFF_Texto(tRegApuracaoItens.sDescricao) & "|"
        End If
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203633
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
        If lErro <> AD_SQL_SUCESSO Then gError 203634

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203635

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E530 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E530:

   SpedFiscal_E530 = gErr

   Select Case gErr
    
        Case 203630
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203631, 203632, 203635
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPIITEM", gErr)
    
        Case 203633
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203634
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203635)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_H010(ByVal iFilialEmpresa As Integer, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro0200 As ClassSpedRegistro, ByVal objSpedRegistro0205 As ClassSpedRegistro, ByVal objSpedRegistro0150 As ClassSpedRegistro, ByVal objSpedRegistro0175 As ClassSpedRegistro, ByVal objSpedRegistroH010 As ClassSpedRegistro, ByVal objSpedRegistro0190 As ClassSpedRegistro, ByVal objSpedRegistro0220 As ClassSpedRegistro, ByVal colUM As Collection, dValorTotalEst As Double, ByVal objTela As Object, ByVal dtDataInvAux As Date, ByVal iFiltroNatureza As Integer, ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
'Dim sComandoSQL As String
'Dim iAno As Integer
'Dim tSldMesEst As typeSldMesEst
'Dim tProduto As typeProduto
Dim iIndice As Integer
Dim sUMResult As String
'Dim tSldMesEst1 As typeSldMesEst1
'Dim tSldMesEst2 As typeSldMesEst2
Dim alComando(0 To 25) As Long
Dim lCliente As Long
Dim iFilialCli As Integer
Dim lFornecedor As Long
Dim iFilialForn As Integer
Dim dQuantInicial3 As Double
Dim dValorInicial3 As Double
Dim dQuantInicial As Double
Dim dValorInicial As Double
'Dim iAlmoxPadrao As Integer
'Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sCodParticipante As String
Dim lSeq0200 As Long
'Dim vbMsgRes As VbMsgBoxResult
Dim dtDataInv As Date
Dim iTipoReg As Integer, sContaContabil As String, sProd As String, sUM As String, dQuantidade As Double, dCusto As Double
Dim dQtdeTotalTerc As Double, iClasseUM As Integer, dCustoUnit As Double
Dim dtDataAux As Date, objEstTerc As ClassEstoqueTerc, sFiltro As String

On Error GoTo Erro_SpedFiscal_H010

    dValorTotalEst = 0
    
'    iAno = Year(dtDataIni) - 1
'    dtDataInv = StrParaDate("31/12/" & CStr(iAno))

    If dtDataInvAux = DATA_NULA Then
        dtDataAux = DateAdd("m", -1, dtDataIni)  'Retorna um mês
        dtDataAux = DateAdd("d", -Day(dtDataAux), dtDataAux) 'Retorna para o último dia do mês anterior
    Else
        dtDataAux = dtDataInvAux
    End If
    dtDataInv = dtDataAux

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203636
    Next
    
    sFiltro = ""
    If iFiltroNatureza = MARCADO Then
        sFiltro = " AND P.Natureza IN (1,2,3,4,5,9) "
    End If
   
    sContaContabil = String(STRING_CONTA, 0)
    sUM = String(STRING_UM_SIGLA, 0)
    sProd = String(STRING_PRODUTO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT I.TipoRegInventario, I.Produto, I.ContaContabil, I.SiglaUmEstoque, I.Quantidade, I.Custo, P.ClasseUM FROM RegInventarioRel AS I, Produtos AS P WHERE P.Codigo = I.Produto AND I.Data = ? AND I.FilialEmpresa = ? AND (I.Quantidade > ? ) " & sFiltro, iTipoReg, sProd, sContaContabil, sUM, dQuantidade, dCusto, iClasseUM, dtDataInv, iFilialEmpresa, QTDE_ESTOQUE_DELTA2)
    If lErro <> AD_SQL_SUCESSO Then gError 203637

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203638

    Do While lErro = AD_SQL_SUCESSO
    
        lErro = SpedFiscal_0200(sProd, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, lSeq0200, objFisSpedAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = SpedFiscal_0190(sUM, iClasseUM, objSpedFiscal, objSpedRegistro0190, objSpedRegistro0220, colUM, sUMResult, sProd, lSeq0200)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        objSpedFiscal.sTipo = "H010"
        objSpedRegistroH010.sTipo = objSpedFiscal.sTipo
    
'        Indicador de propriedade/posse do item:
'        0- Item de propriedade do informante e em seu poder;
'        1- Item de propriedade do informante em posse de terceiros;
'        2- Item de propriedade de terceiros em posse do informante
    
        Select Case iTipoReg
        
            Case 1 'Saldo Nosso na Própria Empresa
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistroH010.iCount = objSpedRegistroH010.iCount + 1
                
                dCustoUnit = 0
                If Abs(dQuantidade) < QTDE_ESTOQUE_DELTA Then
                    dQuantidade = 0
                End If
                If Abs(dCusto) < DELTA_VALORMONETARIO Then
                    dCusto = 0
                End If
                If dQuantidade < 0 Then
                    dQuantidade = 0
                End If
                If dCusto < 0 Then
                    dCusto = 0
                End If
                If dQuantidade > 0 Then
                    dCustoUnit = dCusto / dQuantidade
                End If
                
                objSpedFiscal.sRegistro = "|H010|" & SFF_Texto(sProd) & "|" & sUMResult & "|" & Format(dQuantidade, "###########0.000") & "|"
                If dQuantidade <> 0 Then
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dCustoUnit, "########0.000000") & "|"
                Else
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "0|"
                End If
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dCusto, "fixed") & "|0|||" & sContaContabil & "|"
        
                If giVersao >= 9 Then objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "|"
        
                dValorTotalEst = dValorTotalEst + dCusto
        
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203643
        
                lErro = Comando_Executar(alComando(2), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
                If lErro <> AD_SQL_SUCESSO Then gError 203644
            
            Case 2 'Saldo Nosso em Poder de Terceiros
                        
                dCustoUnit = 0
                If Abs(dQuantidade) < QTDE_ESTOQUE_DELTA Then
                    dQuantidade = 0
                End If
                If Abs(dCusto) < DELTA_VALORMONETARIO Then
                    dCusto = 0
                End If
                If dQuantidade < 0 Then
                    dQuantidade = 0
                End If
                If dCusto < 0 Then
                    dCusto = 0
                End If
                If dQuantidade > 0 Then
                    dCustoUnit = dCusto / dQuantidade
                End If
                
                dQtdeTotalTerc = 0
'                lErro = Comando_Executar(alComando(1), "SELECT Cliente, FilialCli, Fornecedor, FilialForn, QuantInicialConsig + QuantInicialDemo + QuantInicialConserto + QuantInicialOutros + QuantInicialBenef, ValorInicialConsig+ValorInicialDemo+ValorInicialConserto+ValorInicialOutros+ValorInicialBenef FROM SldMesEst2Terc WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? AND ((Cliente <> 0 OR Fornecedor <> 0) AND (FilialCli <> 0 OR FilialForn <> 0))", _
'                    lCliente, iFilialCli, lFornecedor, iFilialForn, dQuantInicial, dValorInicial, iAno + 1, iFilialEmpresa, sProd)
'                If lErro <> AD_SQL_SUCESSO Then gError 203652

                lErro = Comando_Executar(alComando(1), "SELECT Cliente, FilialCli, Fornecedor, FilialForn FROM SldMesEst2Terc WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? AND ((Cliente <> 0 OR Fornecedor <> 0) AND (FilialCli <> 0 OR FilialForn <> 0))", _
                    lCliente, iFilialCli, lFornecedor, iFilialForn, Year(dtDataInv), iFilialEmpresa, sProd)
                If lErro <> AD_SQL_SUCESSO Then gError 203652
        
                lErro = Comando_BuscarPrimeiro(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203653
        
                Do While lErro = AD_SQL_SUCESSO
                
                    dQuantInicial = 0
                
                    For iIndice = 2 To 6
                    
                        Set objEstTerc = New ClassEstoqueTerc
                        
                        objEstTerc.dtData = dtDataInv
                        objEstTerc.iEscaninho = iIndice '2 = Conserto; 3 = Consig; 4 = Demo; 5 = Outros; 6 = Beneficiamento
                        objEstTerc.iFilialEmpresa = iFilialEmpresa
                        objEstTerc.sProduto = sProd
                        
                        If lCliente > 0 Then
                            objEstTerc.iTipoCliForn = TIPO_TERC_CLIENTE
                            objEstTerc.lCliForn = lCliente
                            objEstTerc.iFilial = iFilialCli
                        Else
                            objEstTerc.iTipoCliForn = TIPO_TERC_FORNECEDOR
                            objEstTerc.lCliForn = lFornecedor
                            objEstTerc.iFilial = iFilialForn
                        End If
                        
                        lErro = CF("EstoqueTerc_Le_Saldo", objEstTerc, alComando(5 + (iIndice - 2)), alComando(10 + (iIndice - 2)))
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                        dQuantInicial = dQuantInicial + objEstTerc.dQuantidade
                        
                    Next
        
                    'se é um totalizador de cliente
                    If lCliente > 0 Then
                        lErro = SpedFiscal_0150_Cliente(lCliente, iFilialCli, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    Else
                        lErro = SpedFiscal_0150_Fornecedor(lFornecedor, iFilialForn, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    End If
                    objSpedFiscal.sTipo = "H010"
                    
                    If dQuantInicial < -QTDE_ESTOQUE_DELTA Then gError 211102 'Não pode ter quantidade negativa
                    If Abs(dQuantInicial) < QTDE_ESTOQUE_DELTA Then
                        dQuantInicial = 0
                    End If
                    If Abs(dValorInicial) < DELTA_VALORMONETARIO Then
                        dValorInicial = 0
                    End If
                    If dQuantInicial < 0 Then
                        dQuantInicial = 0
                    End If
                    If dValorInicial < 0 Then
                        dValorInicial = 0
                    End If
                    
                    dQtdeTotalTerc = dQtdeTotalTerc + dQuantInicial
        
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistroH010.iCount = objSpedRegistroH010.iCount + 1
        
                    objSpedFiscal.sRegistro = "|H010|" & SFF_Texto(sProd) & "|" & sUMResult & "|" & Format(dQuantInicial, "###########0.000") & "|"
                    If dQuantInicial <> 0 Then
                        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dCustoUnit, "########0.000000") & "|"
                    Else
                        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "0|"
                    End If
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dCustoUnit * dQuantInicial, "fixed") & "|1|" & sCodParticipante & "||" & sContaContabil & "|"
        
                    If giVersao >= 9 Then objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "|"
        
                    dValorTotalEst = dValorTotalEst + (dCustoUnit * dQuantInicial)
        
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203656
        
                    lErro = Comando_Executar(alComando(2), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
                    If lErro <> AD_SQL_SUCESSO Then gError 203657
        
                    lErro = Comando_BuscarProximo(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203658
                
                Loop
                
                If Abs(dQtdeTotalTerc - dQuantidade) > QTDE_ESTOQUE_DELTA Then gError 211103 ' Nem toda quantidade de terceiro está com o terceiro especificado
        
            Case 3 'Saldo de Terceiros em Nosso Poder
            
                dCustoUnit = 0
                If Abs(dQuantidade) < QTDE_ESTOQUE_DELTA Then
                    dQuantidade = 0
                End If
                If Abs(dCusto) < DELTA_VALORMONETARIO Then
                    dCusto = 0
                End If
                If dQuantidade < 0 Then
                    dQuantidade = 0
                End If
                If dCusto < 0 Then
                    dCusto = 0
                End If
                If dQuantidade > 0 Then
                    dCustoUnit = dCusto / dQuantidade
                End If
            
                dQtdeTotalTerc = 0
'                lErro = Comando_Executar(alComando(3), "SELECT Cliente, FilialCli, Fornecedor, FilialForn, QuantInicialConsig3 + QuantInicialDemo3 + QuantInicialConserto3 + QuantInicialOutros3 + QuantInicialBenef3, ValorInicialConsig3+ValorInicialDemo3+ValorInicialConserto3+ValorInicialOutros3+ValorInicialBenef3 FROM SldMesEst1Terc WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? AND ((Cliente <> 0 OR Fornecedor <> 0) AND (FilialCli <> 0 OR FilialForn <> 0))", _
'                    lCliente, iFilialCli, lFornecedor, iFilialForn, dQuantInicial3, dValorInicial3, iAno + 1, iFilialEmpresa, sProd)
'                If lErro <> AD_SQL_SUCESSO Then gError 203652
        
                 lErro = Comando_Executar(alComando(3), "SELECT Cliente, FilialCli, Fornecedor, FilialForn FROM SldMesEst1Terc WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? AND ((Cliente <> 0 OR Fornecedor <> 0) AND (FilialCli <> 0 OR FilialForn <> 0))", _
                    lCliente, iFilialCli, lFornecedor, iFilialForn, Year(dtDataInv), iFilialEmpresa, sProd)
                If lErro <> AD_SQL_SUCESSO Then gError 203652
        
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203653
        
                Do While lErro = AD_SQL_SUCESSO
                
                    dQuantInicial3 = 0
                
                    For iIndice = 7 To 11
                    
                        Set objEstTerc = New ClassEstoqueTerc
                        
                        objEstTerc.dtData = dtDataInv
                        objEstTerc.iEscaninho = iIndice '7 = Conserto; 8 = Consig; 9 = Demo; 10 = Outros;  11 = Beneficiamento
                        objEstTerc.iFilialEmpresa = iFilialEmpresa
                        objEstTerc.sProduto = sProd
                        
                        If lCliente > 0 Then
                            objEstTerc.iTipoCliForn = TIPO_TERC_CLIENTE
                            objEstTerc.lCliForn = lCliente
                            objEstTerc.iFilial = iFilialCli
                        Else
                            objEstTerc.iTipoCliForn = TIPO_TERC_FORNECEDOR
                            objEstTerc.lCliForn = lFornecedor
                            objEstTerc.iFilial = iFilialForn
                        End If
                        
                        lErro = CF("EstoqueTerc_Le_Saldo", objEstTerc, alComando(15 + (iIndice - 7)), alComando(19 + (iIndice - 7)))
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                        dQuantInicial3 = dQuantInicial3 + objEstTerc.dQuantidade
                        
                    Next
        
                    'se é um totalizador de cliente
                    If lCliente > 0 Then
                        lErro = SpedFiscal_0150_Cliente(lCliente, iFilialCli, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    Else
                        lErro = SpedFiscal_0150_Fornecedor(lFornecedor, iFilialForn, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    End If
                    objSpedFiscal.sTipo = "H010"
                    
                    If dQuantInicial3 < -QTDE_ESTOQUE_DELTA Then gError 211104 'Não pode ter quantidade negativa
                    If Abs(dQuantInicial3) < QTDE_ESTOQUE_DELTA Then
                        dQuantInicial3 = 0
                    End If
                    If Abs(dQuantInicial3) < DELTA_VALORMONETARIO Then
                        dValorInicial3 = 0
                    End If
                    If dQuantInicial3 < 0 Then
                        dQuantInicial3 = 0
                    End If
                    If dValorInicial3 < 0 Then
                        dValorInicial3 = 0
                    End If
                    
                    dQtdeTotalTerc = dQtdeTotalTerc + dQuantInicial3
        
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objSpedRegistroH010.iCount = objSpedRegistroH010.iCount + 1
        
                    objSpedFiscal.sRegistro = "|H010|" & SFF_Texto(sProd) & "|" & sUMResult & "|" & Format(dQuantInicial3, "###########0.000") & "|"
                    If dQuantInicial3 <> 0 Then
                        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dCustoUnit, "########0.000000") & "|"
                    Else
                        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "0|"
                    End If
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dCustoUnit * dQuantInicial3, "fixed") & "|2|" & sCodParticipante & "||" & sContaContabil & "|"
        
                    If giVersao >= 9 Then objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "|"
        
                    dValorTotalEst = dValorTotalEst + (dCustoUnit * dQuantInicial3)
        
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203656
        
                    lErro = Comando_Executar(alComando(2), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
                    If lErro <> AD_SQL_SUCESSO Then gError 203657
        
                    lErro = Comando_BuscarProximo(alComando(3))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203658
                
                Loop
                
                If Abs(dQtdeTotalTerc - dQuantidade) > QTDE_ESTOQUE_DELTA Then gError 211105 ' Nem toda quantidade de terceiro está com o terceiro especificado
            
        End Select
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203638
    
    Loop

'    iAno = Year(dtDataIni) - 1
'    tProduto.sCodigo = String(STRING_PRODUTO, 0)
'    tProduto.sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
'
'
'    '----------------------------- Leitura de SldMesEst ------------------
'    'Lê as quantidades Inciais para o Produto
'    sComandoSQL = "SELECT Produto, SiglaUMEstoque , ClasseUM, QuantInicial, ValorInicial, QuantInicialRecIndl, ValorInicialRecIndl FROM SldMesEst, Produtos WHERE Ano = ? AND FilialEmpresa = ? AND SldMesEst.Produto = Produtos.Codigo ORDER BY Produto"
'
'
'    With tSldMesEst
'        lErro = Comando_Executar(alComando(2), sComandoSQL, tProduto.sCodigo, tProduto.sSiglaUMEstoque, tProduto.iClasseUM, .dQuantInicial, .dValorInicial, _
'            .dQuantInicialRecIndl, .dValorInicialRecIndl, iAno, iFilialEmpresa)
'        If lErro <> AD_SQL_SUCESSO Then gError 203637
'    End With
'
'    lErro = Comando_BuscarPrimeiro(alComando(2))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203638
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        'Lê o Almoxarifado Padrão do Produto em questão
'        lErro = CF("AlmoxarifadoPadrao_Le", iFilialEmpresa, tProduto.sCodigo, iAlmoxPadrao)
'        If lErro <> SUCESSO And lErro <> 23796 Then gError 203639
'
'
'        objEstoqueProduto.sProduto = tProduto.sCodigo
'        objEstoqueProduto.iAlmoxarifado = iAlmoxPadrao
'
'        'Lê o EstoqueProduto para ler a conta de Estoque
'        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
'        If lErro <> SUCESSO And lErro <> 21306 Then gError 203640
'
'
'        lErro = SpedFiscal_0200(tProduto.sCodigo, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, lSeq0200)
'        If lErro <> SUCESSO Then gError 203641
'
'        lErro = SpedFiscal_0190(tProduto.sSiglaUMEstoque, tProduto.iClasseUM, objSpedFiscal, objSpedRegistro0190, objSpedRegistro0220, colUM, sUMResult, tProduto.sCodigo, lSeq0200)
'        If lErro <> SUCESSO Then gError 203642
'
'
'        '******************* registro H010 *******************
'
'        objSpedFiscal.sTipo = "H010"
'        objSpedRegistroH010.sTipo = objSpedFiscal.sTipo
'
'        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
'        objSpedRegistroH010.iCount = objSpedRegistroH010.iCount + 1
'
'        objSpedFiscal.sRegistro = "|H010|" & tProduto.sCodigo & "|" & sUMResult & "|" & Format(tSldMesEst.dQuantInicial - tSldMesEst.dQuantInicialRecIndl, "###########0.000") & "|"
'        If (tSldMesEst.dQuantInicial - tSldMesEst.dQuantInicialRecIndl) <> 0 Then
'            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format((tSldMesEst.dValorInicial - tSldMesEst.dValorInicialRecIndl) / (tSldMesEst.dQuantInicial - tSldMesEst.dQuantInicialRecIndl), "########0.000000") & "|"
'        Else
'            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "0|"
'        End If
'
'        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tSldMesEst.dValorInicial - tSldMesEst.dValorInicialRecIndl, "fixed") & "|0|||" & objEstoqueProduto.sContaContabil & "|"
'
'        dValorTotalEst = dValorTotalEst + tSldMesEst.dValorInicial - tSldMesEst.dValorInicialRecIndl
'
'        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203643
'
'        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
'        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
'        If lErro <> AD_SQL_SUCESSO Then gError 203644
'
'        sComandoSQL = "SELECT Cliente, FilialCli, Fornecedor, FilialForn, QuantInicialConsig3 + QuantInicialDemo3 + QuantInicialConserto3 + QuantInicialOutros3 + QuantInicialBenef3, ValorInicialConsig3+ValorInicialDemo3+ValorInicialConserto3+ValorInicialOutros3+ValorInicialBenef3 FROM SldMesEst1Terc WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ?"
'
'        lErro = Comando_Executar(alComando(3), sComandoSQL, _
'            lCliente, iFilialCli, lFornecedor, iFilialForn, dQuantInicial3, dValorInicial3, iAno, iFilialEmpresa, tProduto.sCodigo)
'        If lErro <> AD_SQL_SUCESSO Then gError 203645
'
'        lErro = Comando_BuscarPrimeiro(alComando(3))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203646
'
'        Do While lErro = AD_SQL_SUCESSO
'
'            'se é um totalizador de cliente
'            If lCliente > 0 Then
'
'                lErro = SpedFiscal_0150_Cliente(lCliente, iFilialCli, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
'                If lErro <> SUCESSO Then gError 203647
'
'            Else
'
'                lErro = SpedFiscal_0150_Fornecedor(lFornecedor, iFilialForn, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
'                If lErro <> SUCESSO Then gError 203648
'
'            End If
'
'
'            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
'            objSpedRegistroH010.iCount = objSpedRegistroH010.iCount + 1
'
'            objSpedFiscal.sRegistro = "|H010|" & tProduto.sCodigo & "|" & sUMResult & "|" & Format(dQuantInicial3, "###########0.000") & "|"
'            If dQuantInicial3 <> 0 Then
'                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dValorInicial3 / dQuantInicial3, "########0.000000") & "|"
'            Else
'                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "0|"
'            End If
'
'            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dValorInicial3, "fixed") & "|2|" & sCodParticipante & "||" & objEstoqueProduto.sContaContabil & "|"
'
'            dValorTotalEst = dValorTotalEst + dValorInicial3
'
'            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203649
'
'            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
'            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
'            If lErro <> AD_SQL_SUCESSO Then gError 203650
'
'            lErro = Comando_BuscarProximo(alComando(3))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203651
'
'
'        Loop
'
'        sComandoSQL = "SELECT Cliente, FilialCli, Fornecedor, FilialForn, QuantInicialConsig + QuantInicialDemo + QuantInicialConserto + QuantInicialOutros + QuantInicialBenef, ValorInicialConsig+ValorInicialDemo+ValorInicialConserto+ValorInicialOutros+ValorInicialBenef FROM SldMesEst2Terc WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ?"
'
'        lErro = Comando_Executar(alComando(4), sComandoSQL, _
'            lCliente, iFilialCli, lFornecedor, iFilialForn, dQuantInicial, dValorInicial, iAno, iFilialEmpresa, tProduto.sCodigo)
'        If lErro <> AD_SQL_SUCESSO Then gError 203652
'
'        lErro = Comando_BuscarPrimeiro(alComando(4))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203653
'
'        Do While lErro = AD_SQL_SUCESSO
'
'            'se é um totalizador de cliente
'            If lCliente > 0 Then
'
'                lErro = SpedFiscal_0150_Cliente(lCliente, iFilialCli, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
'                If lErro <> SUCESSO Then gError 203654
'
'            Else
'
'                lErro = SpedFiscal_0150_Fornecedor(lFornecedor, iFilialForn, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
'                If lErro <> SUCESSO Then gError 203655
'
'            End If
'
'
'            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
'            objSpedRegistroH010.iCount = objSpedRegistroH010.iCount + 1
'
'            objSpedFiscal.sRegistro = "|H010|" & tProduto.sCodigo & "|" & sUMResult & "|" & Format(dQuantInicial, "###########0.000") & "|"
'            If dQuantInicial <> 0 Then
'                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dValorInicial / dQuantInicial, "########0.000000") & "|"
'            Else
'                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "0|"
'            End If
'            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dValorInicial, "fixed") & "|1|" & sCodParticipante & "||" & objEstoqueProduto.sContaContabil & "|"
'
'            dValorTotalEst = dValorTotalEst + dValorInicial
'
'            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203656
'
'            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
'            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
'            If lErro <> AD_SQL_SUCESSO Then gError 203657
'
'            lErro = Comando_BuscarProximo(alComando(4))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203658
'
'        Loop
'
'        lErro = objTela.Processa_Item
'        If lErro <> SUCESSO Then gError 204135
'
'        lErro = Comando_BuscarProximo(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203659
'
'
'    Loop
        
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
        
    SpedFiscal_H010 = SUCESSO

    Exit Function

Erro_SpedFiscal_H010:

    SpedFiscal_H010 = gErr

    Select Case gErr
        
        Case 203636
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203637, 203638, 203659
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST1", gErr, Year(dtDataInv), iFilialEmpresa)
    
        Case 203639 To 203642, 203647, 203648, 203654, 203655
        
        Case 203643, 203649, 203656
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203644, 203650, 203657
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case 203645, 203646, 203651
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTTERC1", gErr)
        
        Case 203652, 203653, 203658
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTTERC2", gErr)
        
        Case 204135
        
        Case 211102
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST2TERC_NEGATIVO", gErr, Format(dtDataInv, "dd/mm/yyyy"), iFilialEmpresa, sProd, IIf(lCliente > 0, "Cliente", "Fornecedor"), IIf(lCliente > 0, lCliente, lFornecedor), IIf(lCliente > 0, iFilialCli, iFilialForn))
            Resume Next
        
        Case 211103
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDTERC2_DIF_QTDE_REGINV", gErr, Format(dtDataInv, "dd/mm/yyyy"), iFilialEmpresa, sProd, Formata_Estoque(dQtdeTotalTerc), Formata_Estoque(dQuantidade))
            Resume Next
        
        Case 211104
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST1TERC_NEGATIVO", gErr, Format(dtDataInv, "dd/mm/yyyy"), iFilialEmpresa, sProd, IIf(lCliente > 0, "Cliente", "Fornecedor"), IIf(lCliente > 0, lCliente, lFornecedor), IIf(lCliente > 0, iFilialCli, iFilialForn))
            Resume Next
        
        Case 211105
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDTERC1_DIF_QTDE_REGINV", gErr, Format(dtDataInv, "dd/mm/yyyy"), iFilialEmpresa, sProd, Formata_Estoque(dQtdeTotalTerc), Formata_Estoque(dQuantidade))
            Resume Next
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203660)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function SpedFiscal_9900(ByVal objSpedFiscal As ClassSpedFiscal, ByVal colSpedRegistro As Collection) As Long ', ByVal objEFD As ClassEFDPisCofinsSel) As Long
        
Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_SpedFiscal_9900
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203661
    Next
        
    '******************* registro 9900 *******************

    objSpedFiscal.sTipo = "9900"

    For Each objSpedRegistro In colSpedRegistro

        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1

        objSpedFiscal.sRegistro = "|9900|" & objSpedRegistro.sTipo & "|" & objSpedRegistro.iCount & "|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203662
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203663

    Next
    
'    'Bloco B
'    For Each objSpedRegistro In objEFD.colSpedRegistro
'
'        Call objEFD.Inicia_Reg
'        Call objEFD.Reg_Campo_add(1, "REG", "9900")
'        Call objEFD.Reg_Campo_add(2, "REG_BLC", objSpedRegistro.sTipo)
'        Call objEFD.Reg_Campo_add(3, "QTD_REG_BLC", CStr(objSpedRegistro.lCount))
'        Call objEFD.Finaliza_Reg
'
'        lErro = Registro_Insere_BD(objEFD)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    Next

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_9900 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_9900:

   SpedFiscal_9900 = gErr

   Select Case gErr
    
        Case 203661
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203662
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203663
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203664)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_Gravar_Arq(ByVal sDiretorio As String, ByVal objSpedFiscal As ClassSpedFiscal, ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim sRegistro As String
Dim iIndice As Integer
Dim alComando(1 To 93) As Long
Dim lSeq As Long
Dim sTipo As String
Dim lSeqC100 As Long
Dim lSeqC110 As Long
Dim lSeqC140 As Long
Dim lSeqC170 As Long
Dim lSeqC300 As Long
Dim lSeqC320 As Long
Dim lSeqC321 As Long
Dim lSeqC400 As Long
Dim lSeqC405 As Long
Dim lSeqC420 As Long
Dim lSeqC460 As Long
Dim lSeqC500 As Long
Dim lSeqD100 As Long
Dim lSeqD500 As Long
Dim lSeqE100 As Long
Dim lSeqE110 As Long
Dim lSeqE200 As Long
Dim lSeqE210 As Long
Dim lSeqE500 As Long
Dim lSeqE520 As Long, lSeqG110 As Long, lSeqG125 As Long, lSeqG130 As Long
Dim lSeqE300 As Long
Dim lFreeFile As Long

On Error GoTo Erro_SpedFiscal_Gravar_Arq

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203665
    Next
    
    lFreeFile = FreeFile()

    Open sDiretorio For Output As lFreeFile
    
    '********* registros 0000 a 0100 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo >= '0000' And Tipo <= '0100'  ORDER BY Tipo", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203666
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203667


    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203668
    
    Loop


    '********* registro 0150 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(2), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0150'  ORDER BY Seq", _
    sRegistro, lSeq, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203669
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203670

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro
    
        '********* registro 0175 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(3), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0175' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 203671
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203672
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203673
        
        Loop
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203674
    
    Loop


    '********* registro 0190 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(4), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0190'  ORDER BY Seq", _
    sRegistro, lSeq, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203669
    
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203670

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro
    
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203674
    
    Loop

    '********* registro 0200 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)

    lErro = Comando_Executar(alComando(5), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0200'  ORDER BY Seq", _
    sRegistro, lSeq, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203675
    
    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203676

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro
    
        
        '********* registro 0205 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(6), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0205' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 203677
        
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203678
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203679
        
        Loop
    
        '********* registro 0210 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(7), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0210' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 203677
        
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203678
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(7))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203679
        
        Loop
        
        '********* registro 0220 ******************************
    
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(8), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0220' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 203680
        
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203681
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203682
        
        Loop
    
    
        lErro = Comando_BuscarProximo(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203683
    
    Loop

    '********* registro 0300 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)

    lErro = Comando_Executar(alComando(9), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0300'  ORDER BY Seq", _
        sRegistro, lSeq, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203675
    
    lErro = Comando_BuscarPrimeiro(alComando(9))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203676

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro
    
        
        '********* registro 0305 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(10), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0305' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 203677
        
        lErro = Comando_BuscarPrimeiro(alComando(10))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203678
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(10))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203679
        
        Loop
    
        lErro = Comando_BuscarProximo(alComando(9))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203683
    
    Loop
    
    '********* registro 0450 ******************************

    sRegistro = String(STRING_SPED_REGISTRO, 0)

    lErro = Comando_Executar(alComando(11), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0450' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203684
    
    lErro = Comando_BuscarPrimeiro(alComando(11))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203685

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(11))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203686
    
    Loop

    '********* registro 0500 ******************************

    sRegistro = String(STRING_SPED_REGISTRO, 0)

    lErro = Comando_Executar(alComando(12), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0500' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203684
    
    lErro = Comando_BuscarPrimeiro(alComando(12))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203685

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(12))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203686
    
    Loop

    '********* registro 0600 ******************************

    sRegistro = String(STRING_SPED_REGISTRO, 0)

    lErro = Comando_Executar(alComando(13), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0600' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203684
    
    lErro = Comando_BuscarPrimeiro(alComando(13))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203685

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(13))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203686
    
    Loop

    '********* registro 0990 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    sTipo = "0990"

    lErro = Comando_Executar(alComando(14), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = ? ", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, sTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 203687
    
    lErro = Comando_BuscarPrimeiro(alComando(14))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203688

    If lErro <> AD_SQL_SUCESSO Then gError 203689
    
    Print #lFreeFile, sRegistro
    
    If giVersao > 12 Then
        lErro = SpedFiscal_Gravar_BlocoB(objEFD, lFreeFile)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    '********* registro C001 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    sTipo = "C001"

    lErro = Comando_Executar(alComando(15), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = ? ", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, sTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 203690
    
    lErro = Comando_BuscarPrimeiro(alComando(15))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203691

    If lErro <> AD_SQL_SUCESSO Then gError 203692
    
    Print #lFreeFile, sRegistro

    '********* registro C100 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)

    lErro = Comando_Executar(alComando(16), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C100'  ORDER BY Seq", _
    sRegistro, lSeqC100, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203693
    
    lErro = Comando_BuscarPrimeiro(alComando(16))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203694

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro
        
         '********* registro C101 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(91), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = ? AND SeqPai = ? ORDER BY Seq", _
        sRegistro, lSeqC110, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, "C101", lSeqC100)
        If lErro <> AD_SQL_SUCESSO Then gError 203695
        
        lErro = Comando_BuscarPrimeiro(alComando(91))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203696
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
            
            lErro = Comando_BuscarProximo(alComando(91))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203712
        
        Loop
        
        '********* registro C110 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(17), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C110' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, lSeqC110, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC100)
        If lErro <> AD_SQL_SUCESSO Then gError 203695
        
        lErro = Comando_BuscarPrimeiro(alComando(17))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203696
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            '********* registro C111 ******************************
        
            sRegistro = String(STRING_SPED_REGISTRO, 0)
        
            lErro = Comando_Executar(alComando(18), "SELECT Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C111' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC110)
            If lErro <> AD_SQL_SUCESSO Then gError 203697
            
            lErro = Comando_BuscarPrimeiro(alComando(18))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203698
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(18))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203699
            
            Loop
    
            '********* registro C112 ******************************
        
            sRegistro = String(STRING_SPED_REGISTRO, 0)
        
            lErro = Comando_Executar(alComando(19), "SELECT Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C112' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC110)
            If lErro <> AD_SQL_SUCESSO Then gError 203700
            
            lErro = Comando_BuscarPrimeiro(alComando(19))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203701
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(19))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203702
            
            Loop
    
    
            '********* registro C113 ******************************
        
            sRegistro = String(STRING_SPED_REGISTRO, 0)
        
            lErro = Comando_Executar(alComando(20), "SELECT Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C113' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC110)
            If lErro <> AD_SQL_SUCESSO Then gError 203703
            
            lErro = Comando_BuscarPrimeiro(alComando(20))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203704
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(20))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203705
            
            Loop
    
            '********* registro C114 ******************************
        
            sRegistro = String(STRING_SPED_REGISTRO, 0)
        
            lErro = Comando_Executar(alComando(21), "SELECT Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C114' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC110)
            If lErro <> AD_SQL_SUCESSO Then gError 203706
            
            lErro = Comando_BuscarPrimeiro(alComando(21))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203707
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(21))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203708
            
            Loop
    
    
            '********* registro C115 ******************************
        
            sRegistro = String(STRING_SPED_REGISTRO, 0)
        
            lErro = Comando_Executar(alComando(22), "SELECT Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C115' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC110)
            If lErro <> AD_SQL_SUCESSO Then gError 203709
            
            lErro = Comando_BuscarPrimeiro(alComando(22))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203710
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(22))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203711
            
            Loop
    
            lErro = Comando_BuscarProximo(alComando(17))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203712
        
        Loop
    
        '********* registro C120 ******************************
    
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(23), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C120' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC100)
        If lErro <> AD_SQL_SUCESSO Then gError 203713
        
        lErro = Comando_BuscarPrimeiro(alComando(23))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203714
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(23))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203715
        
        Loop
    
    
        '********* registro C130 ******************************
    
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(24), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C130' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC100)
        If lErro <> AD_SQL_SUCESSO Then gError 203716
        
        lErro = Comando_BuscarPrimeiro(alComando(24))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203717
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(24))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203718
        
        Loop
    
    
        '********* registro C140 ******************************
    
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(25), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C140' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, lSeqC140, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC100)
        If lErro <> AD_SQL_SUCESSO Then gError 203719
        
        lErro = Comando_BuscarPrimeiro(alComando(25))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203720
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            '********* registro C141 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
        
            lErro = Comando_Executar(alComando(26), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C141' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC140)
            If lErro <> AD_SQL_SUCESSO Then gError 203721
            
            lErro = Comando_BuscarPrimeiro(alComando(26))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203722
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(26))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203723
            
            Loop
    
    
            lErro = Comando_BuscarProximo(alComando(25))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203724
        
        Loop
    
        '********* registro C160 ******************************
    
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(27), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C160' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC100)
        If lErro <> AD_SQL_SUCESSO Then gError 203725
        
        lErro = Comando_BuscarPrimeiro(alComando(27))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203726
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(27))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203727
        
        Loop
    
        '********* registro C170 ******************************
    
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(28), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C170' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, lSeqC170, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC100)
        If lErro <> AD_SQL_SUCESSO Then gError 203728
        
        lErro = Comando_BuscarPrimeiro(alComando(28))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203729
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            '********* registro C172 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
        
            lErro = Comando_Executar(alComando(29), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C172' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC170)
            If lErro <> AD_SQL_SUCESSO Then gError 203730
            
            lErro = Comando_BuscarPrimeiro(alComando(29))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203731
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(29))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203732
            
            Loop
    
    
            lErro = Comando_BuscarProximo(alComando(28))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203733
        
        Loop
    
        '********* registro C190 ******************************
    
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(30), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo IN ('C190','C191') AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC100)
        If lErro <> AD_SQL_SUCESSO Then gError 203734
        
        lErro = Comando_BuscarPrimeiro(alComando(30))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203735
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(30))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203736
        
        Loop
    
        '********* registro C195 ******************************
    
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(31), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C195' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC100)
        If lErro <> AD_SQL_SUCESSO Then gError 203737
        
        lErro = Comando_BuscarPrimeiro(alComando(31))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203738
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(31))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203739
        
        Loop
    
    
        lErro = Comando_BuscarProximo(alComando(16))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203740
    
    Loop

    '********* registro C300 ******************************

    sRegistro = String(STRING_SPED_REGISTRO, 0)

    lErro = Comando_Executar(alComando(32), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C300' ORDER BY Seq", _
    sRegistro, lSeqC300, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203741
    
    lErro = Comando_BuscarPrimeiro(alComando(32))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203742

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro C320 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(33), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C320' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, lSeqC320, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC300)
        If lErro <> AD_SQL_SUCESSO Then gError 203743
        
        lErro = Comando_BuscarPrimeiro(alComando(33))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203744
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            '********* registro C321 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
            
            lErro = Comando_Executar(alComando(34), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C321' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, lSeqC321, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC320)
            If lErro <> AD_SQL_SUCESSO Then gError 203745
            
            lErro = Comando_BuscarPrimeiro(alComando(34))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203746
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(34))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203750
            
            Loop
    
    
            lErro = Comando_BuscarProximo(alComando(33))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203760
        
        Loop

        lErro = Comando_BuscarProximo(alComando(32))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203761
    
    Loop
    
    '********* registro C400 ******************************

    sRegistro = String(STRING_SPED_REGISTRO, 0)

    lErro = Comando_Executar(alComando(35), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C400' ORDER BY Seq", _
    sRegistro, lSeqC400, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203741
    
    lErro = Comando_BuscarPrimeiro(alComando(35))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203742

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro C405 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(36), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C405' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, lSeqC405, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC400)
        If lErro <> AD_SQL_SUCESSO Then gError 203743
        
        lErro = Comando_BuscarPrimeiro(alComando(36))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203744
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            '********* registro C420 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
            
            lErro = Comando_Executar(alComando(37), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C420' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, lSeqC420, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC405)
            If lErro <> AD_SQL_SUCESSO Then gError 203745
            
            lErro = Comando_BuscarPrimeiro(alComando(37))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203746
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                '********* registro C425 ******************************
                sRegistro = String(STRING_SPED_REGISTRO, 0)
                
                lErro = Comando_Executar(alComando(38), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C425' AND SeqPai = ? ORDER BY Seq", _
                sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC420)
                If lErro <> AD_SQL_SUCESSO Then gError 203747
                
                lErro = Comando_BuscarPrimeiro(alComando(38))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203748
            
                Do While lErro = AD_SQL_SUCESSO
                
                    Print #lFreeFile, sRegistro
            
                    lErro = Comando_BuscarProximo(alComando(38))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203749
                
                Loop
        
        
                lErro = Comando_BuscarProximo(alComando(37))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203750
            
            Loop
    
            '********* registro C460 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
            
            lErro = Comando_Executar(alComando(39), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C460' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, lSeqC460, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC405)
            If lErro <> AD_SQL_SUCESSO Then gError 203751
            
            lErro = Comando_BuscarPrimeiro(alComando(39))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203752
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                '********* registro C470 ******************************
                sRegistro = String(STRING_SPED_REGISTRO, 0)
                
                lErro = Comando_Executar(alComando(40), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C470' AND SeqPai = ? ORDER BY Seq", _
                sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC460)
                If lErro <> AD_SQL_SUCESSO Then gError 203753
                
                lErro = Comando_BuscarPrimeiro(alComando(40))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203754
            
                Do While lErro = AD_SQL_SUCESSO
                
                    Print #lFreeFile, sRegistro
            
                    lErro = Comando_BuscarProximo(alComando(40))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203755
                
                Loop
        
        
                lErro = Comando_BuscarProximo(alComando(39))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203756
            
            
            
            
            Loop
    
            '********* registro C490 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
            
            lErro = Comando_Executar(alComando(41), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C490' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, lSeqC460, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC405)
            If lErro <> AD_SQL_SUCESSO Then gError 203757
            
            lErro = Comando_BuscarPrimeiro(alComando(41))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203758
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(41))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203759
            
            Loop
    
            lErro = Comando_BuscarProximo(alComando(36))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203760
        
        Loop

        lErro = Comando_BuscarProximo(alComando(35))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203761
    
    Loop

    '********* registro C500 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(42), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C500' ORDER BY Seq", _
    sRegistro, lSeqC500, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203762
    
    lErro = Comando_BuscarPrimeiro(alComando(42))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203763

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro C590 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
        
        lErro = Comando_Executar(alComando(43), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C590' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqC500)
        If lErro <> AD_SQL_SUCESSO Then gError 203764
        
        lErro = Comando_BuscarPrimeiro(alComando(43))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203765
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(43))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203766
        
        Loop

        lErro = Comando_BuscarProximo(alComando(42))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203767
    
    Loop

    '********* registro C990 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(44), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'C990' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203768
    
    lErro = Comando_BuscarPrimeiro(alComando(44))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203769

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(44))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203770
    
    Loop

    '********* registro D001 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(45), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'D001' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203771
    
    lErro = Comando_BuscarPrimeiro(alComando(45))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203772

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(45))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203773
    
    Loop

    '********* registro D100 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(46), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'D100' ORDER BY Seq", _
    sRegistro, lSeqD100, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203774
    
    lErro = Comando_BuscarPrimeiro(alComando(46))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203775

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro D190 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
        
        lErro = Comando_Executar(alComando(47), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'D190' And SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqD100)
        If lErro <> AD_SQL_SUCESSO Then gError 203776
        
        lErro = Comando_BuscarPrimeiro(alComando(47))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203777
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(47))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203778
        
        Loop

        lErro = Comando_BuscarProximo(alComando(46))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203779
    
    Loop

    '********* registro D500 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(48), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'D500' ORDER BY Seq", _
    sRegistro, lSeqD500, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203780
    
    lErro = Comando_BuscarPrimeiro(alComando(48))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203781

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro D590 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
        
        lErro = Comando_Executar(alComando(49), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'D590' And SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqD500)
        If lErro <> AD_SQL_SUCESSO Then gError 203782
        
        lErro = Comando_BuscarPrimeiro(alComando(49))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203783
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(49))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203784
        
        Loop


        lErro = Comando_BuscarProximo(alComando(48))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203785
    
    Loop

    '********* registro D990 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(50), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'D990' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203786
    
    lErro = Comando_BuscarPrimeiro(alComando(50))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203787

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(50))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203788
    
    Loop

    '********* registro E001 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(51), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E001' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203789
    
    lErro = Comando_BuscarPrimeiro(alComando(51))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203790

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(51))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203791
    
    Loop

    '********* registro E100 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(52), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E100' ORDER BY Seq", _
    sRegistro, lSeqE100, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203792
    
    lErro = Comando_BuscarPrimeiro(alComando(52))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203793

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro E110 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)

        lErro = Comando_Executar(alComando(53), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E110' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, lSeqE110, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqE100)
        If lErro <> AD_SQL_SUCESSO Then gError 203794
        
        lErro = Comando_BuscarPrimeiro(alComando(53))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203795
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            '********* registro E111 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
    
            lErro = Comando_Executar(alComando(54), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E111' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqE110)
            If lErro <> AD_SQL_SUCESSO Then gError 203796
            
            lErro = Comando_BuscarPrimeiro(alComando(54))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203797
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(54))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203798
            
            Loop
    
    
            '********* registro E116 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
    
            lErro = Comando_Executar(alComando(55), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E116' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqE110)
            If lErro <> AD_SQL_SUCESSO Then gError 203799
            
            lErro = Comando_BuscarPrimeiro(alComando(55))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203800
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(55))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203801
            
            Loop
    
            lErro = Comando_BuscarProximo(alComando(53))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203802
        
        Loop


        lErro = Comando_BuscarProximo(alComando(52))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203803
    
    Loop


    '********* registro E200 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(56), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E200' ORDER BY Seq", _
    sRegistro, lSeqE200, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 204956
    
    lErro = Comando_BuscarPrimeiro(alComando(56))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204957

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro E210 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)

        lErro = Comando_Executar(alComando(57), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E210' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, lSeqE210, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqE200)
        If lErro <> AD_SQL_SUCESSO Then gError 204958
        
        lErro = Comando_BuscarPrimeiro(alComando(57))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204959
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            '********* registro E220 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
    
            lErro = Comando_Executar(alComando(58), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E220' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqE210)
            If lErro <> AD_SQL_SUCESSO Then gError 204960
            
            lErro = Comando_BuscarPrimeiro(alComando(58))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204961
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(58))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204962
            
            Loop
    
    
            '********* registro E250 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
    
            lErro = Comando_Executar(alComando(59), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E250' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqE210)
            If lErro <> AD_SQL_SUCESSO Then gError 204963
            
            lErro = Comando_BuscarPrimeiro(alComando(59))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204964
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(59))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204965
            
            Loop
    
            lErro = Comando_BuscarProximo(alComando(57))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204966
        
        Loop


        lErro = Comando_BuscarProximo(alComando(56))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204967
    
    Loop
    
    '********* registro E100 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(92), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E300' ORDER BY Seq", _
    sRegistro, lSeqE300, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203792
    
    lErro = Comando_BuscarPrimeiro(alComando(92))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203793

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro E310 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)

        lErro = Comando_Executar(alComando(93), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E310' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, lSeqE110, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqE300)
        If lErro <> AD_SQL_SUCESSO Then gError 203794
        
        lErro = Comando_BuscarPrimeiro(alComando(93))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203795
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
            
            lErro = Comando_BuscarProximo(alComando(93))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204966
        
        Loop

        lErro = Comando_BuscarProximo(alComando(92))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204967
    
    Loop

    '********* registro E500 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(60), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E500' ORDER BY Seq", _
    sRegistro, lSeqE500, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203804
    
    lErro = Comando_BuscarPrimeiro(alComando(60))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203805

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro E510 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)

        lErro = Comando_Executar(alComando(61), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E510' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqE500)
        If lErro <> AD_SQL_SUCESSO Then gError 203806
        
        lErro = Comando_BuscarPrimeiro(alComando(61))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203807
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
    
            lErro = Comando_BuscarProximo(alComando(61))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203808
        
        Loop


        '********* registro E520 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)

        lErro = Comando_Executar(alComando(62), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E520' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, lSeqE520, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqE500)
        If lErro <> AD_SQL_SUCESSO Then gError 203809
        
        lErro = Comando_BuscarPrimeiro(alComando(62))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203810
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            '********* registro E530 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
    
            lErro = Comando_Executar(alComando(63), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E530' AND SeqPai = ? ORDER BY Seq", _
            sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqE520)
            If lErro <> AD_SQL_SUCESSO Then gError 203811
            
            lErro = Comando_BuscarPrimeiro(alComando(63))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203812
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                lErro = Comando_BuscarProximo(alComando(63))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203813
            
            Loop
    
            lErro = Comando_BuscarProximo(alComando(62))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203814
        
        Loop

        lErro = Comando_BuscarProximo(alComando(60))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203815
    
    Loop

    '********* registro E990 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(64), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'E990' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203816
    
    lErro = Comando_BuscarPrimeiro(alComando(64))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203817

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(64))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203818
    
    Loop

    '********* registro G001 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(65), "SELECT Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'G001'", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203804
    
    lErro = Comando_BuscarPrimeiro(alComando(65))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203805

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro G110 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)

        lErro = Comando_Executar(alComando(66), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'G110' ORDER BY Seq", _
        sRegistro, lSeqG110, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 203809
        
        lErro = Comando_BuscarPrimeiro(alComando(66))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203810
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            '********* registro G125 ******************************
            sRegistro = String(STRING_SPED_REGISTRO, 0)
    
            lErro = Comando_Executar(alComando(67), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'G125' AND SeqPai = ? ORDER BY Seq", _
                sRegistro, lSeqG125, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqG110)
            If lErro <> AD_SQL_SUCESSO Then gError 203811
            
            lErro = Comando_BuscarPrimeiro(alComando(67))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203812
        
            Do While lErro = AD_SQL_SUCESSO
            
                Print #lFreeFile, sRegistro
        
                '********* registro G126 ******************************
                sRegistro = String(STRING_SPED_REGISTRO, 0)
        
                lErro = Comando_Executar(alComando(68), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'G126' AND SeqPai = ? ORDER BY Seq", _
                    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqG125)
                If lErro <> AD_SQL_SUCESSO Then gError 203811
                
                lErro = Comando_BuscarPrimeiro(alComando(68))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203812
            
                Do While lErro = AD_SQL_SUCESSO
                
                    Print #lFreeFile, sRegistro
            
                    lErro = Comando_BuscarProximo(alComando(68))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203813
                
                Loop
                
                '********* registro G130 ******************************
                sRegistro = String(STRING_SPED_REGISTRO, 0)
        
                lErro = Comando_Executar(alComando(69), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'G130' AND SeqPai = ? ORDER BY Seq", _
                    sRegistro, lSeqG130, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqG125)
                If lErro <> AD_SQL_SUCESSO Then gError 203811
                
                lErro = Comando_BuscarPrimeiro(alComando(69))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203812
            
                Do While lErro = AD_SQL_SUCESSO
                
                    Print #lFreeFile, sRegistro
            
                    '********* registro G140 ******************************
                    sRegistro = String(STRING_SPED_REGISTRO, 0)
            
                    lErro = Comando_Executar(alComando(70), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'G140' AND SeqPai = ? ORDER BY Seq", _
                        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeqG130)
                    If lErro <> AD_SQL_SUCESSO Then gError 203811
                    
                    lErro = Comando_BuscarPrimeiro(alComando(70))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203812
                
                    Do While lErro = AD_SQL_SUCESSO
                    
                        Print #lFreeFile, sRegistro
                
                        lErro = Comando_BuscarProximo(alComando(70))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203813
                    
                    Loop
                    
                    lErro = Comando_BuscarProximo(alComando(69))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203813
                
                Loop
                
                lErro = Comando_BuscarProximo(alComando(67))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203813
            
            Loop
    
            lErro = Comando_BuscarProximo(alComando(66))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203814
        
        Loop

        lErro = Comando_BuscarProximo(alComando(65))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203815
    
    Loop

    '********* registro G990 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(71), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'G990' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203828
    
    lErro = Comando_BuscarPrimeiro(alComando(71))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203829

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(71))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203830
    
    Loop

    '********* registro H001 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(72), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'H001' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203819
    
    lErro = Comando_BuscarPrimeiro(alComando(72))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203820

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(72))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203821
    
    Loop

    '********* registro H005 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(73), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'H005' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203822
    
    lErro = Comando_BuscarPrimeiro(alComando(73))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203823

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(73))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203824
    
    Loop

    '********* registro H010 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(74), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'H010' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203825
    
    lErro = Comando_BuscarPrimeiro(alComando(74))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203826

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(74))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203827
    
    Loop


    '********* registro H990 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(75), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'H990' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203828
    
    lErro = Comando_BuscarPrimeiro(alComando(75))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203829

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(75))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203830
    
    Loop

    '********* registro K001 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(76), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'K001' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203819
    
    lErro = Comando_BuscarPrimeiro(alComando(76))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203820

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(76))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203821
    
    Loop
    
    '********* registro K100 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(77), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'K100' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203819
    
    lErro = Comando_BuscarPrimeiro(alComando(77))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203820

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(77))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203821
    
    Loop
    
    '********* registro K200 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(78), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'K200' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203819
    
    lErro = Comando_BuscarPrimeiro(alComando(78))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203820

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(78))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203821
    
    Loop
    
    '********* registro K230 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(79), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'K230' ORDER BY Seq", _
    sRegistro, lSeq, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203819
    
    lErro = Comando_BuscarPrimeiro(alComando(79))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203820

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro K235 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(80), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'K235' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 203677
        
        lErro = Comando_BuscarPrimeiro(alComando(80))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203678
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(80))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203679
        
        Loop
        
        lErro = Comando_BuscarProximo(alComando(79))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203821
    
    Loop
    
    '********* registro K250 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(81), "SELECT  Registro, Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'K250' ORDER BY Seq", _
    sRegistro, lSeq, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203819
    
    lErro = Comando_BuscarPrimeiro(alComando(81))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203820

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        '********* registro K255 ******************************
        sRegistro = String(STRING_SPED_REGISTRO, 0)
    
        lErro = Comando_Executar(alComando(82), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'K255' AND SeqPai = ? ORDER BY Seq", _
        sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, lSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 203677
        
        lErro = Comando_BuscarPrimeiro(alComando(82))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203678
    
        Do While lErro = AD_SQL_SUCESSO
        
            Print #lFreeFile, sRegistro
    
            lErro = Comando_BuscarProximo(alComando(82))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203679
        
        Loop
        
        lErro = Comando_BuscarProximo(alComando(81))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203821
    
    Loop

    '********* registro K990 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(83), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = 'K990' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203828
    
    lErro = Comando_BuscarPrimeiro(alComando(83))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203829

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(83))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203830
    
    Loop

    '********* registro 1001 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(84), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '1001' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203831
    
    lErro = Comando_BuscarPrimeiro(alComando(84))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203832

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(84))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203833
    
    Loop


    '********* registro 1010 ******************************
    '********* registro 1100 ******************************
    '********* registro 1105 ******************************
    '********* registro 1400 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(85), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo IN ('1010','1100','1105','1400') ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 210776
    
    lErro = Comando_BuscarPrimeiro(alComando(85))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210777

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(85))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210778
    
    Loop


    '********* registro 1990 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(86), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '1990' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203834
    
    lErro = Comando_BuscarPrimeiro(alComando(86))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203835

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(86))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203836
    
    Loop

    '********* registro 9001 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(87), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '9001' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203837
    
    lErro = Comando_BuscarPrimeiro(alComando(87))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203838

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(87))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203839
    
    Loop

    '********* registro 9900 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(88), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '9900' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203840
    
    lErro = Comando_BuscarPrimeiro(alComando(88))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203841

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(88))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203842
    
    Loop

    '********* registro 9990 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(89), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '9990' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203843
    
    lErro = Comando_BuscarPrimeiro(alComando(89))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203844

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(89))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203845
    
    Loop


    '********* registro 9999 ******************************
    sRegistro = String(STRING_SPED_REGISTRO, 0)
    
    lErro = Comando_Executar(alComando(90), "SELECT  Registro FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '9999' ORDER BY Seq", _
    sRegistro, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 203846
    
    lErro = Comando_BuscarPrimeiro(alComando(90))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203847

    Do While lErro = AD_SQL_SUCESSO
    
        Print #lFreeFile, sRegistro

        lErro = Comando_BuscarProximo(alComando(90))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203848
    
    Loop

    Close lFreeFile

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_Gravar_Arq = SUCESSO
    
    Exit Function

Erro_SpedFiscal_Gravar_Arq:

    SpedFiscal_Gravar_Arq = gErr

   Select Case gErr
    
        Case 203665
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203666, 203667 To 203688, 203690, 203691, 203693 To 203848, 204956 To 204967, 210776 To 210778
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISCAL", gErr)
    
        Case 203689, 203692
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_NAO_CADASTRADO", gErr, sTipo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203849)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Close lFreeFile

End Function

Function SpedFiscal_0000(ByVal iFilialEmpresa As Long, ByVal lNumIntSped As Long, objSpedFiscal As ClassSpedFiscal, ByVal colSpedRegistro As Collection, ByVal iCodVersao As Integer, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal objFilialEmpresa As AdmFiliais) As Long
        
Dim lErro As Long
Dim objCidades As New ClassCidades
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim sRazaoSocial As String
Dim lComando As Long
Dim sIE As String

On Error GoTo Erro_SpedFiscal_0000
        

    '******** registro 0000 *****************
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203850
    Next
    
    'abre conexao com o BD
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 204147
    
    sRazaoSocial = String(STRING_NOME_EMPRESA, 0)
    
    'le todas as filiais da empresa atual
    lErro = Comando_Executar(lComando, "SELECT Nome FROM Empresas WHERE Codigo = ? ", sRazaoSocial, glEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 204148
    
    'le a primeira filial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204149
    
    Set objSpedFiscal = New ClassSpedFiscal
    
    objSpedFiscal.sRegistro = "|0000|" & Format(iCodVersao, "000") & "|0|" & Format(dtDataIni, "ddmmyyyy") & "|" & Format(dtDataFim, "ddmmyyyy") & "|"
    
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & sRazaoSocial & "|" & IIf(Len(objFilialEmpresa.sCgc) = 14, objFilialEmpresa.sCgc, "") & "|" & IIf(Len(objFilialEmpresa.sCgc) = 11, objFilialEmpresa.sCgc, "") & "|" & objFilialEmpresa.objEndereco.sSiglaEstado & "|"
    
    objCidades.sDescricao = objFilialEmpresa.objEndereco.sCidade
    
    lErro = CF("Cidade_Le_Nome", objCidades)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203851
    
    If lErro = ERRO_OBJETO_NAO_CADASTRADO Then gError 203852
    
    Call Formata_String_Numero2(objFilialEmpresa.sInscricaoEstadual, sIE)
    
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & sIE & "|" & objCidades.sCodIBGE & "|" & objFilialEmpresa.sInscricaoMunicipal & "||" & objFilialEmpresa.sSpedFiscalPerfil & "|" & IIf(objFilialEmpresa.iTipoTribAtividade = 2, 0, 1) & "|"
    
    gsCodIBGE = objCidades.sCodIBGE
    
    objSpedFiscal.sTipo = "0000"
    objSpedFiscal.iFilialEmpresa = iFilialEmpresa
    objSpedFiscal.lNumIntSped = lNumIntSped
    objSpedFiscal.lSeq = 1
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    objSpedRegistro.sTipo = objSpedFiscal.sTipo
    objSpedRegistro.iCount = objSpedRegistro.iCount + 1
    
    colSpedRegistro.Add objSpedRegistro
    
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203853
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203854

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_0000 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_0000:

   SpedFiscal_0000 = gErr

   Select Case gErr
    
        Case 203850, 204147
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203851
    
        Case 203852
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_NAO_CADASTRADA2", gErr, objCidades.sDescricao)
    
        Case 203853
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203854
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case 204148, 104149
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203855)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Function SpedFiscal_0001(objSpedFiscal As ClassSpedFiscal, ByVal colSpedRegistro As Collection) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_SpedFiscal_0001
        

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203856
    Next


    '******** registro 0001 *****************
    
    objSpedFiscal.sTipo = "0001"
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    
    objSpedFiscal.sRegistro = "|0001|0|"
    
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    objSpedRegistro.sTipo = objSpedFiscal.sTipo
    objSpedRegistro.iCount = objSpedRegistro.iCount + 1
    
    colSpedRegistro.Add objSpedRegistro
    
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203857
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203858

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_0001 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_0001:

   SpedFiscal_0001 = gErr

   Select Case gErr
    
        Case 203856
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203857
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203858
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203859)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_0002(objSpedFiscal As ClassSpedFiscal, ByVal colSpedRegistro As Collection) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_SpedFiscal_0002
        

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203856
    Next


    '******** registro 0002 *****************
    
    objSpedFiscal.sTipo = "0002"
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    
    objSpedFiscal.sRegistro = "|0002|09|"
    
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    objSpedRegistro.sTipo = objSpedFiscal.sTipo
    objSpedRegistro.iCount = objSpedRegistro.iCount + 1
    
    colSpedRegistro.Add objSpedRegistro
    
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203857
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203858

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_0002 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_0002:

   SpedFiscal_0002 = gErr

   Select Case gErr
    
        Case 203856
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203857
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203858
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203859)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function
        

Function SpedFiscal_0005(ByVal objFilialEmpresa As AdmFiliais, objSpedFiscal As ClassSpedFiscal, ByVal colSpedRegistro As Collection) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim lComando As Long
Dim iIndice As Integer
Dim sNomeReduzido As String
Dim sTelefone As String
Dim sFax As String

On Error GoTo Erro_SpedFiscal_0005
        

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203860
    Next
    '******** registro 0005 *****************
    
    objSpedFiscal.sTipo = "0005"
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    
    
    Call Formata_String_Numero2(objFilialEmpresa.objEndereco.sTelefone1, sTelefone)
    Call Formata_String_Numero2(objFilialEmpresa.objEndereco.sFax, sFax)
    
    If Len(Trim(objFilialEmpresa.sNomeReduzido)) = 0 Then
        
        'abre conexao com o BD
        lComando = Comando_AbrirExt(GL_lConexaoDic)
        If lComando = 0 Then gError 204144
        
        sNomeReduzido = String(STRING_FILIAL_NOME, 0)
        
        'le todas as filiais da empresa atual
        lErro = Comando_Executar(lComando, "SELECT NomeReduzido FROM Empresas WHERE Codigo = ? ", sNomeReduzido, glEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 204145
        
        'le a primeira filial
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204146
        
    Else
    
        sNomeReduzido = objFilialEmpresa.sNomeReduzido
        
    End If
    
    
    objSpedFiscal.sRegistro = "|0005|" & sNomeReduzido & "|" & objFilialEmpresa.objEndereco.sCEP & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Desacentua_SpedFiscal(objFilialEmpresa.objEndereco.sEndereco) & "|||" & Desacentua_SpedFiscal(objFilialEmpresa.objEndereco.sBairro) & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(sTelefone) & "|" & SFF_Texto(sFax) & "|" & SFF_Texto(objFilialEmpresa.objEndereco.sEmail) & "|"
    
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    objSpedRegistro.sTipo = objSpedFiscal.sTipo
    objSpedRegistro.iCount = objSpedRegistro.iCount + 1
    
    colSpedRegistro.Add objSpedRegistro
    
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203861
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203862

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_0005 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_0005:

   SpedFiscal_0005 = gErr

   Select Case gErr
    
        Case 203860, 204144
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203861
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203862
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case 204145, 104146
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203863)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_0015(ByVal objFilialEmpresa As AdmFiliais, objSpedFiscal As ClassSpedFiscal, ByVal colSpedRegistro As Collection) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim sInscricaoEstadual As String
Dim sUF As String
Dim sIE As String

On Error GoTo Erro_SpedFiscal_0015
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203864
    Next

    '******** registro 0015 *****************
    
    sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    sUF = String(STRING_ESTADO, 0)
    
    lErro = Comando_Executar(alComando(3), "SELECT UF, InscricaoEstadual FROM ContribuinteSubstituto WHERE FilialEmpresa = ? ", _
    sUF, sInscricaoEstadual, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 203865

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203866
    
    Set objSpedRegistro = New ClassSpedRegistro
    
    Do While lErro = AD_SQL_SUCESSO
    
        objSpedFiscal.sTipo = "0015"
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    
        Call Formata_String_Numero2(sInscricaoEstadual, sIE)
    
        objSpedFiscal.sRegistro = "|0015|" & sUF & "|" & sUF & "|" & sIE & "|"
    
        objSpedRegistro.iCount = objSpedRegistro.iCount + 1
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203867
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203868
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203869
    
    Loop
    
    If objSpedRegistro.iCount > 0 Then
        objSpedRegistro.sTipo = objSpedFiscal.sTipo
        colSpedRegistro.Add objSpedRegistro
    End If


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_0015 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_0015:

   SpedFiscal_0015 = gErr

   Select Case gErr
    
        Case 203864
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203865, 203866, 203869
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTRIBUINTESUBSTITUTO", gErr)
    
        Case 203867
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203868
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203870)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function
    
Function SpedFiscal_0100(ByVal objFilialEmpresa As AdmFiliais, objSpedFiscal As ClassSpedFiscal, ByVal colSpedRegistro As Collection) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim sTelefone As String
Dim sFax As String
Dim objCidades As New ClassCidades

On Error GoTo Erro_SpedFiscal_0100
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203871
    Next
    
    objSpedFiscal.sTipo = "0100"
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    
    
    Call Formata_String_Numero2(objFilialEmpresa.objEnderecoContador.sTelefone1, sTelefone)
    Call Formata_String_Numero2(objFilialEmpresa.objEnderecoContador.sFax, sFax)
    
    If Len(Trim(objFilialEmpresa.objEnderecoContador.sCidade)) <> 0 Then
    
        objCidades.sDescricao = objFilialEmpresa.objEnderecoContador.sCidade
        
        lErro = CF("Cidade_Le_Nome", objCidades)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203872
        
        If lErro = ERRO_OBJETO_NAO_CADASTRADO Then gError 203873
    
    End If
    
    objSpedFiscal.sRegistro = "|0100|" & SFF_Texto(objFilialEmpresa.sContador) & "|" & objFilialEmpresa.sCPFContador & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(objFilialEmpresa.sCRCContador) & "|" & objFilialEmpresa.sCNPJContador & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & objFilialEmpresa.objEnderecoContador.sCEP & "|" & SFF_Texto(objFilialEmpresa.objEnderecoContador.sEndereco) & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "||" & SFF_Texto(objFilialEmpresa.objEnderecoContador.sBairro) & "|" & SFF_Texto(sTelefone) & "|" & SFF_Texto(sFax) & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(objFilialEmpresa.objEnderecoContador.sEmail) & "|" & objCidades.sCodIBGE & "|"

    Set objSpedRegistro = New ClassSpedRegistro
    
    objSpedRegistro.sTipo = objSpedFiscal.sTipo
    objSpedRegistro.iCount = objSpedRegistro.iCount + 1
    
    colSpedRegistro.Add objSpedRegistro
    
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203874
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203875

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_0100 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_0100:

   SpedFiscal_0100 = gErr

   Select Case gErr
    
        Case 203871
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203872
    
        Case 203873
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_NAO_CADASTRADA2", gErr, objCidades.sDescricao)
        
        Case 203874
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203875
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203876)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_C101(ByVal objFisSpedAux As ClassFISSpedAux, ByVal lNumIntNF As Long, ByVal sUFRemet As String, ByVal sUFDest As String, ByVal iIndOper As Integer, ByVal iCodSit As Integer, ByVal lSeqPai As Long) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objC101 As ClassSpedC101
Dim objSpedFiscal As ClassSpedFiscal
Dim objTributacaoDoc As New ClassTributacaoDoc
Dim vUF As Variant, bAchou As Boolean

On Error GoTo Erro_SpedFiscal_C101

    objTributacaoDoc.iTipoDoc = 0
    objTributacaoDoc.lNumIntDoc = lNumIntNF
    
    lErro = CF("TributacaoDoc_Le", objTributacaoDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objTributacaoDoc.dICMSInterestVlrFCPUFDest + objTributacaoDoc.dICMSInterestVlrUFDest + objTributacaoDoc.dICMSInterestVlrUFRemet > DELTA_VALORMONETARIO Then
        
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 203871
        Next
        
        Set objC101 = New ClassSpedC101
        
        'Set objC101.objTrib = objTributacaoDoc
        objC101.dVlrFCPDest = objTributacaoDoc.dICMSInterestVlrFCPUFDest
        objC101.dVlrICMSDest = objTributacaoDoc.dICMSInterestVlrUFDest
        objC101.dVlrICMSRemet = objTributacaoDoc.dICMSInterestVlrUFRemet
        objC101.lNumIntNF = lNumIntNF
        objC101.sUFDest = sUFDest
        objC101.sUFRemet = sUFRemet
        objC101.iIndOper = iIndOper
        objC101.iCodSit = iCodSit
        
        Set objSpedFiscal = objFisSpedAux.objSpedFiscal
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objFisSpedAux.objSpedRegistroC101.iCount = objFisSpedAux.objSpedRegistroC101.iCount + 1
        
        objSpedFiscal.sTipo = "C101"
        objFisSpedAux.objSpedRegistroC101.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.sRegistro = "|C101|" '1 REG
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(objC101.dVlrFCPDest) & "|" '2 VL_FCP_UF_DEST
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(objC101.dVlrICMSDest) & "|" '3 VL_ICMS_UF_DEST
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(objC101.dVlrICMSRemet) & "|" '4 VL_ICMS_UF_REM
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 201224
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
        If lErro <> AD_SQL_SUCESSO Then gError 201225
    
        For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
        Next
        
        'Acumula para gerar os registros E300
        objFisSpedAux.colC101.Add objC101
        
        bAchou = False
        For Each vUF In objFisSpedAux.colC101UF
            If vUF = objC101.sUFDest Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then objFisSpedAux.colC101UF.Add objC101.sUFDest
        bAchou = False
        For Each vUF In objFisSpedAux.colC101UF
            If vUF = objC101.sUFRemet Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then objFisSpedAux.colC101UF.Add objC101.sUFRemet
        
    End If

    SpedFiscal_C101 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C101:

   SpedFiscal_C101 = gErr

   Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 201223
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201224
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201225
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204078)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_C110(ByVal objMensagem As ClassMensagem, objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC110 As ClassSpedRegistro, lSeqPai As Long, ByVal lSeqPaiNF As Long) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_SpedFiscal_C110
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203877
    Next

    If objMensagem.iCodigo <> 0 And gsPerfil <> "C" Then

        objSpedFiscal.sTipo = "C110"
        objSpedRegistroC110.sTipo = objSpedFiscal.sTipo
        
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC110.iCount = objSpedRegistroC110.iCount + 1
        lSeqPai = objSpedFiscal.lSeq
        
        objSpedFiscal.sRegistro = "|C110|" & objMensagem.iCodigo & "||"
        
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203878
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
        If lErro <> AD_SQL_SUCESSO Then gError 203879


    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C110 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C110:

   SpedFiscal_C110 = gErr

   Select Case gErr
    
        Case 203877
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203878
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203879
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203880)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_C111(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC111 As ClassSpedRegistro, ByVal lSeqPai As Long, ByRef tNFiscal As typeNFiscal) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim tSpedFisProcessoRef As TypeSpedFisProcessoRef

On Error GoTo Erro_SpedFiscal_C111
        
    If gsPerfil <> "C" Then
        
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 203881
        Next
    
        tSpedFisProcessoRef.sNumProc = String(STRING_SPEDFIS_NUMPROC, 0)
        
        lErro = Comando_Executar(alComando(2), "SELECT NumProc, IndProc FROM SpedFisProcessoRef WHERE NumIntNF = ?", _
        tSpedFisProcessoRef.sNumProc, tSpedFisProcessoRef.iIndProc, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203882
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203883
        
        Do While lErro = AD_SQL_SUCESSO
        
            objSpedFiscal.sTipo = "C111"
            objSpedRegistroC111.sTipo = objSpedFiscal.sTipo
            
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            objSpedRegistroC111.iCount = objSpedRegistroC111.iCount + 1
            
            objSpedFiscal.sRegistro = "|C111|" & tSpedFisProcessoRef.sNumProc & "|" & tSpedFisProcessoRef.iIndProc & "|"
            
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203884
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
            If lErro <> AD_SQL_SUCESSO Then gError 203885
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203886
        
        Loop
    
        For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
        Next

    End If

    SpedFiscal_C111 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C111:

   SpedFiscal_C111 = gErr

   Select Case gErr
    
        Case 203881
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203882, 203883, 203886
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISPROCESSOREF", gErr)
    
        Case 203884
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203885
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203887)
            
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_C112(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC112 As ClassSpedRegistro, ByVal lSeqPai As Long, ByRef tNFiscal As typeNFiscal) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim tSpedFisArrecadacaoRef As TypeSpedFisArrecadacaoRef

On Error GoTo Erro_SpedFiscal_C112
        
    If gsPerfil <> "C" Then
    
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 203888
        Next
    
        tSpedFisArrecadacaoRef.sCodigoDA = String(STRING_SPEDFIS_CODIGODA, 0)
        tSpedFisArrecadacaoRef.sNumDA = String(STRING_SPEDFIS_NUMDA, 0)
        tSpedFisArrecadacaoRef.sCodAut = String(STRING_SPEDFIS_CODAUT, 0)
        tSpedFisArrecadacaoRef.sUF = String(STRING_ESTADOS_SIGLA, 0)
        
        
        lErro = Comando_Executar(alComando(2), "SELECT CodigoDA, UF, NumDA, CodAut, ValorDA, DataVcto, DataPagto FROM SpedFisArrecadacaoRef WHERE NumIntNF = ?", _
        tSpedFisArrecadacaoRef.sCodigoDA, tSpedFisArrecadacaoRef.sUF, tSpedFisArrecadacaoRef.sNumDA, tSpedFisArrecadacaoRef.sCodAut, tSpedFisArrecadacaoRef.dValorDA, tSpedFisArrecadacaoRef.dtDataVcto, tSpedFisArrecadacaoRef.dtDataPagto, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203889
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203890
        
        Do While lErro = AD_SQL_SUCESSO
        
        
            objSpedFiscal.sTipo = "C112"
            objSpedRegistroC112.sTipo = objSpedFiscal.sTipo
            
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            objSpedRegistroC112.iCount = objSpedRegistroC112.iCount + 1
            
            objSpedFiscal.sRegistro = "|C112|" & tSpedFisArrecadacaoRef.sCodigoDA & "|" & tSpedFisArrecadacaoRef.sUF & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tSpedFisArrecadacaoRef.sNumDA & "|" & tSpedFisArrecadacaoRef.sCodAut & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tSpedFisArrecadacaoRef.dValorDA & "|" & tSpedFisArrecadacaoRef.dtDataVcto & "|" & tSpedFisArrecadacaoRef.dtDataPagto & "|"
            
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203891
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
            If lErro <> AD_SQL_SUCESSO Then gError 203892
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203893
        
        Loop
    
    
        For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
        Next

    End If

    SpedFiscal_C112 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C112:

   SpedFiscal_C112 = gErr

   Select Case gErr
    
        Case 203888
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203889, 203890, 203893
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISARRECADACAOREF", gErr)
    
        Case 203891
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203892
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203894)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_C113(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC113 As ClassSpedRegistro, ByVal lSeqPai As Long, ByRef tNFiscal As typeNFiscal, ByVal objSpedRegistro0150 As ClassSpedRegistro, ByVal objSpedRegistro0175 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim tSpedFisDocRef As TypeSpedFisDocRef
Dim sCodParticipante As String

On Error GoTo Erro_SpedFiscal_C113
        
    If gsPerfil <> "C" Then
        
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 203895
        Next
    
        tSpedFisDocRef.sIndOper = String(STRING_SPEDFIS_INDOPER, 0)
        tSpedFisDocRef.sIndEmit = String(STRING_SPEDFIS_INDEMIT, 0)
        tSpedFisDocRef.sCodMod = String(STRING_SPEDFIS_CODMOD, 0)
        tSpedFisDocRef.sSerie = String(STRING_SPEDFIS_SERIE, 0)
        
        lErro = Comando_Executar(alComando(2), "SELECT IndOper, IndEmit, Cliente, FilialCli, Fornecedor, FilialForn, CodMod, Serie, SubSerie, NumDoc, DataDoc FROM SpedFisDocRef WHERE NumIntNF = ?", _
        tSpedFisDocRef.sIndOper, tSpedFisDocRef.sIndEmit, tSpedFisDocRef.lCliente, tSpedFisDocRef.iFilialCli, tSpedFisDocRef.lFornecedor, tSpedFisDocRef.iFilialForn, tSpedFisDocRef.sCodMod, tSpedFisDocRef.sSerie, tSpedFisDocRef.iSubSerie, tSpedFisDocRef.lnumdoc, tSpedFisDocRef.dtDataDoc, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203896
    
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203897
        
        Do While lErro = AD_SQL_SUCESSO
        
            If tSpedFisDocRef.lCliente > 0 Then
            
                lErro = SpedFiscal_0150_Cliente(tSpedFisDocRef.lCliente, tSpedFisDocRef.iFilialCli, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
                If lErro <> SUCESSO Then gError 203898
    
    
            Else
        
        
                lErro = SpedFiscal_0150_Fornecedor(tSpedFisDocRef.lFornecedor, tSpedFisDocRef.iFilialForn, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
                If lErro <> SUCESSO Then gError 203899
        
            End If
            
        
            objSpedFiscal.sTipo = "C113"
            objSpedRegistroC113.sTipo = objSpedFiscal.sTipo
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            objSpedRegistroC113.iCount = objSpedRegistroC113.iCount + 1
            
            objSpedFiscal.sRegistro = "|C113|" & tSpedFisDocRef.sIndOper & "|" & tSpedFisDocRef.sIndEmit & "|" & sCodParticipante & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tSpedFisDocRef.sCodMod & "|" & tSpedFisDocRef.sSerie & "|" & tSpedFisDocRef.iSubSerie & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tSpedFisDocRef.lnumdoc & "|" & tSpedFisDocRef.dtDataDoc & "|"
            
            If giVersao >= 11 Then objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "|"
           
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203900
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
            If lErro <> AD_SQL_SUCESSO Then gError 203901
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203902
        
        Loop
    
        For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
        Next

    End If

    SpedFiscal_C113 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C113:

   SpedFiscal_C113 = gErr

   Select Case gErr
        
        Case 203895
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203896, 203897, 203902
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISDOCREF", gErr)
    
        Case 293898, 203899
    
        Case 203900
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203901
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203903)
            
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_C114(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC114 As ClassSpedRegistro, ByVal lSeqPai As Long, ByRef tNFiscal As typeNFiscal) As Long


Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim tSpedFisCupomRef As TypeSpedFisCupomRef

On Error GoTo Erro_SpedFiscal_C114
        
    If gsPerfil <> "C" Then
        
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 203904
        Next
    
        tSpedFisCupomRef.sCodMod = String(STRING_SPEDFIS_CODMOD, 0)
        tSpedFisCupomRef.sECFFab = String(STRING_SPEDFIS_ECFFAB, 0)
        
        lErro = Comando_Executar(alComando(2), "SELECT CodMod, ECFFab, ECFCx, NumDoc, DataDoc FROM SpedFisCupomRef WHERE NumIntNF = ?", _
        tSpedFisCupomRef.sCodMod, tSpedFisCupomRef.sECFFab, tSpedFisCupomRef.iECFCx, tSpedFisCupomRef.lnumdoc, tSpedFisCupomRef.dtDataDoc, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203905
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203906
        
        Do While lErro = AD_SQL_SUCESSO
        
            objSpedFiscal.sTipo = "C114"
            objSpedRegistroC114.sTipo = objSpedFiscal.sTipo
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            objSpedRegistroC114.iCount = objSpedRegistroC114.iCount + 1
            
            objSpedFiscal.sRegistro = "|C114|" & tSpedFisCupomRef.sCodMod & "|" & tSpedFisCupomRef.sECFFab & "|" & tSpedFisCupomRef.iECFCx & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tSpedFisCupomRef.lnumdoc & "|" & tSpedFisCupomRef.dtDataDoc & "|"
            
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203907
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
            If lErro <> AD_SQL_SUCESSO Then gError 203908
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203909
        
        Loop
    
        For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
        Next

    End If

    SpedFiscal_C114 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C114:

   SpedFiscal_C114 = gErr

   Select Case gErr
    
        Case 203904
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203905, 203906, 203909
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISCUPOMREF", gErr)
    
        Case 203907
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203908
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203910)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_C115(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC115 As ClassSpedRegistro, ByVal lSeqPai As Long, ByRef tNFiscal As typeNFiscal, ByRef tTipoDocInfo As typeTipoDocInfo, ByVal objFilialEmpresa As AdmFiliais) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 6) As Long
Dim iIndice As Integer
Dim iViaTransporte As Integer
Dim iIndCarga As Integer
Dim objCidades1 As New ClassCidades
Dim tFilialForn As typeFilialFornecedor
Dim tEndereco As typeEndereco
Dim objCidades As New ClassCidades
Dim tFilialCli As typeFilialCliente
Dim tEnderecoEntrega As typeEndereco
Dim sIE As String

On Error GoTo Erro_SpedFiscal_C115
       
    If gsPerfil <> "C" Then
       
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 203911
        Next
    
    
        'se for nota fiscal (codigo 01) ou nota do produto (codigo 04)
        If tTipoDocInfo.iModeloArqICMS = 1 Or tTipoDocInfo.iModeloArqICMS = 4 Then
            
            'se for uma nota de saida
            If tTipoDocInfo.iTipo = 2 Then
        
                '***************** registro C115 ************************
                
                If tNFiscal.iCodTransportadora <> 0 Then
                
                    'pegar os dados do fornecedor
                    lErro = Comando_Executar(alComando(2), "SELECT ViaTransporte FROM Transportadoras WHERE Codigo = ?", _
                    iViaTransporte, tNFiscal.iCodTransportadora)
                    If lErro <> AD_SQL_SUCESSO Then gError 203912
                
                    lErro = Comando_BuscarPrimeiro(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203913
                    
                    Select Case iViaTransporte
                    
                        Case 1 'Aereo
                            iIndCarga = 5
                            
                        Case 2, 5 'Maritimo, Fluvial
                            iIndCarga = 3
                            
                        Case 3 'Rodoviario
                            iIndCarga = 0
                            
                        Case 4 'Ferroviario
                            iIndCarga = 1
                            
                        Case 7 'Rodoviario/Ferroviario
                            iIndCarga = 2
                            
                        Case Else
                            iIndCarga = 9
                            
                    End Select
                    
                Else
                
                    iIndCarga = 9
                    
                End If
            
                objCidades1.sDescricao = objFilialEmpresa.objEndereco.sCidade
                
                lErro = CF("Cidade_Le_Nome", objCidades1)
                If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203914
                
            
                'se for um fornecedor
                If tNFiscal.lFornecedor <> 0 Then
        
                    'se tiver uma filial de entrega especificada e for diferente da filial do fornecedor
                    If tNFiscal.iFilialEntrega <> 0 And tNFiscal.iFilialEntrega <> tNFiscal.iFilialForn Then
        
                        tFilialForn.sCgc = String(STRING_CLIENTE_CGC, 0)
                        tFilialForn.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
                        tEndereco.sEndereco = String(STRING_ENDERECO, 0)
                        tEndereco.sBairro = String(STRING_BAIRRO, 0)
                        tEndereco.sCidade = String(STRING_CIDADE, 0)
            
                        'pegar os dados do fornecedor
                        lErro = Comando_Executar(alComando(3), "SELECT CGC, InscricaoEstatual, E.Endereco, E.Bairro, E.Cidade FROM FiliaisFornecedores AS FF, Enderecos AS E WHERE E.Codigo = FF.Endereco AND CodFornecedor = ? AND CodFilial = ?", _
                        tFilialForn.sCgc, tFilialForn.sInscricaoEstadual, tEndereco.sEndereco, tEndereco.sBairro, tEndereco.sCidade, tNFiscal.lFornecedor, tNFiscal.iFilialEntrega)
                        If lErro <> AD_SQL_SUCESSO Then gError 203915
                    
                        lErro = Comando_BuscarPrimeiro(alComando(3))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203916
        
                        objCidades.sDescricao = tEndereco.sCidade
                        
                        lErro = CF("Cidade_Le_Nome", objCidades)
                        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203917
                
                        Call Formata_String_Numero2(tFilialForn.sInscricaoEstadual, sIE)
                
                
                        objSpedFiscal.sRegistro = "|C115|" & iIndCarga & "||||" & objCidades1.sCodIBGE & "|" & IIf(Len(tFilialForn.sCgc) = 14, tFilialForn.sCgc, "") & "|"
                        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & sIE & "|" & IIf(Len(tFilialForn.sCgc) = 11, tFilialForn.sCgc, "") & "|" & objCidades.sCodIBGE & "|"
        
                    Else
                    
                        tEndereco.sCidade = String(STRING_CIDADE, 0)
            
                        'pegar os dados do fornecedor
                        lErro = Comando_Executar(alComando(4), "SELECT Cidade FROM FiliaisFornecedores AS FF, Enderecos AS E WHERE E.Codigo = FF.Endereco AND CodFornecedor = ? AND CodFilial = ?", _
                        tEndereco.sCidade, tNFiscal.lFornecedor, tNFiscal.iFilialForn)
                        If lErro <> AD_SQL_SUCESSO Then gError 203918
                    
                        lErro = Comando_BuscarPrimeiro(alComando(4))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203919
                    
                        objCidades.sDescricao = tEndereco.sCidade
                        
                        lErro = CF("Cidade_Le_Nome", objCidades)
                        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203920
                
                        objSpedFiscal.sRegistro = "|C115|" & iIndCarga & "||||" & objCidades1.sCodIBGE & "||||"
                        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & objCidades.sCodIBGE & "|"
                            
                    End If
                    
            
                Else
            
                    'se tiver uma filial de entrega especificada e for diferente da filial do Cliente
                    If tNFiscal.iFilialEntrega <> 0 And tNFiscal.iFilialEntrega <> tNFiscal.iFilialCli Then
            
                        tFilialCli.sCgc = String(STRING_CLIENTE_CGC, 0)
                        tFilialCli.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
                        tEnderecoEntrega.sEndereco = String(STRING_ENDERECO, 0)
                        tEnderecoEntrega.sBairro = String(STRING_BAIRRO, 0)
                        tEnderecoEntrega.sCidade = String(STRING_CIDADE, 0)
                        tEndereco.sEndereco = String(STRING_ENDERECO, 0)
                        tEndereco.sBairro = String(STRING_BAIRRO, 0)
                        tEndereco.sCidade = String(STRING_CIDADE, 0)
                
                        'pegar os dados do cliente
                        lErro = Comando_Executar(alComando(5), "SELECT CGC, InscricaoEstadual, E.Endereco, E.Bairro, E.Cidade, EE.Endereco, EE.Bairro, EE.Cidade FROM Clientes AS C, FiliaisClientes AS F, Enderecos AS E, Enderecos AS EE WHERE C.Codigo = CodCliente AND E.Codigo = F.Endereco AND EE.Codigo = F.EnderecoEntrega AND CodCliente = ? AND CodFilial = ?", _
                        tFilialCli.sCgc, tFilialCli.sInscricaoEstadual, tEndereco.sEndereco, tEndereco.sBairro, tEndereco.sCidade, tEnderecoEntrega.sEndereco, tEnderecoEntrega.sBairro, tEnderecoEntrega.sCidade, tNFiscal.lCliente, tNFiscal.iFilialEntrega)
                        If lErro <> AD_SQL_SUCESSO Then gError 203921
                
                        lErro = Comando_BuscarPrimeiro(alComando(5))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203922
                
                
                        If Len(Trim(tEnderecoEntrega.sEndereco)) > 0 Then
                        
                            objCidades.sDescricao = tEnderecoEntrega.sCidade
                            
                            lErro = CF("Cidade_Le_Nome", objCidades)
                            If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203923
                    
                            Call Formata_String_Numero2(tFilialCli.sInscricaoEstadual, sIE)
                    
                    
                            objSpedFiscal.sRegistro = "|C115|" & iIndCarga & "||||" & objCidades1.sCodIBGE & "|" & IIf(Len(tFilialCli.sCgc) = 14, tFilialCli.sCgc, "") & "|"
                            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & sIE & "|" & IIf(Len(tFilialCli.sCgc) = 11, tFilialCli.sCgc, "") & "|" & objCidades.sCodIBGE & "|"
                        
                        Else
                    
                            objCidades.sDescricao = tEndereco.sCidade
                            
                            lErro = CF("Cidade_Le_Nome", objCidades)
                            If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203924
                    
                            Call Formata_String_Numero2(tFilialCli.sInscricaoEstadual, sIE)
                    
                    
                            objSpedFiscal.sRegistro = "|C115|" & iIndCarga & "||||" & objCidades1.sCodIBGE & "|" & IIf(Len(tFilialCli.sCgc) = 14, tFilialCli.sCgc, "") & "|"
                            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & sIE & "|" & IIf(Len(tFilialCli.sCgc) = 11, tFilialCli.sCgc, "") & "|" & objCidades.sCodIBGE & "|"
                    
                        End If
                    
                    Else
                    
                        tFilialCli.sCgc = String(STRING_CLIENTE_CGC, 0)
                        tFilialCli.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
                        tEnderecoEntrega.sEndereco = String(STRING_ENDERECO, 0)
                        tEnderecoEntrega.sBairro = String(STRING_BAIRRO, 0)
                        tEnderecoEntrega.sCidade = String(STRING_CIDADE, 0)
                        tEndereco.sCidade = String(STRING_CIDADE, 0)
                
                        'pegar os dados do cliente
                        lErro = Comando_Executar(alComando(6), "SELECT CGC, InscricaoEstadual, EE.Endereco, EE.Bairro, EE.Cidade, E.Cidade FROM Clientes AS C, FiliaisClientes AS F, Enderecos AS E, Enderecos AS EE WHERE C.Codigo = CodCliente AND E.Codigo = F.Endereco AND EE.Codigo = F.EnderecoEntrega AND CodCliente = ? AND CodFilial = ?", _
                        tFilialCli.sCgc, tFilialCli.sInscricaoEstadual, tEnderecoEntrega.sEndereco, tEnderecoEntrega.sBairro, tEnderecoEntrega.sCidade, tEndereco.sCidade, tNFiscal.lCliente, tNFiscal.iFilialCli)
                        If lErro <> AD_SQL_SUCESSO Then gError 203925
                
                        lErro = Comando_BuscarPrimeiro(alComando(6))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203926
                        
                        
                        If Len(Trim(tEnderecoEntrega.sEndereco)) > 0 Then
                        
                            objCidades.sDescricao = tEnderecoEntrega.sCidade
                            
                            lErro = CF("Cidade_Le_Nome", objCidades)
                            If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203927
                    
                            Call Formata_String_Numero2(tFilialCli.sInscricaoEstadual, sIE)
                    
                            objSpedFiscal.sRegistro = "|C115|" & iIndCarga & "||||" & objCidades1.sCodIBGE & "|" & IIf(Len(tFilialCli.sCgc) = 14, tFilialCli.sCgc, "") & "|"
                            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & sIE & "|" & IIf(Len(tFilialCli.sCgc) = 11, tFilialCli.sCgc, "") & "|" & objCidades.sCodIBGE & "|"
                        
                        Else
                        
                            objCidades.sDescricao = tEndereco.sCidade
                            
                            lErro = CF("Cidade_Le_Nome", objCidades)
                            If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 203928
                    
                            objSpedFiscal.sRegistro = "|C115|" & iIndCarga & "||||" & objCidades1.sCodIBGE & "||||"
                            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & objCidades.sCodIBGE & "|"
                        
                        End If
                        
                        
                    End If
                    
            
                End If
                
                objSpedFiscal.sTipo = "C115"
                objSpedRegistroC115.sTipo = objSpedFiscal.sTipo
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistroC115.iCount = objSpedRegistroC115.iCount + 1
                
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203929
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                If lErro <> AD_SQL_SUCESSO Then gError 203930
                
                
            End If
            
        End If

    End If


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C115 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C115:

   SpedFiscal_C115 = gErr

   Select Case gErr
    
        Case 203911
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203912, 203913
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORA2", gErr)
    
        Case 203914, 203917, 203920, 203923, 203924, 203927, 203928
        
        Case 203915, 203916, 203918, 203919
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES", gErr)
        
        Case 203921, 203922, 203925, 203926
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
    
        Case 203929
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203930
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203931)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next


End Function


Private Function SpedFiscal_C120(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC120 As ClassSpedRegistro, ByVal lSeqPaiNF As Long, ByRef tNFiscal As typeNFiscal, ByRef tTipoDocInfo As typeTipoDocInfo) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim tDIInfo As typeDIInfo

On Error GoTo Erro_SpedFiscal_C120
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203932
    Next

    'se for nota fiscal (codigo 01)
    If tTipoDocInfo.iModeloArqICMS = 1 And gsPerfil <> "C" Then
    
        'se for uma nota de entrada
        If tTipoDocInfo.iTipo = 1 Or tTipoDocInfo.iTipo = 3 Then

            '******************** registro C120 ************************************
    
            tDIInfo.sNumero = String(STRING_DI_NUMERO, 0)
        
            'pegar os dados de DIInfo
            lErro = Comando_Executar(alComando(2), "SELECT DI.Numero, DI.PISValor, DI.COFINSValor FROM ItensNFiscal AS INF, ItemAdicaoDIItemNF AS IAD, ItensADicaoDI AS IDI, AdicaoDI AS ADI, DIINFO AS DI WHERE INF.NumIntNF = ? AND IAD.NumIntItemNF = INF.NumIntDoc AND IDI.NumIntDoc =IAD.NumIntItemAdicaoDI AND ADI.NumIntDoc = IDI.NumIntAdicaoDI AND ADI.NumIntDI = DI.NumIntDoc", _
            tDIInfo.sNumero, tDIInfo.dPISValor, tDIInfo.dCOFINSValor, tNFiscal.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 203933
    
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203934
        
            If lErro = AD_SQL_SUCESSO Then
        
                objSpedFiscal.sTipo = "C120"
                objSpedRegistroC120.sTipo = objSpedFiscal.sTipo
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistroC120.iCount = objSpedRegistroC120.iCount + 1
                
                
                objSpedFiscal.sRegistro = "|C120|0|" & SFF_Texto(tDIInfo.sNumero, 10) & "|" & Format(tDIInfo.dPISValor, "Fixed") & "|" & Format(tDIInfo.dCOFINSValor, "Fixed") & "||"
                
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203935
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
                If lErro <> AD_SQL_SUCESSO Then gError 203936
        
        
            End If
    
        End If
        
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C120 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C120:

   SpedFiscal_C120 = gErr

   Select Case gErr
    
        Case 203932
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203933, 203934
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)
    
        Case 203935
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203936
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203937)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Private Function SpedFiscal_C130(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC130 As ClassSpedRegistro, ByVal lSeqPaiNF As Long, ByRef tNFiscal As typeNFiscal, ByRef tTipoDocInfo As typeTipoDocInfo) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim dServNTribICMS As Double
Dim tTributacaoNF As typeTributacaoNF

On Error GoTo Erro_SpedFiscal_C130
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203938
    Next


    'se for uma nota de saida
    If tTipoDocInfo.iTipo = 2 And gsPerfil <> "C" Then
    
        '******************** registro C130 ************************************
    
        lErro = Comando_Executar(alComando(2), "SELECT SUM(INF.PrecoUnitario * INF.Quantidade) FROM ItensNFiscal AS INF, Produtos AS P, TributacaoItemNF AS T WHERE INF.NumIntNF = ?  AND INF.Produto = P.Codigo AND INF.NumIntNF = T.NumIntNF AND INF.Item = T.Item AND P.Natureza = 8 AND T.ICMSTipo <> 1", _
        dServNTribICMS, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203939

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203940
        
        lErro = Comando_Executar(alComando(3), "SELECT ISSBase, ISSValor, IRRFBase, IRRFValor, INSSValorBase, INSSValorDeducoes FROM TributacaoNF WHERE NumIntDoc = ?", _
        tTributacaoNF.dISSBase, tTributacaoNF.dISSValor, tTributacaoNF.dIRRFBase, tTributacaoNF.dIRRFValor, tTributacaoNF.dINSSBase, tTributacaoNF.dINSSDeducoes, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203941

        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203942
        
        
        objSpedFiscal.sTipo = "C130"
        objSpedRegistroC130.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC130.iCount = objSpedRegistroC130.iCount + 1
        
        
        objSpedFiscal.sRegistro = "|C130|" & IIf(dServNTribICMS = 0, 0, Format(dServNTribICMS, "Fixed")) & "|" & IIf(tTributacaoNF.dISSBase = 0, 0, Format(tTributacaoNF.dISSBase, "Fixed")) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dISSValor = 0, 0, Format(tTributacaoNF.dISSValor, "Fixed")) & "|" & IIf(tTributacaoNF.dIRRFBase = 0, 0, Format(tTributacaoNF.dIRRFBase, "Fixed")) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dIRRFValor = 0, 0, Format(tTributacaoNF.dIRRFValor, "Fixed")) & "|" & IIf(tTributacaoNF.dINSSBase = 0, 0, Format(tTributacaoNF.dINSSBase, "Fixed")) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dINSSDeducoes = 0, 0, Format(tTributacaoNF.dINSSDeducoes, "Fixed")) & "|"
        
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203943
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
        If lErro <> AD_SQL_SUCESSO Then gError 203944
    
    
    End If


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    

    SpedFiscal_C130 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C130:

   SpedFiscal_C130 = gErr

   Select Case gErr
    
        Case 203938
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203939, 203940
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)
    
        Case 203941, 203942
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAONF", gErr)
    
        Case 203943
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203944
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203945)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Private Function SpedFiscal_C140(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC140 As ClassSpedRegistro, ByVal objSpedRegistroC141 As ClassSpedRegistro, ByVal lSeqPaiNF As Long, ByRef tNFiscal As typeNFiscal, ByRef tTipoDocInfo As typeTipoDocInfo, ByVal iIndPag As Integer) As Long


Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 5) As Long
Dim iIndice As Integer
Dim lSeqPai As Long
Dim tTitPag As typeTituloPagar
Dim tParcPag As typeParcelaPagar
Dim tTitRec As typeTituloReceber
Dim tParcRec As typeParcelaReceber

On Error GoTo Erro_SpedFiscal_C140
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203946
    Next


    'se for nota fiscal (codigo 01) e for condicao de pagto a prazo
    If tTipoDocInfo.iModeloArqICMS = 1 And iIndPag = 1 Then
    
        '******************** registro C140 ************************************
    
        'se fir um titulo a pagar
        If tNFiscal.iClasseDocCPR = 1 Then

            lErro = Comando_Executar(alComando(2), "SELECT NumTitulo, NumParcelas, ValorTotal FROM TitulosPagTodos  WHERE NumIntDoc  = ?", _
            tTitPag.lNumTitulo, tTitPag.iNumParcelas, tTitPag.dValorTotal, tNFiscal.lNumIntDocCPR)
            If lErro <> AD_SQL_SUCESSO Then gError 203947
    
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203948
            
            objSpedFiscal.sTipo = "C140"
            objSpedRegistroC140.sTipo = objSpedFiscal.sTipo
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            lSeqPai = objSpedFiscal.lSeq
            objSpedRegistroC140.iCount = objSpedRegistroC140.iCount + 1
            
            objSpedFiscal.sRegistro = "|C140|1|00||" & tTitPag.lNumTitulo & "|" & tTitPag.iNumParcelas & "|" & IIf(tTitPag.dValorTotal = 0, 0, Format(tTitPag.dValorTotal, "Fixed")) & "|"
            
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203949
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
            If lErro <> AD_SQL_SUCESSO Then gError 203950

            lErro = Comando_Executar(alComando(3), "SELECT NumParcela, DataVencimento, Valor FROM ParcelasPagTodas  WHERE NumIntTitulo  = ? ORDER BY NumParcela", _
            tParcPag.iNumParcela, tParcPag.dtDataVencimento, tParcPag.dValor, tNFiscal.lNumIntDocCPR)
            If lErro <> AD_SQL_SUCESSO Then gError 203951
    
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203952

            Do While lErro = AD_SQL_SUCESSO

                '******************** registro C141 ************************************

                objSpedFiscal.sTipo = "C141"
                objSpedRegistroC141.sTipo = objSpedFiscal.sTipo
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistroC141.iCount = objSpedRegistroC141.iCount + 1
                
                
                objSpedFiscal.sRegistro = "|C141|" & tParcPag.iNumParcela & "|" & Format(tParcPag.dtDataVencimento, "ddmmyyyy") & "|" & IIf(tParcPag.dValor = 0, 0, Format(tParcPag.dValor, "Fixed")) & "|"
                
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203953
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                If lErro <> AD_SQL_SUCESSO Then gError 203954

                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203955

            Loop
            
        ElseIf tNFiscal.iClasseDocCPR = 2 Then
            
            '******************** registro C140 ************************************
            
            lErro = Comando_Executar(alComando(4), "SELECT NumTitulo, NumParcelas, Valor FROM TitulosRecTodos  WHERE NumIntDoc  = ?", _
            tTitRec.lNumTitulo, tTitRec.iNumParcelas, tTitRec.dValor, tNFiscal.lNumIntDocCPR)
            If lErro <> AD_SQL_SUCESSO Then gError 203956
    
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203957
            
            objSpedFiscal.sTipo = "C140"
            objSpedRegistroC140.sTipo = objSpedFiscal.sTipo
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            lSeqPai = objSpedFiscal.lSeq
            objSpedRegistroC140.iCount = objSpedRegistroC140.iCount + 1
            
            
            objSpedFiscal.sRegistro = "|C140|1|00||" & tTitRec.lNumTitulo & "|" & tTitRec.iNumParcelas & "|" & IIf(tTitRec.dValor = 0, 0, Format(tTitRec.dValor, "Fixed")) & "|"
            
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203958
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
            If lErro <> AD_SQL_SUCESSO Then gError 203959

            lErro = Comando_Executar(alComando(5), "SELECT NumParcela, DataVencimento, Valor FROM ParcelasRecTodas  WHERE NumIntTitulo  = ? ORDER BY NumParcela", _
            tParcRec.iNumParcela, tParcRec.dtDataVencimento, tParcRec.dValor, tNFiscal.lNumIntDocCPR)
            If lErro <> AD_SQL_SUCESSO Then gError 203960
    
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203961

            Do While lErro = AD_SQL_SUCESSO

                '******************** registro C141 ************************************

                objSpedFiscal.sTipo = "C141"
                objSpedRegistroC141.sTipo = objSpedFiscal.sTipo
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistroC141.iCount = objSpedRegistroC141.iCount + 1
                
                
                objSpedFiscal.sRegistro = "|C141|" & tParcRec.iNumParcela & "|" & Format(tParcRec.dtDataVencimento, "ddmmyyyy") & "|" & IIf(tParcRec.dValor = 0, 0, Format(tParcRec.dValor, "Fixed")) & "|"
                
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203962
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                If lErro <> AD_SQL_SUCESSO Then gError 203963

                lErro = Comando_BuscarProximo(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203964

            Loop

        End If
    
    End If


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C140 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C140:

   SpedFiscal_C140 = gErr

   Select Case gErr
    
        Case 203946
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203947, 203948
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAGTODOS", gErr)
    
        Case 203949, 203953, 203958, 203962
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203950, 203954, 203959, 203963
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case 203951, 203952, 203955
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPAGTODAS", gErr)
    
        Case 203956, 203957
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTODOS", gErr)
        
        Case 203960, 203961, 203964
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASRECTODAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203965)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_C160(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC160 As ClassSpedRegistro, ByVal lSeqPaiNF As Long, ByRef tNFiscal As typeNFiscal, ByRef tTipoDocInfo As typeTipoDocInfo, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal objSpedRegistro0150 As ClassSpedRegistro, ByVal objSpedRegistro0175 As ClassSpedRegistro) As Long


Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 5) As Long
Dim iIndice As Integer
Dim sCodParticipante As String

On Error GoTo Erro_SpedFiscal_C160
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203966
    Next


    'se for nota fiscal (codigo 01) ou nota do produto (codigo 04)
    If tTipoDocInfo.iModeloArqICMS = 1 Or tTipoDocInfo.iModeloArqICMS = 4 Then
    
        '******************** registro C160 ************************************
    
        'se for uma nota de saida
        If tTipoDocInfo.iTipo = 2 Then
    
            lErro = SpedFiscal_0150_Transportadora(tNFiscal.iCodTransportadora, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
            If lErro <> SUCESSO Then gError 203967
                
            objSpedFiscal.sTipo = "C160"
            objSpedRegistroC160.sTipo = objSpedFiscal.sTipo
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            objSpedRegistroC160.iCount = objSpedRegistroC160.iCount + 1
            
            
            objSpedFiscal.sRegistro = "|C160|" & sCodParticipante & "|" & tNFiscal.sPlaca & "|" & tNFiscal.lVolumeQuant & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tNFiscal.dPesoBruto = 0, 0, Format(tNFiscal.dPesoBruto, "Fixed")) & "|" & IIf(tNFiscal.dPesoLiq = 0, 0, Format(tNFiscal.dPesoLiq, "Fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tNFiscal.sPlacaUF & "|"
            
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 203968
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
            If lErro <> AD_SQL_SUCESSO Then gError 203969
    
        End If
    
    End If


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C160 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C160:

   SpedFiscal_C160 = gErr

   Select Case gErr
    
        Case 203966
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203967
    
        Case 203968
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203969
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203970)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_C170(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC170 As ClassSpedRegistro, ByVal objSpedRegistroC172 As ClassSpedRegistro, ByVal objSpedRegistro0200 As ClassSpedRegistro, ByVal objSpedRegistro0205 As ClassSpedRegistro, ByVal objSpedRegistro0190 As ClassSpedRegistro, ByVal objSpedRegistro0220 As ClassSpedRegistro, ByVal lSeqPaiNF As Long, ByRef tNFiscal As typeNFiscal, ByRef tTipoDocInfo As typeTipoDocInfo, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal iNFe As Integer, dTotalRateioNF As Double, ByVal colUM As Collection, ByVal colSpedE510 As Collection, dValDebICMS_objSpedC190 As Double, dValCredICMS_objSpedC190 As Double, colSpedC190 As Collection, ByVal objFilialEmpresa As AdmFiliais, ByRef tTributacaoNF As typeTributacaoNF, ByVal sData As String, colSpedC190UF As Collection, ByVal objFisSpedAux As ClassFISSpedAux) As Long

'Private Function SpedFiscal_C170(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC170 As ClassSpedRegistro, ByVal objSpedRegistroC172 As ClassSpedRegistro, ByVal objSpedRegistro0200 As ClassSpedRegistro, ByVal objSpedRegistro0205 As ClassSpedRegistro, ByVal objSpedRegistro0190 As ClassSpedRegistro, ByVal objSpedRegistro0220 As ClassSpedRegistro, ByVal lSeqPaiNF As Long, ByRef tNFiscal As typeNFiscal, ByRef tTipoDocInfo As typeTipoDocInfo, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal iNFe As Integer, dTotalRateioNF As Double, ByVal colUM As Collection, ByVal colSpedE510 As Collection, dValDebICMS_objSpedC190 As Double, dValCredICMS_objSpedC190 As Double, colSpedC190 As Collection, ByVal objFilialEmpresa As AdmFiliais, ByRef tTributacaoNF As typeTributacaoNF, ByVal sData As String, _
'dTotalRateioNFICMSBase As Double, dTotalRateioNFICMSValor As Double, dTotalRateioNFICMSSubstBase As Double, dTotalRateioNFICMSSubstValor As Double, dTotalRateioNFIPIBaseCalculo As Double, dTotalRateioNFIPIValor As Double) As Long


Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 5) As Long
Dim iIndice As Integer
Dim tItemNF As typeItemNF
Dim tTribItemNF As typeTributacaoDocItem
Dim iTipoTribCST_ICMS As Integer
Dim iIPICSTEntrada As Integer
Dim iIPICSTSaida As Integer
Dim iPISCST As Integer
Dim dPISAliquota As Double
Dim iCOFINSCST As Integer
Dim dCOFINSAliquota As Double
Dim tProduto As typeProduto
Dim objSpedE510Aux As ClassSpedE510Aux
Dim colSpedE510Aux As Collection
Dim dValorItemNFTotal As Double
Dim sUMResult As String
Dim lSeqPai As Long
Dim iISSIncide As Integer
Dim objSpedC190 As ClassSpedC190
Dim objSpedE510 As ClassSpedE510
Dim lSeq0200 As Long
Dim iAchou As Integer
Dim iTipoRegistro As Integer
Dim dTotalRateioNFICMSBase As Double
Dim objSpedC190UF As ClassSpedC190UF
Dim dIPIValor As Double, objTribItem As ClassTributacaoDocItem
Dim dValorTotalItem As Double, dICMSSubstValorExterno As Double
Dim iTipoTribCST_ICMSAux As Integer
Dim dValorFCP As Double
Dim dValorFCPST As Double
Dim dValorFCPRet As Double

On Error GoTo Erro_SpedFiscal_C170
        
    Set colSpedC190 = New Collection
        
    If Len(Trim(sData)) > 0 Then sData = left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 4)
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203971
    Next

    '****************************** registro C170 *****************************************
    'se for nota fiscal (codigo 01) ou nota do produto (codigo 04)
    If tTipoDocInfo.iModeloArqICMS = 1 Or tTipoDocInfo.iModeloArqICMS = 4 Or iNFe = 1 Then
    
        dTotalRateioNF = 0
    
        tItemNF.sProduto = String(STRING_PRODUTO, 0)
        tItemNF.sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
        tItemNF.sUnidadeMed = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
    
        lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Item, Produto, DescricaoItem, Quantidade, UnidadeMed, ClasseUM, PrecoUnitario, ValorDesconto FROM ItensNFiscal, Produtos  WHERE Produto = Codigo AND NumIntNF  = ?", _
        tItemNF.lNumIntDoc, tItemNF.iItem, tItemNF.sProduto, tItemNF.sDescricaoItem, tItemNF.dQuantidade, tItemNF.sUnidadeMed, tProduto.iClasseUM, tItemNF.dPrecoUnitario, tItemNF.dValorDesconto, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203972
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203973
    
        Set colSpedE510Aux = New Collection
    
        Do While lErro = AD_SQL_SUCESSO
    
            tTribItemNF.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
            tTribItemNF.sCST = String(STRING_CST, 0)
     
            lErro = Comando_Executar(alComando(3), "SELECT CST, OrigemMercadoria, COFINSCredito, TipoTribCST, NaturezaOP, ICMSBase, ICMSAliquota, ICMSValor, ICMSCredito, ICMSSubstBase, ICMSSubstAliquota, ICMSSubstValor, IPI.CSTEntrada, IPI.CSTSaida, " & _
                "IPIBaseCalculo, IPIAliquota, IPIValor, IPICredito, PISCredito, TipoTributacao, ICMSPercRedBase, DespImpICMSBase, DespImpICMSValor, PISTipo, PISTipoCalculo, PISBase, PISAliquota, PISAliquotaValor, PISQtde, PISValor, COFINSTipo, COFINSTipoCalculo, COFINSBase, COFINSAliquota, COFINSAliquotaValor, COFINSQtde, COFINSValor, ValorFreteItem, ValorSeguroItem, ValorOutrasDespesasItem, DescontoGrid, ICMSvFCP, ICMSvFCPST, ICMSvFCPSTRet FROM TributacaoDocItem, tiposTribICMS AS ICMS, TiposTribIPI AS IPI  WHERE NumIntDoc = ? AND Item = ? AND ICMS.Tipo = ICMSTipo AND IPI.Tipo = IPITipo", _
                tTribItemNF.sCST, tTribItemNF.iOrigemMercadoria, tTribItemNF.dCOFINSCredito, iTipoTribCST_ICMSAux, tTribItemNF.sNaturezaOp, tTribItemNF.dICMSBase, tTribItemNF.dICMSAliquota, tTribItemNF.dICMSValor, tTribItemNF.dICMSCredito, tTribItemNF.dICMSSubstBase, tTribItemNF.dICMSSubstAliquota, tTribItemNF.dICMSSubstValor, _
                iIPICSTEntrada, iIPICSTSaida, tTribItemNF.dIPIBaseCalculo, tTribItemNF.dIPIAliquota, tTribItemNF.dIPIValor, tTribItemNF.dIPICredito, tTribItemNF.dPISCredito, tTribItemNF.iTipoTributacao, tTribItemNF.dICMSPercRedBase, tTribItemNF.dDespImpICMSBase, tTribItemNF.dDespImpICMSValor, _
                tTribItemNF.iPISTipo, tTribItemNF.iPISTipoCalculo, tTribItemNF.dPISBase, tTribItemNF.dPISAliquota, tTribItemNF.dPISAliquotaValor, tTribItemNF.dPISQtde, tTribItemNF.dPISValor, tTribItemNF.iCOFINSTipo, tTribItemNF.iCOFINSTipoCalculo, tTribItemNF.dCOFINSBase, tTribItemNF.dCOFINSAliquota, tTribItemNF.dCOFINSAliquotaValor, tTribItemNF.dCOFINSQtde, tTribItemNF.dCOFINSValor, tTribItemNF.dValorFreteItem, tTribItemNF.dValorSeguroItem, tTribItemNF.dValorOutrasDespesasItem, tTribItemNF.dDescontoGrid, dValorFCP, dValorFCPST, dValorFCPRet, _
                tNFiscal.lNumIntDoc, tItemNF.iItem)
            If lErro <> AD_SQL_SUCESSO Then gError 203974
        
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203975
    
'            If iTipoTribCST_ICMS = 0 And tTribItemNF.dICMSValor <> 0 And tTribItemNF.dICMSCredito = 0 Then
'                iTipoTribCST_ICMS = 90
'            End If
'
'            If iTipoTribCST_ICMS = 90 Then
'                tTribItemNF.dICMSBase = 0
'                tTribItemNF.dICMSAliquota = 0
'                tTribItemNF.dICMSValor = 0
'                tTribItemNF.dICMSCredito = 0
'            End If
'
'            'Não vejo sentido em zerar o IPI quando não tem crédito
'            'O IPI afeta o valor das NFs ... vou comentar esse código (Wagner)
'            dIPIValor = tTribItemNF.dIPIValor
'            If iIPICSTEntrada = 0 And tTribItemNF.dIPIValor <> 0 And tTribItemNF.dIPICredito = 0 Then
'                iIPICSTEntrada = 49
'            End If
'
'            If iIPICSTEntrada = 49 Then
'                tTribItemNF.dIPIBaseCalculo = 0
'                tTribItemNF.dIPIAliquota = 0
'                tTribItemNF.dIPIValor = 0
'                tTribItemNF.dIPICredito = 0
'            End If

            iTipoTribCST_ICMS = iTipoTribCST_ICMSAux

            dICMSSubstValorExterno = tTribItemNF.dICMSSubstValor

            Set objTribItem = New ClassTributacaoDocItem
            objTribItem.lNumIntDocItem = tItemNF.lNumIntDoc
            objTribItem.iTipoDoc = TIPODOC_TRIB_NF
            lErro = ItemNF_Le_BasePropria(objTribItem, iTipoTribCST_ICMS, iIPICSTEntrada)
            If lErro <> SUCESSO Then gError 203976
            
            iTipoTribCST_ICMS = iTipoTribCST_ICMS + (tTribItemNF.iOrigemMercadoria * 100)
            
            tTribItemNF.dICMSBase = objTribItem.dICMSBase
            tTribItemNF.dICMSAliquota = objTribItem.dICMSAliquota
            tTribItemNF.dICMSValor = objTribItem.dICMSValor
            tTribItemNF.dICMSCredito = objTribItem.dICMSCredito
            
            tTribItemNF.dICMSSubstBase = objTribItem.dICMSSubstBase
            tTribItemNF.dICMSSubstAliquota = objTribItem.dICMSSubstAliquota
            tTribItemNF.dICMSSubstValor = objTribItem.dICMSSubstValor
            tTribItemNF.dICMSSubstPercMVA = objTribItem.dICMSSubstPercMVA
            tTribItemNF.dICMSSubstPercRedBase = objTribItem.dICMSSubstPercRedBase
        
            dIPIValor = tTribItemNF.dIPIValor

            tTribItemNF.dIPIBaseCalculo = objTribItem.dIPIBaseCalculo
            tTribItemNF.dIPIAliquota = objTribItem.dIPIAliquota
            tTribItemNF.dIPIValor = objTribItem.dIPIValor
            tTribItemNF.dIPICredito = objTribItem.dIPICredito

'            lErro = PIS_CST(iPISCST, tTribItemNF)
'            If lErro <> SUCESSO Then gError 203976
'
'            If iPISCST = 1 Then
'
'                lErro = PIS_Aliquota(dPISAliquota, objFilialEmpresa)
'                If lErro <> SUCESSO Then gError 203977
'
'            Else
'                dPISAliquota = 0
'            End If
'
'            lErro = COFINS_CST(iCOFINSCST, tTribItemNF)
'            If lErro <> SUCESSO Then gError 203978
'
'            If iCOFINSCST = 1 Then
'
'                lErro = COFINS_Aliquota(dCOFINSAliquota, objFilialEmpresa)
'                If lErro <> SUCESSO Then gError 203979
'
'            Else
'                dCOFINSAliquota = 0
'            End If
            
            'nao foi rateado pois senao os valores dos itens nao bateriam com os totalizadores C190
'            lErro = Rateia_Compl_ItemNF(tNFiscal, tItemNF, tTribItemNF, dTotalRateioNFICMSBase, dTotalRateioNFICMSValor, dTotalRateioNFICMSSubstBase, dTotalRateioNFICMSSubstValor, dTotalRateioNFIPIBaseCalculo, dTotalRateioNFIPIValor)
        
'            lErro = Rateia_Compl_ItemNF(tNFiscal, tItemNF, tTribItemNF)
'            If lErro <> SUCESSO Then gError 204206
        
        
            If iNFe <> 1 Or (tTipoDocInfo.iTipo = 3) Then
        
                lErro = SpedFiscal_0200(tItemNF.sProduto, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, lSeq0200, objFisSpedAux)
                If lErro <> SUCESSO Then gError 203980
        
                lErro = SpedFiscal_0190(tItemNF.sUnidadeMed, tProduto.iClasseUM, objSpedFiscal, objSpedRegistro0190, objSpedRegistro0220, colUM, sUMResult, tItemNF.sProduto, lSeq0200)
                If lErro <> SUCESSO Then gError 203981
        
        
                objSpedFiscal.sTipo = "C170"
                objSpedRegistroC170.sTipo = objSpedFiscal.sTipo
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                lSeqPai = objSpedFiscal.lSeq
                objSpedRegistroC170.iCount = objSpedRegistroC170.iCount + 1
                
                
                objSpedFiscal.sRegistro = "|C170|" & tItemNF.iItem & "|" & SFF_Texto(tItemNF.sProduto) & "|" & SFF_Texto(tItemNF.sDescricaoItem) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tItemNF.dQuantidade, "##############0.0##") & "|" & sUMResult & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tItemNF.dQuantidade * tItemNF.dPrecoUnitario, "Fixed") & "|" & IIf(tItemNF.dValorDesconto = 0, "0", Format(tItemNF.dValorDesconto, "Fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTipoDocInfo.iTipoMovtoEstoque = 0, 1, 0) & "|" & Format(iTipoTribCST_ICMS, "000") & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tTribItemNF.sNaturezaOp & "||" & IIf(tTribItemNF.dICMSBase = 0, 0, Format(tTribItemNF.dICMSBase - (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase), "Fixed")) & "|" & IIf(tTribItemNF.dICMSAliquota = 0, 0, Format(tTribItemNF.dICMSAliquota * 100, "Fixed")) & "|"
'mario                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dICMSValor = 0, 0, Format(tTribItemNF.dICMSValor, "Fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dICMSCredito = 0, 0, Format(tTribItemNF.dICMSCredito, "Fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dICMSSubstBase = 0, 0, Format(tTribItemNF.dICMSSubstBase, "Fixed")) & "|" & IIf(tTribItemNF.dICMSSubstAliquota = 0, 0, Format(tTribItemNF.dICMSSubstAliquota * 100, "Fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dICMSSubstValor = 0, 0, Format(tTribItemNF.dICMSSubstValor, "Fixed")) & "||"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTipoDocInfo.iTipo = 2, Format(iIPICSTSaida, "00"), Format(iIPICSTEntrada, "00")) & "||"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dIPIBaseCalculo = 0, 0, Format(tTribItemNF.dIPIBaseCalculo, "Fixed")) & "|" & IIf(tTribItemNF.dIPIAliquota = 0, 0, Format(tTribItemNF.dIPIAliquota * 100, "Fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dIPICredito = 0, 0, Format(tTribItemNF.dIPICredito, "Fixed")) & "|"
                
                If tTribItemNF.dPISCredito = 0 Then
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "04|0|0|0|0|0|"
                Else
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tTribItemNF.iPISTipo, "00") & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tTribItemNF.dPISBase, "Fixed") & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dPISAliquota = 0, 0, Format(tTribItemNF.dPISAliquota * 100, "Fixed")) & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dPISQtde = 0, 0, Format(tTribItemNF.dPISQtde, "##############0.0##")) & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dPISAliquotaValor = 0, 0, Format(tTribItemNF.dPISAliquotaValor, "Fixed")) & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tTribItemNF.dPISValor, "Fixed") & "|"
                End If
                
                If tTribItemNF.dCOFINSCredito = 0 Then
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "04|0|0|0|0|0|"
                Else
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tTribItemNF.iCOFINSTipo, "00") & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tTribItemNF.dCOFINSBase, "Fixed") & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dCOFINSAliquota = 0, 0, Format(tTribItemNF.dCOFINSAliquota * 100, "Fixed")) & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dCOFINSQtde = 0, 0, Format(tTribItemNF.dCOFINSQtde, "##############0.0##")) & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTribItemNF.dCOFINSAliquotaValor = 0, 0, Format(tTribItemNF.dCOFINSAliquotaValor, "Fixed")) & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tTribItemNF.dCOFINSValor, "Fixed") & "|"
                End If
                
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "|"
                
                If giVersao > 12 Then
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "0|" '38 VL_ABAT_NT Valor do abatimento não tributado e não comercial
                End If
                 
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203982
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
                If lErro <> AD_SQL_SUCESSO Then gError 203983
    
            End If
            
            Set objSpedE510Aux = New ClassSpedE510Aux
            
            objSpedE510Aux.sCFOP = tTribItemNF.sNaturezaOp
            objSpedE510Aux.iCSTIPI = IIf(tTipoDocInfo.iTipo = 2, Format(iIPICSTSaida, "00"), Format(iIPICSTEntrada, "00"))
            objSpedE510Aux.dValorItemNF = tItemNF.dQuantidade * tItemNF.dPrecoUnitario - tItemNF.dValorDesconto + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem + tTribItemNF.dICMSSubstValor
            objSpedE510Aux.dValorIPI = tTribItemNF.dIPICredito
            objSpedE510Aux.dValorBCIPI = tTribItemNF.dIPIBaseCalculo
            
            colSpedE510Aux.Add objSpedE510Aux
            
    
    
            '*************** registro C172 ***************************************
            '??? informar dados de prestcao de servicos: pegar direto de tributacaodocitem
            'se for nota fiscal (codigo 01) ou nota do produto (codigo 04)
            If tTipoDocInfo.iModeloArqICMS = 1 Then
            
                'se for uma nota de saida
                If tTipoDocInfo.iTipo = 2 And iNFe <> 1 Then
            
    
                    lErro = Comando_Executar(alComando(4), "SELECT ISSIncide FROM TiposDeTributacaoMovto WHERE  Tipo = ?", _
                    iISSIncide, tTribItemNF.iTipoTributacao)
                    If lErro <> AD_SQL_SUCESSO Then gError 203984
                
                    lErro = Comando_BuscarPrimeiro(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203985
    
                    If iISSIncide = 1 Then
    
                        objSpedFiscal.sTipo = "C172"
                        objSpedRegistroC172.sTipo = objSpedFiscal.sTipo
                        
                        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                        objSpedRegistroC172.iCount = objSpedRegistroC172.iCount + 1
    
                        objSpedFiscal.sRegistro = "|C172|" & Format(tItemNF.dPrecoUnitario * tItemNF.dQuantidade, "fixed") & "|"
                        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tTributacaoNF.dISSAliquota * 100, "##0.00") & "|"
                        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tItemNF.dPrecoUnitario * tItemNF.dQuantidade * tTributacaoNF.dISSAliquota, "fixed") & "|"
                        
                        
                        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203986
                        
                        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
                        If lErro <> AD_SQL_SUCESSO Then gError 203987
    
    
                    End If
                            
                End If
                
            End If
    
            iAchou = 0
    
            For Each objSpedC190 In colSpedC190
            
                If objSpedC190.iCSTICMS = iTipoTribCST_ICMS And objSpedC190.sCFOP = tTribItemNF.sNaturezaOp And objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota Then
                    
                    dValorTotalItem = Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + dICMSSubstValorExterno)
                    If tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC Then
                        If tTribItemNF.dICMSBase <> 0 And tTribItemNF.dICMSBase > dValorTotalItem And gobjCRFAT.iNFImportacaoTribFlag03 = DESMARCADO Then
                            dValorTotalItem = tTribItemNF.dICMSBase
                        Else
                            If gobjCRFAT.iNFImportacaoTribFlag04 = DESMARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dICMSValor
                            If gobjCRFAT.iNFImportacaoTribFlag08 = MARCADO Then dValorTotalItem = dValorTotalItem - tTribItemNF.dValorOutrasDespesasItem
                            If gobjCRFAT.iNFImportacaoTribFlag05 = MARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dIIValor
                            
                            If gobjCRFAT.iNFImportacaoTribFlag12 = MARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dPISValor + tTribItemNF.dCOFINSValor

                        End If
                    End If
                    
                    objSpedC190.dValorOper = objSpedC190.dValorOper + dValorTotalItem 'Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + tTribItemNF.dICMSSubstValor)
                    dTotalRateioNF = dTotalRateioNF + Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + dICMSSubstValorExterno)
                    
'                    objSpedC190.dValorOper = objSpedC190.dValorOper + (tItemNF.dQuantidade * tItemNF.dPrecoUnitario)
'                    objSpedC190.dValorOper = objSpedC190.dValorOper + tTribItemNF.dICMSSubstValor + tTribItemNF.dIPIValor
'                    If tNFiscal.dValorProdutos > 0 Then
'                        objSpedC190.dValorOper = objSpedC190.dValorOper + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'                        dTotalRateioNF = dTotalRateioNF + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
''                        objSpedC190.dValorOper = objSpedC190.dValorOper - Round(((tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'                    End If
                    
                    
                    objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS + tTribItemNF.dICMSBase
'                    If tNFiscal.iTipoNFiscal = 119 Then
'                        objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS + tTribItemNF.dDespImpICMSBase
'                    End If
                    
'mario                    objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dICMSValor
                    objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dICMSCredito
'                    If tNFiscal.iTipoNFiscal = 119 Then
'                        objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dDespImpICMSValor
'                    End If

                    
                    objSpedC190.dValorBCICMSST = objSpedC190.dValorBCICMSST + tTribItemNF.dICMSSubstBase
                    objSpedC190.dValorICMSST = objSpedC190.dValorICMSST + tTribItemNF.dICMSSubstValor
                    
                    objSpedC190.dValorRedBC = objSpedC190.dValorRedBC + Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase, 2) '??????? esta certo????
                    objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS - Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase, 2) 'objSpedC190.dValorRedBC
                    objSpedC190.dValorIPI = objSpedC190.dValorIPI + tTribItemNF.dIPICredito
                    objSpedC190.dValorFCP = objSpedC190.dValorFCP + dValorFCP
                    objSpedC190.dValorFCPST = objSpedC190.dValorFCPST + dValorFCPST
                    objSpedC190.dValorFCPRet = objSpedC190.dValorFCPRet + dValorFCPRet
                    iAchou = 1
                    Exit For
                End If
                
            Next
    
            If iAchou = 0 Then
            
                Set objSpedC190 = New ClassSpedC190
                
                objSpedC190.iCSTICMS = iTipoTribCST_ICMS
                objSpedC190.sCFOP = tTribItemNF.sNaturezaOp
                objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota
                
                dValorTotalItem = Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + dICMSSubstValorExterno)
                If tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC Then
                    If tTribItemNF.dICMSBase <> 0 And tTribItemNF.dICMSBase > dValorTotalItem And gobjCRFAT.iNFImportacaoTribFlag03 = DESMARCADO Then
                        dValorTotalItem = tTribItemNF.dICMSBase
                    Else
                        If gobjCRFAT.iNFImportacaoTribFlag04 = DESMARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dICMSValor
                        If gobjCRFAT.iNFImportacaoTribFlag08 = MARCADO Then dValorTotalItem = dValorTotalItem - tTribItemNF.dValorOutrasDespesasItem
                        If gobjCRFAT.iNFImportacaoTribFlag05 = MARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dIIValor
                    
                        If gobjCRFAT.iNFImportacaoTribFlag12 = MARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dPISValor + tTribItemNF.dCOFINSValor
                    
                    End If
                End If
                
                objSpedC190.dValorOper = dValorTotalItem 'Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + tTribItemNF.dICMSSubstValor)
                dTotalRateioNF = dTotalRateioNF + Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + dICMSSubstValorExterno)
                
'                objSpedC190.dValorOper = (tItemNF.dQuantidade * tItemNF.dPrecoUnitario)
'                objSpedC190.dValorOper = objSpedC190.dValorOper + tTribItemNF.dICMSSubstValor + tTribItemNF.dIPIValor
'
'                If tNFiscal.dValorProdutos > 0 Then
'                    objSpedC190.dValorOper = objSpedC190.dValorOper + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'                    dTotalRateioNF = dTotalRateioNF + Round(((tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
''                     objSpedC190.dValorOper = objSpedC190.dValorOper - Round(((tNFiscal.dValorDesconto) * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'
'                End If
                
                objSpedC190.dValorBCICMS = tTribItemNF.dICMSBase
                'Esses valores já estão incluídos na base (Wagner)
                'se for nota de importacao interna ==> somar o ICMS a base
'                If tNFiscal.iTipoNFiscal = 119 Then
'                    objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS + tTribItemNF.dDespImpICMSBase
'                End If
'mari                objSpedC190.dValorICMS = tTribItemNF.dICMSValor
                objSpedC190.dValorICMS = tTribItemNF.dICMSCredito
'                If tNFiscal.iTipoNFiscal = 119 Then
'                    objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dDespImpICMSValor
'                End If
                objSpedC190.dValorBCICMSST = tTribItemNF.dICMSSubstBase
                objSpedC190.dValorICMSST = tTribItemNF.dICMSSubstValor
                objSpedC190.dValorRedBC = Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
                objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS - Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase) 'objSpedC190.dValorRedBC
                objSpedC190.dValorIPI = tTribItemNF.dIPICredito
                objSpedC190.dValorFCP = dValorFCP
                objSpedC190.dValorFCPST = dValorFCPST
                objSpedC190.dValorFCPRet = dValorFCPRet
                
                colSpedC190.Add objSpedC190
                
            End If
    
            iAchou = 0

            For Each objSpedC190UF In colSpedC190UF
            
                If objSpedC190UF.sUF = tNFiscal.sUFSubstTrib And objSpedC190UF.sCFOP = tTribItemNF.sNaturezaOp Then
                
                    objSpedC190UF.dValorBCICMSST = objSpedC190UF.dValorBCICMSST + tTribItemNF.dICMSSubstBase
                    objSpedC190UF.dValorICMSST = objSpedC190UF.dValorICMSST + tTribItemNF.dICMSSubstValor
                    
                    iAchou = 1
                    Exit For
                End If
                
            Next
    
            If iAchou = 0 Then
            
                Set objSpedC190UF = New ClassSpedC190UF
                
                objSpedC190UF.sUF = tNFiscal.sUFSubstTrib
                objSpedC190UF.sCFOP = tTribItemNF.sNaturezaOp
                
                objSpedC190UF.dValorBCICMSST = tTribItemNF.dICMSSubstBase
                objSpedC190UF.dValorICMSST = tTribItemNF.dICMSSubstValor
                                
                colSpedC190UF.Add objSpedC190UF
                
            End If
    
            'soma os valores de ICMS dos registros C100 de saida que estiverem com data de saida dentro do periodo para poder colocar
            'posteriormente no registro E110
            If Len(Trim(sData)) > 0 Then
                If CDate(sData) >= dtDataIni And CDate(sData) <= dtDataFim And tTipoDocInfo.iTipo = 2 Then
    'mario                dValDebICMS_objSpedC190 = dValDebICMS_objSpedC190 + tTribItemNF.dICMSValor
                    dValDebICMS_objSpedC190 = dValDebICMS_objSpedC190 + tTribItemNF.dICMSCredito
                End If
            End If
    
            'soma os valores de ICMS dos registros C100 de entrada que estiverem com data de entrada dentro do periodo para poder colocar
            'posteriormente no registro E110
            If Len(Trim(sData)) > 0 Then
                If CDate(sData) >= dtDataIni And CDate(sData) <= dtDataFim And (tTipoDocInfo.iTipo = 1 Or tTipoDocInfo.iTipo = 3) Then
    '                dValCredICMS_objSpedC190 = dValCredICMS_objSpedC190 + tTribItemNF.dICMSValor
                    dValCredICMS_objSpedC190 = dValCredICMS_objSpedC190 + tTribItemNF.dICMSCredito
    '                If tNFiscal.iTipoNFiscal = 119 Then
    '                    dValCredICMS_objSpedC190 = dValCredICMS_objSpedC190 + tTribItemNF.dDespImpICMSValor
    '                End If
                    
                End If
            End If
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203988
    
        Loop

'        'trata os complementos das notas fiscais
'
'        tTribItemNF.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'        lErro = Comando_Executar(alComando(5), "SELECT T.Tipo, TipoTribCST, NaturezaOP, ICMSBase, ICMSAliquota, ICMSValor, ICMSSubstBase, ICMSSubstValor, IPI.CSTEntrada, IPI.CSTSaida, " & _
'        "IPIValor, ICMSPercRedBase FROM TributacaoComplNF AS T, tiposTribICMS AS ICMS, TiposTribIPI AS IPI  WHERE T.NumIntDoc  = ? AND ICMS.Tipo = ICMSTipo AND IPI.Tipo = IPITipo AND (T.Tipo = 1 Or T.Tipo=2 Or T.Tipo = 4)", _
'        iTipoRegistro, iTipoTribCST_ICMS, tTribItemNF.sNaturezaOp, tTribItemNF.dICMSBase, tTribItemNF.dICMSAliquota, tTribItemNF.dICMSValor, tTribItemNF.dICMSSubstBase, tTribItemNF.dICMSSubstValor, _
'        iIPICSTEntrada, iIPICSTSaida, tTribItemNF.dIPIValor, tTribItemNF.dICMSPercRedBase, tNFiscal.lNumIntDoc)
'        If lErro <> AD_SQL_SUCESSO Then gError 204190
'
'        lErro = Comando_BuscarPrimeiro(alComando(5))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204191
'
'        Do While lErro = AD_SQL_SUCESSO
'
'
'            iAchou = 0
'
'            For Each objSpedC190 In colSpedC190
'
'                If objSpedC190.iCSTICMS = iTipoTribCST_ICMS And objSpedC190.sCFOP = tTribItemNF.sNaturezaOp And objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota Then
'                    If iTipoRegistro = ITEM_TIPO_FRETE Then
'                        objSpedC190.dValorOper = objSpedC190.dValorOper + tNFiscal.dValorFrete
'                    ElseIf iTipoRegistro = ITEM_TIPO_SEGURO Then
'                        objSpedC190.dValorOper = objSpedC190.dValorOper + tNFiscal.dValorSeguro
'                    ElseIf iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                        objSpedC190.dValorOper = objSpedC190.dValorOper + tNFiscal.dValorOutrasDespesas
''                    ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''                        objSpedC190.dValorOper = objSpedC190.dValorOper - tNFiscal.dValorDesconto
'                    End If
'                    If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                        objSpedC190.dValorOper = objSpedC190.dValorOper + tTribItemNF.dICMSSubstValor + tTribItemNF.dIPIValor
'
'                        objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS + tTribItemNF.dICMSBase
'                        objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dICMSValor
'                        objSpedC190.dValorBCICMSST = objSpedC190.dValorBCICMSST + tTribItemNF.dICMSSubstBase
'                        objSpedC190.dValorICMSST = objSpedC190.dValorICMSST + tTribItemNF.dICMSSubstValor
'                        objSpedC190.dValorRedBC = objSpedC190.dValorRedBC + (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'                        objSpedC190.dValorIPI = objSpedC190.dValorIPI + tTribItemNF.dIPIValor
'
''                    ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''                        objSpedC190.dValorOper = objSpedC190.dValorOper - tTribItemNF.dICMSSubstValor
''
''                        objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS - tTribItemNF.dICMSBase
''                        objSpedC190.dValorICMS = objSpedC190.dValorICMS - tTribItemNF.dICMSValor
''                        objSpedC190.dValorBCICMSST = objSpedC190.dValorBCICMSST - tTribItemNF.dICMSSubstBase
''                        objSpedC190.dValorICMSST = objSpedC190.dValorICMSST - tTribItemNF.dICMSSubstValor
''                        objSpedC190.dValorRedBC = objSpedC190.dValorRedBC - (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
'                    End If
'
'
'                    iAchou = 1
'                    Exit For
'                End If
'
'            Next
'
'            If iAchou = 0 Then
'
'                Set objSpedC190 = New ClassSpedC190
'
'                objSpedC190.iCSTICMS = iTipoTribCST_ICMS
'                objSpedC190.sCFOP = tTribItemNF.sNaturezaOp
'                objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota
'
'                If iTipoRegistro = ITEM_TIPO_FRETE Then
'                    objSpedC190.dValorOper = tNFiscal.dValorFrete
'                ElseIf iTipoRegistro = ITEM_TIPO_SEGURO Then
'                    objSpedC190.dValorOper = tNFiscal.dValorSeguro
'                ElseIf iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                    objSpedC190.dValorOper = tNFiscal.dValorOutrasDespesas
''                ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''                    objSpedC190.dValorOper = objSpedC190.dValorOper - tNFiscal.dValorDesconto
'                End If
'
'
'                If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                    objSpedC190.dValorOper = objSpedC190.dValorOper + tTribItemNF.dICMSSubstValor + tTribItemNF.dIPIValor
'
'                    objSpedC190.dValorBCICMS = tTribItemNF.dICMSBase
'                    objSpedC190.dValorICMS = tTribItemNF.dICMSValor
'                    objSpedC190.dValorBCICMSST = tTribItemNF.dICMSSubstBase
'                    objSpedC190.dValorICMSST = tTribItemNF.dICMSSubstValor
'                    objSpedC190.dValorRedBC = (tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'                    objSpedC190.dValorIPI = tTribItemNF.dIPIValor
'
''                ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''                    objSpedC190.dValorOper = objSpedC190.dValorOper - tTribItemNF.dICMSSubstValor
''
''                    objSpedC190.dValorBCICMS = -tTribItemNF.dICMSBase
''                    objSpedC190.dValorICMS = -tTribItemNF.dICMSValor
''                    objSpedC190.dValorBCICMSST = -tTribItemNF.dICMSSubstBase
''                    objSpedC190.dValorICMSST = -tTribItemNF.dICMSSubstValor
''                    objSpedC190.dValorRedBC = -(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
'
'                End If
'
'                colSpedC190.Add objSpedC190
'
'            End If
'
'            If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'
'                'soma os valores de ICMS dos registros C100 de saida que estiverem com data de saida dentro do periodo para poder colocar
'                'posteriormente no registro E110
'                If CDate(sData) >= dtDataIni And CDate(sData) <= dtDataFim And tTipoDocInfo.iTipo = 2 Then
'                    dValDebICMS_objSpedC190 = dValDebICMS_objSpedC190 + tTribItemNF.dICMSValor
'                End If
'
'                'soma os valores de ICMS dos registros C100 de entrada que estiverem com data de entrada dentro do periodo para poder colocar
'                'posteriormente no registro E110
'                If CDate(sData) >= dtDataIni And CDate(sData) <= dtDataFim And (tTipoDocInfo.iTipo = 1 Or tTipoDocInfo.iTipo = 3) Then
'                    dValCredICMS_objSpedC190 = dValCredICMS_objSpedC190 + tTribItemNF.dICMSValor
'                End If
'
''            ElseIf iTipoRegistro = ITEM_TIPO_DESCONTO Then
''
''                'soma os valores de ICMS dos registros C100 de saida que estiverem com data de saida dentro do periodo para poder colocar
''                'posteriormente no registro E110
''                If CDate(sData) >= dtDataIni And CDate(sData) <= dtDataFim And tTipoDocInfo.iTipo = 2 Then
''                    dValDebICMS_objSpedC190 = dValDebICMS_objSpedC190 - tTribItemNF.dICMSValor
''                End If
''
''                'soma os valores de ICMS dos registros C100 de entrada que estiverem com data de entrada dentro do periodo para poder colocar
''                'posteriormente no registro E110
''                If CDate(sData) >= dtDataIni And CDate(sData) <= dtDataFim And (tTipoDocInfo.iTipo = 1 Or tTipoDocInfo.iTipo = 3) Then
''                    dValCredICMS_objSpedC190 = dValCredICMS_objSpedC190 - tTribItemNF.dICMSValor
''                End If
'
'            End If
'
'            If iTipoRegistro = ITEM_TIPO_FRETE Or iTipoRegistro = ITEM_TIPO_SEGURO Or iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'
'                Set objSpedE510Aux = New ClassSpedE510Aux
'
'                objSpedE510Aux.sCFOP = tTribItemNF.sNaturezaOp
'                objSpedE510Aux.iCSTIPI = IIf(tTipoDocInfo.iTipo = 2, Format(iIPICSTSaida, "00"), Format(iIPICSTEntrada, "00"))
'
'                objSpedE510Aux.dValorIPI = tTribItemNF.dIPIValor
'                objSpedE510Aux.dValorBCIPI = tTribItemNF.dIPIBaseCalculo
'
'                If iTipoRegistro = ITEM_TIPO_FRETE Then
'                    objSpedE510Aux.dValorItemNF = tNFiscal.dValorFrete
'                ElseIf iTipoRegistro = ITEM_TIPO_SEGURO Then
'                    objSpedE510Aux.dValorItemNF = tNFiscal.dValorSeguro
'                ElseIf iTipoRegistro = ITEM_TIPO_OUTRAS_DESP Then
'                    objSpedE510Aux.dValorItemNF = tNFiscal.dValorOutrasDespesas
'                End If
'
'                colSpedE510Aux.Add objSpedE510Aux
'
'            End If
'
'
'            lErro = Comando_BuscarProximo(alComando(5))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204192
'
'        Loop



        For Each objSpedE510Aux In colSpedE510Aux
            dValorItemNFTotal = dValorItemNFTotal + objSpedE510Aux.dValorItemNF
        Next
        
        
        'acumula os valores de ItemNF para posteriormente poder criar os registro E510
        For Each objSpedE510Aux In colSpedE510Aux
            
            iAchou = 0
        
            For Each objSpedE510 In colSpedE510
            
                If objSpedE510Aux.sCFOP = objSpedE510.sCFOP And objSpedE510Aux.iCSTIPI = objSpedE510.iCSTIPI Then
                            
                    objSpedE510.dValorBCIPI = objSpedE510.dValorBCIPI + objSpedE510Aux.dValorBCIPI
                    objSpedE510.dValorIPI = objSpedE510.dValorIPI + objSpedE510Aux.dValorIPI
                    objSpedE510.dValorContabilIPI = objSpedE510.dValorContabilIPI + objSpedE510Aux.dValorItemNF + objSpedE510Aux.dValorIPI
                    'O VALOR DOS ITENS JÁ VEM COM DESCONTO. Ass: Wagner
'                    If dValorItemNFTotal > 0 Then
'                        objSpedE510.dValorContabilIPI = objSpedE510.dValorContabilIPI - (tNFiscal.dValorDesconto * objSpedE510Aux.dValorItemNF / dValorItemNFTotal)
'                    End If
                    
                    iAchou = 1
                    Exit For
                End If
            
            Next
            
            If iAchou = 0 Then
                
                Set objSpedE510 = New ClassSpedE510
                
                objSpedE510.sCFOP = objSpedE510Aux.sCFOP
                objSpedE510.iCSTIPI = objSpedE510Aux.iCSTIPI
                objSpedE510.dValorIPI = objSpedE510Aux.dValorIPI
                objSpedE510.dValorBCIPI = objSpedE510Aux.dValorBCIPI
                objSpedE510.dValorContabilIPI = objSpedE510Aux.dValorItemNF + objSpedE510Aux.dValorIPI
                
                'O VALOR DOS ITENS JÁ VEM COM DESCONTO. Ass: Wagner
'                If dValorItemNFTotal > 0 Then
'                    objSpedE510.dValorContabilIPI = objSpedE510.dValorContabilIPI - (tNFiscal.dValorDesconto * objSpedE510Aux.dValorItemNF / dValorItemNFTotal)
'                End If
                
                colSpedE510.Add objSpedE510
                
            End If
        
        Next
        
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C170 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C170:

   SpedFiscal_C170 = gErr

   Select Case gErr
    
        Case 203971
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203972, 203973, 203988
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)
    
        Case 203974, 203975
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOITEMNF", gErr)
    
        Case 203976 To 203981, 204190 To 204192, 204206
    
        Case 203982, 203986
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203983, 203987
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case 203984, 203985
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBUTACAOMOVTO", gErr)
        
        Case 204190 To 204192
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLNF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203989)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Private Function SpedFiscal_C190(tTributacaoNF As typeTributacaoNF, objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC190 As ClassSpedRegistro, ByVal lSeqPaiNF As Long, ByRef tNFiscal As typeNFiscal, ByRef tTipoDocInfo As typeTipoDocInfo, ByVal colSpedC190 As Collection, ByVal iNFe As Integer, ByVal dTotalRateioNF As Double, ByVal objFisSpedAux As ClassFISSpedAux) As Long


Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objSpedC190 As ClassSpedC190

On Error GoTo Erro_SpedFiscal_C190
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203990
    Next

    'se for nota fiscal (codigo 01) ou nota do produto (codigo 04)
    If tTipoDocInfo.iModeloArqICMS = 1 Or tTipoDocInfo.iModeloArqICMS = 4 Then 'Or iNFe = 1 Then
    
        'Está errado porque falta o IPI e ICMS ST (Wagner)
'        If tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto <> dTotalRateioNF And colSpedC190.Count > 0 Then
'
'            Set objSpedC190 = colSpedC190.Item(1)
'
'            objSpedC190.dValorOper = objSpedC190.dValorOper + Round(tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto - dTotalRateioNF, 2)
'
'        End If
        
        For Each objSpedC190 In colSpedC190
        
            If objSpedC190.dValorOper > 0 Or objSpedC190.dValorBCICMS > 0 Or objSpedC190.dValorICMS > 0 Or _
               objSpedC190.dValorBCICMSST > 0 Or objSpedC190.dValorICMSST > 0 Or objSpedC190.dValorRedBC > 0 Or _
               objSpedC190.dValorIPI > 0 Then
        
                objSpedFiscal.sTipo = "C190"
                objSpedRegistroC190.sTipo = objSpedFiscal.sTipo
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistroC190.iCount = objSpedRegistroC190.iCount + 1
    
                objSpedFiscal.sRegistro = "|C190|" & Format(objSpedC190.iCSTICMS, "000") & "|" & Format(objSpedC190.sCFOP, "0000") & "|" & Format(objSpedC190.dAliquotaICMS * 100, "fixed") & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorOper <= 0, 0, Format(objSpedC190.dValorOper, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorBCICMS <= 0, 0, Format(objSpedC190.dValorBCICMS, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorICMS <= 0, 0, Format(objSpedC190.dValorICMS, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorBCICMSST <= 0, 0, Format(objSpedC190.dValorBCICMSST, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorICMSST <= 0, 0, Format(objSpedC190.dValorICMSST, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorRedBC <= 0, 0, Format(objSpedC190.dValorRedBC, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorIPI <= 0, 0, Format(objSpedC190.dValorIPI, "fixed")) & "||"
                
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203991
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
                If lErro <> AD_SQL_SUCESSO Then gError 203992
                
                If giVersao > 12 And iNFe = 1 And tTipoDocInfo.iModeloArqICMS = 1 Then
                
                     objSpedFiscal.sTipo = "C191"
                    objFisSpedAux.objSpedRegistroC191.sTipo = objSpedFiscal.sTipo
                    
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objFisSpedAux.objSpedRegistroC191.iCount = objFisSpedAux.objSpedRegistroC191.iCount + 1
                    
                    objSpedFiscal.sRegistro = "|C191|" & IIf(objSpedC190.dValorFCP <= 0, 0, Format(objSpedC190.dValorFCP, "fixed")) & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorFCP <= 0, 0, Format(objSpedC190.dValorFCPST, "fixed")) & "|"
                    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorFCP <= 0, 0, Format(objSpedC190.dValorFCPRet, "fixed")) & "|"
                
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203991
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
                    If lErro <> AD_SQL_SUCESSO Then gError 203992
                
                End If
                
                
            End If
            
        Next
    
    End If


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C190 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C190:

   SpedFiscal_C190 = gErr

   Select Case gErr
    
        Case 203990
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203991
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203992
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203993)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Private Function SpedFiscal_C405(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC405 As ClassSpedRegistro, ByVal lSeqECF As Long, ByRef tArqMestre As typeLojaArqFisMestre, ByVal lSeqReducaoZ As Long, ByVal dVendaBruta As Double) As Long


Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_SpedFiscal_C405
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203994
    Next

    objSpedFiscal.sTipo = "C405"
    objSpedRegistroC405.sTipo = objSpedFiscal.sTipo
    
    objSpedRegistroC405.iCount = objSpedRegistroC405.iCount + 1

    objSpedFiscal.sRegistro = "|C405|" & Format(tArqMestre.dtData, "ddmmyyyy") & "|" & tArqMestre.lCRO & "|" & tArqMestre.lCRZ & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tArqMestre.lCOOFim & "|" & IIf(tArqMestre.dGrandeTotal = 0, 0, Format(tArqMestre.dGrandeTotal, "fixed")) & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(dVendaBruta = 0, 0, Format(dVendaBruta, "fixed")) & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203995
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, lSeqReducaoZ, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqECF)
    If lErro <> AD_SQL_SUCESSO Then gError 203996


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C405 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C405:

   SpedFiscal_C405 = gErr

   Select Case gErr
    
        Case 203994
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203995
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203996
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203997)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_C420(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC420 As ClassSpedRegistro, ByVal lSeqReducaoZ As Long, ByRef tArqMestre As typeLojaArqFisMestre, ByVal iFilialEmpresa As Integer, ByVal iCodigoECF As Integer, ByVal objSpedRegistro0200 As ClassSpedRegistro, ByVal objSpedRegistro0205 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal objSpedRegistro0190 As ClassSpedRegistro, ByVal objSpedRegistro0220 As ClassSpedRegistro, ByVal colUM As Collection, ByVal objSpedRegistroC425 As ClassSpedRegistro, dVendaBruta As Double, ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim tLojaReducao As typeLojaReducao
Dim tLojaReducaoAliquota As typeLojaReducaoAliquota
Dim sCodigo As String
Dim dValorTotal As Double
Dim lSeqC420 As Long


On Error GoTo Erro_SpedFiscal_C420
        
    dVendaBruta = 0
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203998
    Next

    lErro = Comando_Executar(alComando(2), "SELECT AcrescimoICMS, AcrescimoISS, AcrescimoNaoFiscal, CancelamentoICMS, CancelamentoISS, CancelamentoNaoFiscal, DescontoICMS, DescontoISS, DescontoNaoFiscal, IsentoICMS, IsentoISS, NaoIncideICMS, NaoIncideISS, SubstTribICMS, SubstTribISS, TotalNaoFiscal FROM LojaReducao WHERE CodECF = ? And FilialEmpresa = ? And DataMovimento = ?", _
    tLojaReducao.dAcrescimoICMS, tLojaReducao.dAcrescimoISS, tLojaReducao.dAcrescimoNaoFiscal, tLojaReducao.dCancelamentoICMS, tLojaReducao.dCancelamentoISS, tLojaReducao.dCancelamentoNaoFiscal, tLojaReducao.dDescontoICMS, tLojaReducao.dDescontoISS, tLojaReducao.dDescontoNaoFiscal, tLojaReducao.dIsentoICMS, tLojaReducao.dIsentoISS, tLojaReducao.dNaoIncideICMS, tLojaReducao.dNaoIncideISS, tLojaReducao.dSubstTribICMS, tLojaReducao.dSubstTribISS, tLojaReducao.dTotalNaoFiscal, iCodigoECF, iFilialEmpresa, tArqMestre.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 203999

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204000

    If tLojaReducao.dAcrescimoICMS > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1


        objSpedFiscal.sRegistro = "|C420|AT|" & IIf(tLojaReducao.dAcrescimoICMS = 0, 0, Format(tLojaReducao.dAcrescimoICMS, "fixed")) & "|||"

        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204001
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204002

    End If
    
    If tLojaReducao.dAcrescimoISS > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|AS|" & IIf(tLojaReducao.dAcrescimoISS = 0, 0, Format(tLojaReducao.dAcrescimoISS, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204003
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204004

    End If
    
    If tLojaReducao.dAcrescimoNaoFiscal > 0 Then
    
        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|AO|" & IIf(tLojaReducao.dAcrescimoNaoFiscal = 0, 0, Format(tLojaReducao.dAcrescimoNaoFiscal, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204005
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204006

    End If
    
    If tLojaReducao.dCancelamentoICMS > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|Can-T|" & IIf(tLojaReducao.dCancelamentoICMS = 0, 0, Format(tLojaReducao.dCancelamentoICMS, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204007
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204008

        dVendaBruta = dVendaBruta + tLojaReducao.dCancelamentoICMS

    End If
    
    If tLojaReducao.dCancelamentoISS > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|Can-S|" & IIf(tLojaReducao.dCancelamentoISS = 0, 0, Format(tLojaReducao.dCancelamentoISS, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204009
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204010

        dVendaBruta = dVendaBruta + tLojaReducao.dCancelamentoISS

    End If

    If tLojaReducao.dCancelamentoNaoFiscal > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|Can-O|" & IIf(tLojaReducao.dCancelamentoNaoFiscal = 0, 0, Format(tLojaReducao.dCancelamentoNaoFiscal, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204011
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204012

    End If
    
    If tLojaReducao.dDescontoICMS > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|DT|" & IIf(tLojaReducao.dDescontoICMS = 0, 0, Format(tLojaReducao.dDescontoICMS, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204013
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204014

        dVendaBruta = dVendaBruta + tLojaReducao.dDescontoICMS

    End If

    If tLojaReducao.dDescontoISS > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|DS|" & IIf(tLojaReducao.dDescontoISS = 0, 0, Format(tLojaReducao.dDescontoISS, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204015
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204016

        dVendaBruta = dVendaBruta + tLojaReducao.dDescontoISS

    End If
    
    If tLojaReducao.dDescontoNaoFiscal > 0 Then
    
        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|DO|" & IIf(tLojaReducao.dDescontoNaoFiscal = 0, 0, Format(tLojaReducao.dDescontoNaoFiscal, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204017
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204018
    
    End If
    

'    objSpedFiscal.sTipo = "C420"
'    objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    lSeqC420 = objSpedFiscal.lSeq
'    objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1

'    objSpedFiscal.sRegistro = "|C420|I1|" & IIf(tLojaReducao.dIsentoICMS = 0, 0, Format(tLojaReducao.dIsentoICMS, "fixed")) & "|||"

'    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204019
'
'    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
'    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
'    If lErro <> AD_SQL_SUCESSO Then gError 204020

    dValorTotal = 0

    '******************* registro 425 (referente a isencao) *******************
    lErro = SpedFiscal_C425(iCodigoECF, iFilialEmpresa, tArqMestre, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, objSpedRegistroC425, "I1", dValorTotal, objFisSpedAux)
    If lErro <> SUCESSO Then gError 204021
    
    If dValorTotal > 0 Then
    
        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|I1|" & IIf(dValorTotal = 0, 0, Format(dValorTotal, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204019
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, lSeqC420, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204020
    
        dVendaBruta = dVendaBruta + dValorTotal
    
    Else
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq - 1
    
    End If
    
    If tLojaReducao.dIsentoISS Then
    
        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|IS1|" & IIf(tLojaReducao.dIsentoISS = 0, 0, Format(tLojaReducao.dIsentoISS, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204022
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204023

        dVendaBruta = dVendaBruta + tLojaReducao.dIsentoISS

    End If
    
'    objSpedFiscal.sTipo = "C420"
'    objSpedRegistroC420.sTipo = objSpedFiscal.sTipo

    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    lSeqC420 = objSpedFiscal.lSeq
'    objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1

'    objSpedFiscal.sRegistro = "|C420|N1|" & IIf(tLojaReducao.dNaoIncideICMS = 0, 0, Format(tLojaReducao.dNaoIncideICMS, "fixed")) & "|||"

'    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204024
    
'    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
'    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
'    If lErro <> AD_SQL_SUCESSO Then gError 204025

    dValorTotal = 0

    '******************* registro 425 (referente a nao tributado) *******************
    lErro = SpedFiscal_C425(iCodigoECF, iFilialEmpresa, tArqMestre, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, objSpedRegistroC425, "N1", dValorTotal, objFisSpedAux)
    If lErro <> SUCESSO Then gError 204026

    If dValorTotal > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
    
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|N1|" & IIf(dValorTotal = 0, 0, Format(dValorTotal, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204024
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, lSeqC420, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204025

        dVendaBruta = dVendaBruta + dValorTotal

    Else
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq - 1

    End If

    If tLojaReducao.dNaoIncideISS > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|NS1|" & IIf(tLojaReducao.dNaoIncideISS = 0, 0, Format(tLojaReducao.dNaoIncideISS, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204027
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204028

        dVendaBruta = dVendaBruta + tLojaReducao.dNaoIncideISS

    End If

'    objSpedFiscal.sTipo = "C420"
'    objSpedRegistroC420.sTipo = objSpedFiscal.sTipo

    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    lSeqC420 = objSpedFiscal.lSeq
'    objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1

'    objSpedFiscal.sRegistro = "|C420|F1|" & IIf(tLojaReducao.dSubstTribICMS = 0, 0, Format(tLojaReducao.dSubstTribICMS, "fixed")) & "|||"
'
'    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204029
'
'    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
'    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
'    If lErro <> AD_SQL_SUCESSO Then gError 204030

    dValorTotal = 0

    '******************* registro 425 (referente a substituicao) *******************
    lErro = SpedFiscal_C425(iCodigoECF, iFilialEmpresa, tArqMestre, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, objSpedRegistroC425, "F1", dValorTotal, objFisSpedAux)
    If lErro <> SUCESSO Then gError 204031

    If dValorTotal > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
    
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|F1|" & IIf(dValorTotal = 0, 0, Format(dValorTotal, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204029
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, lSeqC420, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204030

        dVendaBruta = dVendaBruta + dValorTotal

     Else
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq - 1

   End If
    
    If tLojaReducao.dNaoIncideISS > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|FS1|" & IIf(tLojaReducao.dSubstTribISS = 0, 0, Format(tLojaReducao.dSubstTribISS, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204032
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204033

        dVendaBruta = dVendaBruta + tLojaReducao.dSubstTribISS

    End If
    

    If tLojaReducao.dTotalNaoFiscal > 0 Then

        objSpedFiscal.sTipo = "C420"
        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
        objSpedFiscal.sRegistro = "|C420|OPNF|" & IIf(tLojaReducao.dTotalNaoFiscal = 0, 0, Format(tLojaReducao.dTotalNaoFiscal, "fixed")) & "|||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204034
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
        If lErro <> AD_SQL_SUCESSO Then gError 204035

    End If

    lErro = Comando_Executar(alComando(3), "SELECT Aliquota, ISS, Totalizador FROM LojaReducaoAliquota WHERE CodECF = ? And FilialEmpresa = ? And Data = ?", _
    tLojaReducaoAliquota.dAliquota, tLojaReducaoAliquota.iISS, tLojaReducaoAliquota.dTotalizador, iCodigoECF, iFilialEmpresa, tArqMestre.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 204036

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204037
        
    Do While lErro = AD_SQL_SUCESSO

        If tLojaReducaoAliquota.iISS = ALIQUOTA_ISS Then
            sCodigo = "S"
        Else
            sCodigo = "T"
        End If

        sCodigo = sCodigo & Format(tLojaReducaoAliquota.dAliquota * 100, "0000")
        
'        objSpedFiscal.sTipo = "C420"
'        objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        lSeqC420 = objSpedFiscal.lSeq
'        objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1

'        objSpedFiscal.sRegistro = "|C420|" & sCodigo & "|" & IIf(tLojaReducaoAliquota.dTotalizador = 0, 0, Format(tLojaReducaoAliquota.dTotalizador, "fixed")) & "|||"
'
'        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204038
'
'        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
'        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
'        If lErro <> AD_SQL_SUCESSO Then gError 204039

        dValorTotal = 0

        '******************* registro C425 (referente a tributacao normal) *******************
        lErro = SpedFiscal_C425(iCodigoECF, iFilialEmpresa, tArqMestre, objSpedFiscal, objSpedRegistro0200, objSpedRegistro0205, dtDataIni, dtDataFim, objSpedRegistro0190, objSpedRegistro0220, colUM, objSpedRegistroC425, sCodigo, dValorTotal, objFisSpedAux)
        If lErro <> SUCESSO Then gError 204041

        'se o valortotal for 0 nao vai imprimir o C420
        If dValorTotal > 0 Then


            objSpedFiscal.sTipo = "C420"
            objSpedRegistroC420.sTipo = objSpedFiscal.sTipo
            
            objSpedRegistroC420.iCount = objSpedRegistroC420.iCount + 1
    
            objSpedFiscal.sRegistro = "|C420|01" & sCodigo & "|" & IIf(dValorTotal = 0, 0, Format(dValorTotal, "fixed")) & "|01||"
    
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 204038
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, lSeqC420, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqReducaoZ)
            If lErro <> AD_SQL_SUCESSO Then gError 204039
    
            dVendaBruta = dVendaBruta + dValorTotal
    
        Else
        
            objSpedFiscal.lSeq = objSpedFiscal.lSeq - 1
        
        End If

        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204040


    Loop


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C420 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C420:

   SpedFiscal_C420 = gErr

   Select Case gErr
    
        Case 203998
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203999, 204000
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJAREDUCAO", gErr)
    
        Case 204001, 204003, 204005, 204007, 204009, 204011, 204013, 204015, 204017, 204019, 204022, 204024, 204027, 204029, 204032, 204034, 204038
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204002, 204004, 204006, 204008, 204010, 204012, 204014, 204016, 204018, 204020, 204023, 204025, 204028, 204030, 204033, 204035, 204039
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case 204021, 204026, 204031, 204041
        
        Case 204036, 204037, 204040
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJAREDUCAOALIQUOTA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204042)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_D001(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroD001 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date) As Long


Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_SpedFiscal_D001
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204043
    Next


    'se houver uma nota fiscal de entrada ou saida e for nota de conhecimento de frete ou telecomunicacoes
    lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc  " & _
    " From NFiscal AS N, TiposDocInfo AS T WHERE TipoNFiscal = Codigo AND ((Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) Or ((Tipo=1 Or Tipo=3) AND DataEntrada >= ? AND DataEntrada <= ?))  AND (ModeloArqICMS = 8 Or ModeloArqICMS = 22) AND FilialEmpresa = ? ", _
    tNFiscal.lNumIntDoc, dtDataIni, dtDataFim, dtDataIni, dtDataFim, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 204044
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204045
    
    If lErro = AD_SQL_SUCESSO Then
    
        '******************* registro D001 *******************
        objSpedFiscal.sTipo = "D001"
        objSpedRegistroD001.sTipo = objSpedFiscal.sTipo
    
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroD001.iCount = objSpedRegistroD001.iCount + 1
        
        objSpedFiscal.sRegistro = "|D001|0|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204046
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 204047
    
    Else
    
        '******************* registro D001 *******************
        objSpedFiscal.sTipo = "D001"
        objSpedRegistroD001.sTipo = objSpedFiscal.sTipo
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroD001.iCount = objSpedRegistroD001.iCount + 1
        
        objSpedFiscal.sRegistro = "|D001|1|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204048
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 204049
    
    End If


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_D001 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_D001:

   SpedFiscal_D001 = gErr

   Select Case gErr
    
        Case 204043
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204044, 204045
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
    
        Case 204046, 204048
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204047, 204049
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204050)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_D990(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroD990 As ClassSpedRegistro, ByVal objSpedRegistroD001 As ClassSpedRegistro, ByVal objSpedRegistroD100 As ClassSpedRegistro, ByVal objSpedRegistroD190 As ClassSpedRegistro, ByVal objSpedRegistroD500 As ClassSpedRegistro, ByVal objSpedRegistroD590 As ClassSpedRegistro) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lTotal As Long

On Error GoTo Erro_SpedFiscal_D990
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204051
    Next

    objSpedFiscal.sTipo = "D990"
    objSpedRegistroD990.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroD990.iCount = objSpedRegistroD990.iCount + 1

    lTotal = objSpedRegistroD001.iCount + objSpedRegistroD100.iCount + objSpedRegistroD190.iCount + objSpedRegistroD500.iCount + objSpedRegistroD590.iCount + objSpedRegistroD990.iCount
    
    objSpedFiscal.sRegistro = "|D990|" & lTotal & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204052
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204053


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_D990 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_D990:

   SpedFiscal_D990 = gErr

   Select Case gErr
    
        Case 204051
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204052
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204053
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204054)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_E110(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE110 As ClassSpedRegistro, ByVal objSpedRegistroE111 As ClassSpedRegistro, ByVal objFilialEmpresa As AdmFiliais, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal dValDebICMS_objSpedC190 As Double, ByVal dValDebICMS_objSpedC490 As Double, ByVal dValCredICMS_objSpedC190 As Double, ByVal dValCredICMS_objSpedC590 As Double, ByVal dValCredICMS_objSpedD190 As Double, ByVal dValCredICMS_objSpedD590 As Double, dValorICMSRecolher As Double, lSeqPai As Long, ByVal lSeqE100 As Long) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim dTotAjCreditos As Double
Dim dTotAjDebitos As Double
Dim dTotEstCreditos As Double
Dim dTotEstDebitos As Double
Dim dValTotDebitos As Double
Dim dValTotCreditos As Double
Dim dSaldoCredorInicial As Double
Dim dSaldoApurado As Double
Dim dSaldoCredTransp As Double
Dim dTotDeducoes As Double

On Error GoTo Erro_SpedFiscal_E110
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204055
    Next


    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroE110.iCount = objSpedRegistroE110.iCount + 1
    lSeqPai = objSpedFiscal.lSeq
    
    '*************** registro E111 **********************
    lErro = SpedFiscal_E111(objSpedFiscal, objSpedRegistroE111, objFilialEmpresa, dtDataIni, dtDataFim, lSeqPai, dTotAjDebitos, dTotEstCreditos, dTotAjCreditos, dTotEstDebitos, dTotDeducoes)
    If lErro <> SUCESSO Then gError 204056
    
    dValTotDebitos = dValDebICMS_objSpedC190 + dValDebICMS_objSpedC490
    
    dValTotCreditos = dValCredICMS_objSpedC190 + dValCredICMS_objSpedC590 + dValCredICMS_objSpedD190 + dValCredICMS_objSpedD590
    
    'Verifica se o Item de apuração ICMS está cadastrada
    lErro = Comando_Executar(alComando(2), "SELECT SaldoCredorInicial FROM RegApuracaoICMS WHERE  DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ? ", _
    dSaldoCredorInicial, dtDataIni, dtDataFim, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 204057
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204058
    
    If lErro <> AD_SQL_SUCESSO Then dSaldoCredorInicial = 0
    
    dSaldoApurado = dValTotDebitos + dTotAjDebitos + dTotEstCreditos - dValTotCreditos - dTotAjCreditos - dTotEstDebitos - dSaldoCredorInicial
    
    dValorICMSRecolher = IIf(dSaldoApurado - dTotDeducoes <= 0, 0, dSaldoApurado - dTotDeducoes)
    dSaldoCredTransp = IIf(dSaldoApurado - dTotDeducoes <= 0, Abs(dSaldoApurado - dTotDeducoes), 0)
    
    If dSaldoApurado < 0 Then dSaldoApurado = 0
    
    objSpedFiscal.sTipo = "E110"
    objSpedRegistroE110.sTipo = objSpedFiscal.sTipo
    
    
    objSpedFiscal.sRegistro = "|E110|" & Format(dValTotDebitos, "fixed") & "|0|" & Format(dTotAjDebitos, "fixed") & "|" & Format(dTotEstCreditos, "fixed") & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dValTotCreditos, "fixed") & "|0|" & Format(dTotAjCreditos, "fixed") & "|" & Format(dTotEstDebitos, "fixed") & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dSaldoCredorInicial, "fixed") & "|" & Format(dSaldoApurado, "fixed") & "|" & Format(dTotDeducoes, "fixed") & "|" & Format(dValorICMSRecolher, "fixed") & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dSaldoCredTransp, "fixed") & "|0|"
    
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204060

    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai) VALUES (?,?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, lSeqPai, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqE100)
    If lErro <> AD_SQL_SUCESSO Then gError 204061

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E110 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E110:

   SpedFiscal_E110 = gErr

   Select Case gErr
    
        Case 204055
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204056
    
        Case 204057, 204058
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMS", gErr)
    
       Case 204059
            Call Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOICMS_NAO_CADASTRADO", gErr)
    
        Case 204060
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204061
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204062)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Private Function SpedFiscal_E520(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE520 As ClassSpedRegistro, ByVal objSpedRegistroE530 As ClassSpedRegistro, ByVal objFilialEmpresa As AdmFiliais, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal dValDebIPI As Double, ByVal dValCredIPI As Double, ByVal lSeqE500 As Long) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim dTotAjCreditos As Double
Dim dTotAjDebitos As Double
Dim dSaldoCredorInicial As Double
Dim dSaldoApurado As Double
Dim dSaldoCredTransp As Double
Dim lSeqPai As Long


On Error GoTo Erro_SpedFiscal_E520
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204063
    Next


    objSpedFiscal.sTipo = "E520"
    objSpedRegistroE520.sTipo = objSpedFiscal.sTipo

    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroE520.iCount = objSpedRegistroE520.iCount + 1
    lSeqPai = objSpedFiscal.lSeq
    
    '*************** registro E530 **********************
    dTotAjDebitos = 0
    dTotAjCreditos = 0
    
    lErro = SpedFiscal_E530(objSpedFiscal, objSpedRegistroE530, objFilialEmpresa, dtDataIni, dtDataFim, lSeqPai, dTotAjDebitos, dTotAjCreditos)
    If lErro <> SUCESSO Then gError 204064
    
    'Verifica se o Item de apuração ICMS está cadastrada
    lErro = Comando_ExecutarPos(alComando(2), "SELECT SaldoCredorInicial FROM RegApuracaoIPI WHERE  DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ? ", 0, _
    dSaldoCredorInicial, dtDataIni, dtDataFim, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 204065
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204066
    
    If lErro = AD_SQL_SEM_DADOS Then dSaldoCredorInicial = 0
    
    dSaldoApurado = Arredonda_Moeda(dValDebIPI) + Arredonda_Moeda(dTotAjDebitos) - Arredonda_Moeda(dValCredIPI) - Arredonda_Moeda(dTotAjCreditos) - Arredonda_Moeda(dSaldoCredorInicial)
    
    objSpedFiscal.sRegistro = "|E520|" & Format(dSaldoCredorInicial, "fixed") & "|" & Format(dValDebIPI, "fixed") & "|" & Format(dValCredIPI, "fixed") & "|" & Format(dTotAjDebitos, "fixed") & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dTotAjCreditos, "fixed") & "|" & IIf(dSaldoApurado <= 0, Format(Abs(dSaldoApurado), "fixed"), 0) & "|" & IIf(dSaldoApurado > 0, Format(dSaldoApurado, "fixed"), 0) & "|"
    
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204068

    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai) VALUES (?,?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, lSeqPai, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqE500)
    If lErro <> AD_SQL_SUCESSO Then gError 204069

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E520 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E520:

   SpedFiscal_E520 = gErr

   Select Case gErr
    
        Case 204063
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204064
    
        Case 204065, 204066
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPI", gErr)
    
        Case 204067
            Call Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOIPI_NAO_CADASTRADO", gErr)
    
        Case 204068
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204069
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204070)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_E990(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE990 As ClassSpedRegistro, ByVal objSpedRegistroE001 As ClassSpedRegistro, ByVal objSpedRegistroE100 As ClassSpedRegistro, ByVal objSpedRegistroE110 As ClassSpedRegistro, ByVal objSpedRegistroE111 As ClassSpedRegistro, ByVal objSpedRegistroE116 As ClassSpedRegistro, objSpedRegistroE200 As ClassSpedRegistro, objSpedRegistroE210 As ClassSpedRegistro, objSpedRegistroE220 As ClassSpedRegistro, objSpedRegistroE250 As ClassSpedRegistro, ByVal objSpedRegistroE500 As ClassSpedRegistro, ByVal objSpedRegistroE510 As ClassSpedRegistro, ByVal objSpedRegistroE520 As ClassSpedRegistro, ByVal objSpedRegistroE530 As ClassSpedRegistro, ByVal objSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lTotal As Long


On Error GoTo Erro_SpedFiscal_E990
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204071
    Next


    objSpedFiscal.sTipo = "E990"
    objSpedRegistroE990.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroE990.iCount = objSpedRegistroE990.iCount + 1

    lTotal = objSpedRegistroE001.iCount + objSpedRegistroE100.iCount + objSpedRegistroE110.iCount + objSpedRegistroE111.iCount + objSpedRegistroE116.iCount + objSpedRegistroE200.iCount + objSpedRegistroE210.iCount + objSpedRegistroE220.iCount + objSpedRegistroE250.iCount + objSpedRegistroE500.iCount + objSpedRegistroE510.iCount + objSpedRegistroE520.iCount + objSpedRegistroE530.iCount + objSpedRegistroE990.iCount + objSpedAux.objSpedRegistroE300.iCount + objSpedAux.objSpedRegistroE310.iCount + objSpedAux.objSpedRegistroE311.iCount + objSpedAux.objSpedRegistroE312.iCount + objSpedAux.objSpedRegistroE313.iCount + objSpedAux.objSpedRegistroE316.iCount
    
    objSpedFiscal.sRegistro = "|E990|" & lTotal & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204072
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204073

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E990 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E990:

   SpedFiscal_E990 = gErr

   Select Case gErr
    
        Case 204071
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204072
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204073
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204074)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Private Function SpedFiscal_H001(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroH001 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal iIncluiRegInv As Integer) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim iIndMov As Integer


On Error GoTo Erro_SpedFiscal_H001
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204075
    Next

    'If Month(dtDataIni) = 2 Then
    If iIncluiRegInv = MARCADO Then
        iIndMov = 0
    Else
        iIndMov = 1
    End If
    
    objSpedFiscal.sTipo = "H001"
    objSpedRegistroH001.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroH001.iCount = objSpedRegistroH001.iCount + 1

    objSpedFiscal.sRegistro = "|H001|" & iIndMov & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204076
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204077

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_H001 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_H001:

   SpedFiscal_H001 = gErr

   Select Case gErr
    
        Case 204075
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204076
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204077
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204078)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_H005(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroH005 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dValorTotalEst As Double, ByVal iMotivoRegInv As Integer, ByVal dtDataInv As Date) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim iAno As Integer, sData As String, dtDataAux As Date

On Error GoTo Erro_SpedFiscal_H005
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204079
    Next

    objSpedFiscal.sTipo = "H005"
    objSpedRegistroH005.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroH005.iCount = objSpedRegistroH005.iCount + 1
    
'    O inventário deverá ser apresentado no arquivo da EFD-ICMS/IPI, no segundo mês subsequente ao evento. Ex.
'    inventário realizado em 31/12/08 deverá ser apresentado na EFD-ICMS/IPI de período de referência fevereiro de 2009.
    
    If dtDataInv = DATA_NULA Then
        dtDataAux = DateAdd("m", -1, dtDataIni)  'Retorna mais um mês
        dtDataAux = DateAdd("d", -Day(dtDataAux), dtDataAux) 'Retorna para o último dia do mês anterior
    Else
        dtDataAux = dtDataInv
    End If
    sData = Format(dtDataAux, "ddmmyyyy")

    iAno = Year(dtDataIni) - 1

    If giVersao >= 6 Then
        'objSpedFiscal.sRegistro = "|H005|" & "3112" & iAno & "|" & Format(dValorTotalEst, "fixed") & "|" & Format(iMotivoRegInv, "00") & "|"
        objSpedFiscal.sRegistro = "|H005|" & sData & "|" & Format(dValorTotalEst, "fixed") & "|" & Format(iMotivoRegInv, "00") & "|"
    Else
        'objSpedFiscal.sRegistro = "|H005|" & "3112" & iAno & "|" & Format(dValorTotalEst, "fixed") & "|"
        objSpedFiscal.sRegistro = "|H005|" & sData & "|" & Format(dValorTotalEst, "fixed") & "|"
    End If
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204080
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204081


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_H005 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_H005:

   SpedFiscal_H005 = gErr

   Select Case gErr
    
        Case 204079
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204080
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204081
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204082)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_H990(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroH990 As ClassSpedRegistro, ByVal objSpedRegistroH001 As ClassSpedRegistro, ByVal objSpedRegistroH005 As ClassSpedRegistro, ByVal objSpedRegistroH010 As ClassSpedRegistro) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lTotal As Long


On Error GoTo Erro_SpedFiscal_H990
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204083
    Next


    objSpedFiscal.sTipo = "H990"
    objSpedRegistroH990.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroH990.iCount = objSpedRegistroH990.iCount + 1


    lTotal = objSpedRegistroH001.iCount + objSpedRegistroH005.iCount + objSpedRegistroH010.iCount + objSpedRegistroH990.iCount

    objSpedFiscal.sRegistro = "|H990|" & lTotal & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204084
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204085


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_H990 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_H990:

   SpedFiscal_H990 = gErr

   Select Case gErr
    
        Case 204083
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204084
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204085
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204086)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_1001(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro1001 As ClassSpedRegistro) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lTotal As Long


On Error GoTo Erro_SpedFiscal_1001
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204087
    Next

    objSpedFiscal.sTipo = "1001"
    objSpedRegistro1001.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistro1001.iCount = objSpedRegistro1001.iCount + 1

    If giVersao >= 6 Then
        objSpedFiscal.sRegistro = "|1001|0|"
    Else
        objSpedFiscal.sRegistro = "|1001|1|"
    End If

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204088
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204089


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_1001 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_1001:

   SpedFiscal_1001 = gErr

   Select Case gErr
    
        Case 204087
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204088
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204089
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204090)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_1010(objSpedFiscal As ClassSpedFiscal, ByVal iFilialEmpresa As Integer, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal objSpedRegistro1010 As ClassSpedRegistro) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim lTotal As Long, sTemExportacao As String, sTemVA As String
Dim lNumIntDocAux As Long

On Error GoTo Erro_SpedFiscal_1010
        
    If giVersao >= 6 Then
        
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 210772
        Next
    
        objSpedFiscal.sTipo = "1010"
        objSpedRegistro1010.sTipo = objSpedFiscal.sTipo
               
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistro1010.iCount = objSpedRegistro1010.iCount + 1
        
        sTemExportacao = "N"

        lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc FROM DEInfo WHERE FilialEmpresa = ? AND DataAverbacao BETWEEN ? AND ?", _
        lNumIntDocAux, iFilialEmpresa, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 216046
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216047
        
        If lErro = AD_SQL_SUCESSO Then sTemExportacao = "S"
    
        sTemVA = "N"
    
        lErro = Comando_Executar(alComando(3), "SELECT NumNotaFiscal FROM DFC_Prod WHERE FilialEmpresa = ? AND Data BETWEEN ? AND ?", _
        lNumIntDocAux, iFilialEmpresa, dtDataIni, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 216046
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216047
        
        If lErro = AD_SQL_SUCESSO Then sTemVA = "S"
    
        objSpedFiscal.sRegistro = "|1010|" & sTemExportacao & "|N|N|N|" & sTemVA & "|N|N|N|N|"

        If giVersao >= 14 Then objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "N|"
        
        If giVersao > 12 Then
            'Principais alterações no Guia Prático da EFD-ICMS/IPI – versão 3.01
            'Inclusão no registro 1010 dos campos 11, 12 e 13 (IND_GIAFn), que poderá ser preenchido com “Sim”, se
            'estabelecimento informante do arquivo for domiciliado no estado de Pernambuco. Demais contribuintes deverão informar
            '“Não”
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro + "N|N|N|"
        End If
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 210773
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 210774
    
        For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If

    SpedFiscal_1010 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_1010:

   SpedFiscal_1010 = gErr

   Select Case gErr
    
        Case 210772
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 210773
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 210774
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
            
        Case 216045, 216047
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "DEInfo")
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210775)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_1100(ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim objSpedFiscal As ClassSpedFiscal
Dim alComando(1 To 4) As Long
Dim iIndice As Integer, sFiltro As String
Dim tDE As typeDEInfo, tNF As typeNFiscal, sProd As String
Dim sNumRE As String, dtDataRE As Date
Dim iModDocFis As Integer, iModDocFisE As Integer
Dim sChvNFe As String, iModelo As Integer
Dim objModDocFis As ClassModelosDocFiscais

On Error GoTo Erro_SpedFiscal_1100
        
    If giVersao >= 6 Then
        
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 216047
        Next
        
        'Alocação de espaço no buffer
        tDE.sNumero = String(STRING_MAXIMO, 0)
        tDE.sDescricao = String(STRING_MAXIMO, 0)
        tDE.sNumConhEmbarque = String(STRING_MAXIMO, 0)
        tDE.sUFEmbarque = String(STRING_MAXIMO, 0)
        tDE.sLocalEmbarque = String(STRING_MAXIMO, 0)
        sNumRE = String(STRING_MAXIMO, 0)
           
        lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Numero, Data, FilialEmpresa, Descricao, TipoDoc, " & _
                    "Natureza, NumConhEmbarque, UFEmbarque, LocalEmbarque, DataConhEmbarque, TipoConhEmbarque, CodPais, " & _
                    "DataAverbacao, NumRegistro, DataRegistro FROM DERE WHERE FilialEmpresa = ? AND DataAverbacao BETWEEN ? AND ? ORDER BY Data, Numero, NumRegistro", _
                    tDE.lNumIntDoc, tDE.sNumero, tDE.dtData, tDE.iFilialEmpresa, _
                    tDE.sDescricao, tDE.iTipoDoc, tDE.iNatureza, tDE.sNumConhEmbarque, tDE.sUFEmbarque, _
                    tDE.sLocalEmbarque, tDE.dtDataConhEmbarque, tDE.iTipoConhEmbarque, tDE.iCodPais, tDE.dtDataAverbacao, _
                    sNumRE, dtDataRE, objFisSpedAux.iFilialEmpresa, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 216048
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216049
        
        Do While lErro = AD_SQL_SUCESSO
        
            Set objSpedFiscal = objFisSpedAux.objSpedFiscal
        
            objSpedFiscal.sTipo = "1100"
            objFisSpedAux.objSpedRegistro1100.sTipo = objSpedFiscal.sTipo
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            objFisSpedAux.objSpedRegistro1100.iCount = objFisSpedAux.objSpedRegistro1100.iCount + 1
    
            objSpedFiscal.sRegistro = "|1100|" '1-REG
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & CStr(tDE.iTipoDoc) & "|" '2-IND_DOC
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tDE.sNumero & "|" '3-NRO_DE
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tDE.dtData, "ddmmyyyy") & "|" '4-DT_DE
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & CStr(tDE.iNatureza) & "|" '5-NAT_EXP
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & sNumRE & "|" '6-NRO_RE
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dtDataRE, "ddmmyyyy") & "|" '7-DT_RE
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & tDE.sNumConhEmbarque & "|" '8-CHC_EMB
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tDE.dtDataConhEmbarque, "ddmmyyyy") & "|" '9-DT_CHC
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tDE.dtDataAverbacao, "ddmmyyyy") & "|" '10-DT_AVB
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tDE.iTipoConhEmbarque, "00") & "|" '11-TP_CHC
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & CStr(tDE.iCodPais) & "|" '12-PAIS
        
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 216050
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
            If lErro <> AD_SQL_SUCESSO Then gError 216051
            
            sFiltro = "AND (I.NumIntDoc IN " & _
                "(SELECT NumIntDocItem FROM InfoAdicDocItem WHERE TipoDoc = 9 AND NumIntDE = ? AND NumRE = ?) " & _
                    "OR I.NumIntDoc IN " & _
                "(SELECT NumIntDocItem FROM InfoAdicDocItem AS X, ItensNFiscal AS I, InfoAdicExportacao AS Y " & _
                "WHERE X.TipoDoc = 9 AND X.NumIntDE = 0 AND I.NumIntDoc = X.NumIntDocItem AND " & _
                "Y.TipoDoc = 0 AND Y.NumIntDoc = I.NumIntNF AND Y.NumIntDE = ? AND Y.NumRE = ?))"
                
            tNF.sSerie = String(STRING_SERIE, 0)
            sProd = String(STRING_PRODUTO, 0)
            
            lErro = Comando_Executar(alComando(3), "SELECT N.NumIntDoc, N.Serie, N.NumNotaFiscal, N.DataEmissao, I.Produto, N.NFe, T.ModDocFis, T.ModDocFisE  FROM NFiscal As N, ItensNFiscal As I, TiposDocInfo As T WHERE N.NumIntDoc = I.NumIntNF AND N.Status <> 7 AND T.Codigo = N.TipoNFiscal " & sFiltro, _
            tNF.lNumIntDoc, tNF.sSerie, tNF.lNumNotaFiscal, tNF.dtDataEmissao, sProd, tNF.iNFe, iModDocFis, iModDocFisE, tDE.lNumIntDoc, sNumRE, tDE.lNumIntDoc, sNumRE)
            If lErro <> AD_SQL_SUCESSO Then gError 216052
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216053
            
            Do While lErro = AD_SQL_SUCESSO
            
                lErro = SpedFiscal_0200(sProd, objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, 0, objFisSpedAux)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                objSpedFiscal.sTipo = "1105"
                objFisSpedAux.objSpedRegistro1105.sTipo = objSpedFiscal.sTipo
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objFisSpedAux.objSpedRegistro1105.iCount = objFisSpedAux.objSpedRegistro1105.iCount + 1
                
                If tNF.iNFe = MARCADO Then
                    iModelo = iModDocFisE
                Else
                    iModelo = iModDocFis
                End If
                                
                Set objModDocFis = New ClassModelosDocFiscais
                
                objModDocFis.iCodigo = iModelo
                
                lErro = CF("ModelosDocFiscais_Le", objModDocFis)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
                                
                sChvNFe = String(STRING_NFE_CHNFE, 0)
                
                'se for uma nota eletronica autorizada no ambiente de producao
                lErro = Comando_Executar(alComando(4), "SELECT chNFe FROM NFEFedProtNFe WHERE NumIntNF = ? AND cStat IN ('100', '101', '205', '110', '301', '302') AND tpAmb = 0 ", sChvNFe, tNF.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 216054
            
                lErro = Comando_BuscarPrimeiro(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216055
                
                If lErro <> AD_SQL_SUCESSO Then sChvNFe = ""
        
                objSpedFiscal.sRegistro = "|1105|" '1-REG
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & objModDocFis.sModelo & "|" '2-COD_MOD
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Ajusta_Serie(Desconverte_Serie_Eletronica(tNF.sSerie)) & "|"  '3-SERIE
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & CStr(tNF.lNumNotaFiscal) & "|" '4-NUM_DOC
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Trim(sChvNFe) & "|" '5-CHV_NFE
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tNF.dtDataEmissao, "ddmmyyyy") & "|" '6-DT_DOC
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(sProd) & "|" '7-COD_ITEM
            
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 216056
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
                If lErro <> AD_SQL_SUCESSO Then gError 216057
            
                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216058
                
            Loop
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216059
            
        Loop
    
        For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If

    SpedFiscal_1100 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_1100:

   SpedFiscal_1100 = gErr

   Select Case gErr
    
        Case 216047
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 216048, 216049, 216059
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "DERE")
    
        Case 216050, 216056
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 216051, 216057
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
            
        Case 216052, 216053, 216058
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "NFiscal")
        
        Case 216054, 216055
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "NFEFedProtNFe")
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216060)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_1400(ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim objSpedFiscal As ClassSpedFiscal
Dim alComando(1 To 2) As Long
Dim iIndice As Integer, sCodIBGE As String, sProduto As String, dValor As Double

On Error GoTo Erro_SpedFiscal_1400
        
    If giVersao >= 6 Then
        
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 216047
        Next
        
        'Alocação de espaço no buffer
        sCodIBGE = String(STRING_MAXIMO, 0)
        sProduto = String(STRING_MAXIMO, 0)
           
        lErro = Comando_Executar(alComando(2), "SELECT CodIBGE, Produto, SUM(ValorProduto) FROM [DFC_Prod] WHERE FilialEmpresa = ? AND Data BETWEEN ? AND ? GROUP BY CodIBGE, Produto ORDER BY CodIBGE, Produto", _
                    sCodIBGE, sProduto, dValor, objFisSpedAux.iFilialEmpresa, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 216048
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216049
        
        Do While lErro = AD_SQL_SUCESSO
        
            Set objSpedFiscal = objFisSpedAux.objSpedFiscal
        
            lErro = SpedFiscal_0200(sProduto, objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, 0, objFisSpedAux)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            objSpedFiscal.sTipo = "1400"
            objFisSpedAux.objSpedRegistro1400.sTipo = objSpedFiscal.sTipo
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            objFisSpedAux.objSpedRegistro1400.iCount = objFisSpedAux.objSpedRegistro1400.iCount + 1
    
            objSpedFiscal.sRegistro = "|1400|" '1-REG
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(sProduto) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & sCodIBGE & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dValor, "Fixed") & "|"
        
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 216050
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
            If lErro <> AD_SQL_SUCESSO Then gError 216051
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216059
            
        Loop
    
        For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If

    SpedFiscal_1400 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_1400:

   SpedFiscal_1400 = gErr

   Select Case gErr
    
        Case 216047
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 216048, 216049, 216059
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "DFC_Prod")
    
        Case 216050, 216056
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 216051, 216057
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216060)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_1990(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro1990 As ClassSpedRegistro, ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lTotal As Long


On Error GoTo Erro_SpedFiscal_1990
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204091
    Next

    objSpedFiscal.sTipo = "1990"
    objSpedRegistro1990.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistro1990.iCount = objSpedRegistro1990.iCount + 1
    
    lTotal = objFisSpedAux.objSpedRegistro1001.iCount + objFisSpedAux.objSpedRegistro1010.iCount + objFisSpedAux.objSpedRegistro1100.iCount + objFisSpedAux.objSpedRegistro1105.iCount + objFisSpedAux.objSpedRegistro1400.iCount + objSpedRegistro1990.iCount

'    If giVersao >= 6 Then
'        objSpedFiscal.sRegistro = "|1990|3|"
'    Else
'        objSpedFiscal.sRegistro = "|1990|2|"
'    End If
    objSpedFiscal.sRegistro = "|1990|" & CStr(lTotal) & "|"
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204092
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204093

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_1990 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_1990:

   SpedFiscal_1990 = gErr

   Select Case gErr
    
        Case 204091
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204092
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204093
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204094)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_9001(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro9001 As ClassSpedRegistro) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer


On Error GoTo Erro_SpedFiscal_9001
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204095
    Next

    objSpedFiscal.sTipo = "9001"
    objSpedRegistro9001.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistro9001.iCount = objSpedRegistro9001.iCount + 1

    objSpedFiscal.sRegistro = "|9001|0|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204096
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204097

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_9001 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_9001:

   SpedFiscal_9001 = gErr

   Select Case gErr
    
        Case 204095
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204096
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204097
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204098)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_9990(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistro9990 As ClassSpedRegistro, ByVal colSpedRegistro As Collection) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer


On Error GoTo Erro_SpedFiscal_9990
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204099
    Next


    objSpedFiscal.sTipo = "9990"
    objSpedRegistro9990.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1

    objSpedFiscal.sRegistro = "|9990|" & colSpedRegistro.Count + 3 & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204100
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204101


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_9990 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_9990:

   SpedFiscal_9990 = gErr

   Select Case gErr
    
        Case 204099
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204100
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204101
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204102)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Function SpedFiscal_9999(objSpedFiscal As ClassSpedFiscal) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer


On Error GoTo Erro_SpedFiscal_9999
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204103
    Next


    objSpedFiscal.sTipo = "9999"
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1

    objSpedFiscal.sRegistro = "|9999|" & objSpedFiscal.lSeq & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204104
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204105


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_9999 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_9999:

   SpedFiscal_9999 = gErr

   Select Case gErr
    
        Case 204103
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204104
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204105
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204106)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_C001(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC001 As ClassSpedRegistro, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 5) As Long
Dim iIndice As Integer
Dim lContador As Long
Dim iTemReg As Integer
Dim iCodigoECF As Integer



On Error GoTo Erro_SpedFiscal_C001
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204137
    Next


    'se for uma nota fiscal de entrada ou saida
    lErro = Comando_Executar(alComando(2), "SELECT Count(*) From NFiscal AS N, TiposDocInfo AS T WHERE TipoNFiscal = Codigo AND ((Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) Or ((Tipo=1 Or Tipo=3) AND DataEntrada >= ? AND DataEntrada <= ?))  AND FilialEmpresa = ? ", _
    lContador, dtDataIni, dtDataFim, dtDataIni, dtDataFim, objSpedFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 204129

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204130

    If lContador = 0 Then


        'se for uma nota fiscal de entrada de energia ou gas
        lErro = Comando_Executar(alComando(3), "SELECT Count(*) From NFiscal AS N, TiposDocInfo AS T WHERE TipoNFiscal = Codigo AND (Tipo = 1 Or Tipo =3) AND DataEntrada >= ? AND DataEntrada <= ? AND (ModeloArqICMS = 06 OR ModeloArqICMS = 28) AND FilialEmpresa = ?", _
        lContador, dtDataIni, dtDataFim, objSpedFiscal.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 203529
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203530


        If lContador = 0 Then

            lErro = Comando_Executar(alComando(4), "SELECT E.Codigo  FROM ECF AS E, ImpressoraECF AS I, ModeloECF AS M WHERE E.FilialEmpresa = ? AND E.ImpressoraECF = I.Codigo AND E.FilialEmpresa = I.FilialEmpresa AND I.CodModelo = M.Codigo", _
            iCodigoECF, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 203330
        
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203331
            
            Do While lErro = AD_SQL_SUCESSO
            
                lErro = Comando_Executar(alComando(5), "SELECT Count(*) FROM LojaArqFisMestre WHERE CodECF = ? And FilialEmpresa = ? And Data >= ? AND Data <= ?", _
                lContador, iCodigoECF, iFilialEmpresa, dtDataIni, dtDataFim)
                If lErro <> AD_SQL_SUCESSO Then gError 203332
            
                lErro = Comando_BuscarPrimeiro(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203333
        
                If lContador > 0 Then Exit Do
        
                lErro = Comando_BuscarProximo(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203331
        
        
            Loop

        End If

    End If

    objSpedFiscal.sTipo = "C001"
    objSpedRegistroC001.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroC001.iCount = objSpedRegistroC001.iCount + 1

    If lContador > 0 Then
        iTemReg = 0
    Else
        iTemReg = 1
    End If

    objSpedFiscal.sRegistro = "|C001|" & iTemReg & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204138
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 204139

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C001 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C001:

   SpedFiscal_C001 = gErr

   Select Case gErr
    
        Case 204137
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204138
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204139
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204140)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function Checa_Duplicidade_0175(ByVal objSpedFiscal As ClassSpedFiscal, ByVal iNumero As Integer, ByVal dtData As Date, iDuplicado As Integer, ByVal lSeqPai As Long) As Long
        
Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lSeq As Long

On Error GoTo Erro_Checa_Duplicidade_0175
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204150
    Next
        
    lErro = Comando_Executar(alComando(1), "SELECT Seq FROM SpedFiscal WHERE NumIntSped = ? AND FilialEmpresa = ? AND Tipo = '0175' AND SUBSTRING(Registro, 7,8) = ? AND SUBSTRING(Registro, 16,2) = ? AND SeqPai = ?", _
    lSeq, objSpedFiscal.lNumIntSped, objSpedFiscal.iFilialEmpresa, Format(dtData, "ddmmyyyy"), Format(iNumero, "00"), lSeqPai)
    If lErro <> AD_SQL_SUCESSO Then gError 204151

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204152

    If lErro = AD_SQL_SUCESSO Then
        iDuplicado = 1
    Else
        iDuplicado = 0
    End If
        
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Checa_Duplicidade_0175 = SUCESSO

    Exit Function
    
Erro_Checa_Duplicidade_0175:

   Checa_Duplicidade_0175 = gErr

   Select Case gErr
    
        Case 204150
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204151, 204152
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISCAL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204153)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


'Function Rateia_Compl_ItemNF(tNFiscal As typeNFiscal, tItemNF As typeItemNF, tTribItemNFParam As typeTribItemNF, dTotalRateioNFICMSBase As Double, dTotalRateioNFICMSValor As Double, dTotalRateioNFICMSSubstBase As Double, dTotalRateioNFICMSSubstValor As Double, dTotalRateioNFIPIBaseCalculo As Double, dTotalRateioNFIPIValor As Double) As Long
Private Function Rateia_Compl_ItemNF(tNFiscal As typeNFiscal, tItemNF As typeItemNF, tTribItemNFParam As typeTribItemNF) As Long

'''Dim lErro As Long
'''Dim alComando(1 To 1) As Long
'''Dim tTribItemNF As typeTribItemNF
'''Dim iIndice As Integer
'''Dim iTipoRegistro As Integer
'''Dim iIPICSTEntrada As Integer
'''Dim iIPICSTSaida As Integer
'''Dim iTipoTribCST_ICMS As Integer
'''
'''On Error GoTo Erro_Rateia_Compl_ItemNF
'''
'''    'Abertura de Comando
'''    For iIndice = LBound(alComando) To UBound(alComando)
'''        alComando(iIndice) = Comando_Abrir()
'''        If alComando(iIndice) = 0 Then gError 204150
'''    Next
'''
'''    'trata os complementos das notas fiscais
'''
'''    tTribItemNF.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'''
'''    lErro = Comando_Executar(alComando(1), "SELECT T.Tipo, TipoTribCST, NaturezaOP, ICMSBase, ICMSAliquota, ICMSValor, ICMSSubstBase, ICMSSubstValor, IPI.CSTEntrada, IPI.CSTSaida, " & _
'''    "IPIValor, ICMSPercRedBase FROM TributacaoComplNF AS T, tiposTribICMS AS ICMS, TiposTribIPI AS IPI  WHERE T.NumIntDoc  = ? AND ICMS.Tipo = ICMSTipo AND IPI.Tipo = IPITipo AND (T.Tipo = 1 Or T.Tipo=2 Or T.Tipo = 4)", _
'''    iTipoRegistro, iTipoTribCST_ICMS, tTribItemNF.sNaturezaOp, tTribItemNF.dICMSBase, tTribItemNF.dICMSAliquota, tTribItemNF.dICMSValor, tTribItemNF.dICMSSubstBase, tTribItemNF.dICMSSubstValor, _
'''    iIPICSTEntrada, iIPICSTSaida, tTribItemNF.dIPIValor, tTribItemNF.dICMSPercRedBase, tNFiscal.lNumIntDoc)
'''    If lErro <> AD_SQL_SUCESSO Then gError 204190
'''
'''    lErro = Comando_BuscarPrimeiro(alComando(1))
'''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204191
'''
'''    Do While lErro = AD_SQL_SUCESSO
'''
'''
'''        tTribItemNFParam.dICMSBase = tTribItemNFParam.dICMSBase + Round((tTribItemNF.dICMSBase * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'''        tTribItemNFParam.dICMSValor = tTribItemNFParam.dICMSValor + Round((tTribItemNF.dICMSValor * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'''        tTribItemNFParam.dICMSSubstBase = tTribItemNFParam.dICMSSubstBase + Round((tTribItemNF.dICMSSubstBase * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'''        tTribItemNFParam.dICMSSubstValor = tTribItemNFParam.dICMSSubstValor + Round((tTribItemNF.dICMSSubstValor * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'''        tTribItemNFParam.dIPIBaseCalculo = tTribItemNFParam.dIPIBaseCalculo + Round((tTribItemNF.dIPIBaseCalculo * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'''        tTribItemNFParam.dIPIValor = tTribItemNFParam.dIPIValor + Round((tTribItemNF.dIPIValor * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'''
''''        dTotalRateioNFICMSBase = dTotalRateioNFICMSBase + Round((tTribItemNF.dICMSBase * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
''''        dTotalRateioNFICMSValor = dTotalRateioNFICMSValor + Round((tTribItemNF.dICMSValor * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
''''        dTotalRateioNFICMSSubstBase = dTotalRateioNFICMSSubstsBase + Round((tTribItemNF.dICMSSubstBase * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
''''        dTotalRateioNFICMSSubstValor = dTotalRateioNFICMSSubstsValor + Round((tTribItemNF.dICMSSubstValor * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
''''        dTotalRateioNFIPIBaseCalculo = dTotalRateioNFIPIBaseCalculo + Round((tTribItemNF.dIPIBaseCalculo * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
''''        dTotalRateioNFIPIValor = dTotalRateioNFIPIValor + Round((tTribItemNF.dIPIValor * ((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) / tNFiscal.dValorProdutos)), 2)
'''
'''        lErro = Comando_BuscarProximo(alComando(1))
'''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204191
'''
'''    Loop
'''
'''
'''    For iIndice = LBound(alComando) To UBound(alComando)
'''         Call Comando_Fechar(alComando(iIndice))
'''    Next

    Rateia_Compl_ItemNF = SUCESSO

'''    Exit Function
'''
'''Erro_Rateia_Compl_ItemNF:
'''
'''   Rateia_Compl_ItemNF = gErr
'''
'''   Select Case gErr
'''
'''        Case 204202
'''            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'''
'''        Case 204203, 204204
'''            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLNF", gErr)
'''
'''        Case Else
'''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204205)
'''
'''    End Select
'''
'''    For iIndice = LBound(alComando) To UBound(alComando)
'''         Call Comando_Fechar(alComando(iIndice))
'''    Next
'''

End Function

Private Function SpedFiscal_Trata_ItemNF(objTela As Object, objFilialEmpresa As AdmFiliais, dtDataIni As Date, dtDataFim As Date, objSpedFiscal As ClassSpedFiscal, tNFiscal As typeNFiscal, objSpedRegistro0150 As ClassSpedRegistro, objSpedRegistro0175 As ClassSpedRegistro, objSpedRegistro0190 As ClassSpedRegistro, objSpedRegistro0200 As ClassSpedRegistro, objSpedRegistro0205 As ClassSpedRegistro, objSpedRegistro0220 As ClassSpedRegistro, tTipoDocInfo As typeTipoDocInfo, alComando() As Long, objSpedRegistroC100 As ClassSpedRegistro, objSpedRegistro0450 As ClassSpedRegistro, objSpedRegistroC110 As ClassSpedRegistro, _
objSpedRegistroC111 As ClassSpedRegistro, objSpedRegistroC112 As ClassSpedRegistro, objSpedRegistroC113 As ClassSpedRegistro, objSpedRegistroC114 As ClassSpedRegistro, objSpedRegistroC115 As ClassSpedRegistro, objSpedRegistroC120 As ClassSpedRegistro, objSpedRegistroC130 As ClassSpedRegistro, objSpedRegistroC140 As ClassSpedRegistro, objSpedRegistroC141 As ClassSpedRegistro, objSpedRegistroC160 As ClassSpedRegistro, _
objSpedRegistroC170 As ClassSpedRegistro, objSpedRegistroC172 As ClassSpedRegistro, objSpedRegistroC190 As ClassSpedRegistro, colSpedE510 As Collection, colSpedC190 As Collection, dValDebICMS_objSpedC190 As Double, dValCredICMS_objSpedC190 As Double, colMensagem As Collection, colUM As Collection, colSpedC190UF As Collection, ByVal objFisSpedAux As ClassFISSpedAux) As Long


Dim lErro As Long
Dim schNFe As String
Dim iNFe As Integer
Dim sCodParticipante As String
Dim iCodSit As Integer
Dim iAmbiente As Integer
Dim iCondPagto As Integer
Dim iIndPag As Integer
Dim iFreteRespons As Integer
Dim tTributacaoNF As typeTributacaoNF
Dim lSeqPaiNF As Long
Dim dValorMercadorias As Double, dValorDesconto As Double
Dim sData As String
Dim lSeqPai As Long
Dim objMensagem As ClassMensagem
Dim dTotalRateioNF As Double
Dim vbMsgRes As VbMsgBoxResult
Dim objTrib As ClassTributacaoDoc

On Error GoTo Erro_SpedFiscal_Trata_ItemNF

    schNFe = String(STRING_NFE_CHNFE, 0)

    'se for uma nota eletronica autorizada no ambiente de producao
    lErro = Comando_Executar(alComando(5), "SELECT chNFe FROM NFEFedProtNFe WHERE NumIntNF = ? AND cStat IN ('100', '101', '205', '110', '301', '302') AND tpAmb = 0 ", schNFe, tNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 203303

    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203304

    If lErro <> AD_SQL_SUCESSO Then
        
        'se for uma nota eletronica autorizada no ambiente de producao
        lErro = Comando_Executar(alComando(34), "SELECT chvNFe FROM ImportNFeXml WHERE NumIntNF = ? ", schNFe, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203303
    
        lErro = Comando_BuscarPrimeiro(alComando(34))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203304

        If lErro <> AD_SQL_SUCESSO Then
        
            schNFe = tNFiscal.sChvNFe
            
        End If
        
    End If
    
    iNFe = SerieEletronica(tNFiscal.sSerie)

    'se for nota fiscal (codigo 01) ou nota do produto (codigo 04) ou NFE o u NFCE
    If tTipoDocInfo.iModeloArqICMS = 1 Or tTipoDocInfo.iModeloArqICMS = 4 Or tTipoDocInfo.iModeloArqICMS = 65 Then 'Or iNFe = 1 Then 'Não pode incluir o modelo 8 por exemplo só por ser eletrônico (vai ser um CT-e que não é modelo 55)

        lErro = Retorna_UF_Subst_Tributaria(objFilialEmpresa, tNFiscal, tTipoDocInfo)
        If lErro <> SUCESSO Then gError 204920
        
        '************ registro C100 ********************************
        
        lErro = NF_Obter_Situacao(alComando(36), tNFiscal, tTipoDocInfo, iCodSit)
        If lErro <> SUCESSO Then gError 204920

'        If tTipoDocInfo.iComplementar = 1 Then
'            iCodSit = 6
'        Else
'            iCodSit = 0
'        End If
'
'        If tNFiscal.iStatus = STATUS_CANCELADO Then iCodSit = 2
'
'        If iNFe = 1 And tTipoDocInfo.iTipo <> 3 Then
'
'            'se for uma nota eletronica cancelada no ambiente de producao
'            lErro = Comando_Executar(alComando(9), "SELECT tpAmb FROM NFEFedRetCancNFe WHERE NumIntNF = ? AND cStat = '101' AND tpAmb = 0", iAmbiente, tNFiscal.lNumIntDoc)
'            If lErro <> AD_SQL_SUCESSO Then gError 203307
'
'            lErro = Comando_BuscarPrimeiro(alComando(9))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203308
'
'            'nota fiscal cancelada
'            If lErro = AD_SQL_SUCESSO Then iCodSit = 4
'
'        End If

        sData = ""

        'se for nota de saida
        If tTipoDocInfo.iTipo = 2 Then
            If tNFiscal.dtDataSaida <= dtDataFim Then sData = Format(tNFiscal.dtDataSaida, "ddmmyyyy")
        Else
            sData = Format(tNFiscal.dtDataEntrada, "ddmmyyyy")
        End If
                        
        Select Case tNFiscal.iClasseDocCPR
        
            'se for um titulo a pagar
            Case CPR_TITULO_PAGAR

                iCondPagto = 0
                
                'pesquisa a condicao de pagto
                lErro = Comando_Executar(alComando(10), "SELECT CondicaoPagto FROM TitulosPagTodos WHERE NumIntDoc = ?", iCondPagto, tNFiscal.lNumIntDocCPR)
                If lErro <> AD_SQL_SUCESSO Then gError 203309
            
                lErro = Comando_BuscarPrimeiro(alComando(10))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203310
    
                If iCondPagto = 1 Then
                    iIndPag = 0
                Else
                    iIndPag = 1
                End If
        
            Case CPR_TITULO_RECEBER
            
                iCondPagto = 0
                
                'pesquisa a condicao de pagto
                lErro = Comando_Executar(alComando(11), "SELECT CondicaoPagto FROM TitulosRecTodos WHERE NumIntDoc = ?", iCondPagto, tNFiscal.lNumIntDocCPR)
                If lErro <> AD_SQL_SUCESSO Then gError 203311
            
                lErro = Comando_BuscarPrimeiro(alComando(11))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203312
    
                If iCondPagto = 1 Then
                    iIndPag = 0
                Else
                    iIndPag = 1
                End If
        
            Case CPR_NF_PAGAR
                iIndPag = 1
            
            Case Else
                If tTipoDocInfo.iFaturavel <> 0 Then
                    iIndPag = 1
                Else
                    If giVersao >= 6 Then
                        iIndPag = 2
                    Else
                        iIndPag = 9
                    End If
                End If
            
        End Select

'        If tNFiscal.dValorFrete <> 0 Then
'            If tNFiscal.iFreteRespons = 0 Or tNFiscal.iFreteRespons = 1 Then
'                iFreteRespons = 1
'            Else
'                iFreteRespons = 2
'            End If
'        Else
'            iFreteRespons = 9
'        End If
        
        Select Case tNFiscal.iFreteRespons
        
            Case 0, 1 'Emitente
                If tNFiscal.dtDataEmissao < #1/1/2012# Then
                    iFreteRespons = 1
                Else
                    iFreteRespons = 0
                End If
            Case 2 'Destinatário
                If tNFiscal.dtDataEmissao < #1/1/2012# Then
                    iFreteRespons = 2
                Else
                    iFreteRespons = 1
                End If
            Case 3 'Terceiros
                If tNFiscal.dtDataEmissao < #1/1/2012# Then
                    iFreteRespons = 0
                Else
                    iFreteRespons = 2
                End If
            Case 4 'Sem Frete
                iFreteRespons = 9

        End Select
        
        'Cód Desc CodNFe
        '1   Emitente 0
        '2   Destinatário 1
        '3   Terceiros 2
        '4   Sem Frete   9
        '5   Remetente (Transp.Próprio)  3
        '6   Destinatário (Transp.Próprio)   4
        If giVersao >= 12 Then
            Select Case tNFiscal.iFreteRespons
                Case 1
                    iFreteRespons = 0
                Case 2
                    iFreteRespons = 1
                Case 3
                    iFreteRespons = 2
                Case 4
                    iFreteRespons = 9
                Case 5
                    iFreteRespons = 3
                Case 6
                    iFreteRespons = 4
            End Select
        End If

        'pesquisa a tributacaoNF
        lErro = Comando_Executar(alComando(12), "SELECT ICMSBase, ICMSValor, ICMSSubstBase, ICMSSubstValor, IPIValor, IPICredito, PISCredito, COFINSCredito, PISRetido, COFINSRetido, ISSAliquota, ICMSCredito FROM TributacaoNF WHERE NumIntDoc = ?", _
        tTributacaoNF.dICMSBase, tTributacaoNF.dICMSValor, tTributacaoNF.dICMSSubstBase, tTributacaoNF.dICMSSubstValor, tTributacaoNF.dIPIValor, tTributacaoNF.dIPICredito, tTributacaoNF.dPISCredito, tTributacaoNF.dCOFINSCredito, tTributacaoNF.dPISRetido, tTributacaoNF.dCOFINSRetido, tTributacaoNF.dISSAliquota, tTributacaoNF.dICMSCredito, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203313
    
        lErro = Comando_BuscarPrimeiro(alComando(12))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203314
        
        'se a nota nao for cancelada
        If iCodSit <> 2 And iCodSit <> 4 And tTipoDocInfo.iModeloArqICMS <> 65 Then
        
            'se é uma nota de cliente
            If tNFiscal.lCliente > 0 Then
                lErro = SpedFiscal_0150_Cliente(tNFiscal.lCliente, tNFiscal.iFilialCli, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
                If lErro <> SUCESSO Then gError 203305
            Else
                lErro = SpedFiscal_0150_Fornecedor(tNFiscal.lFornecedor, tNFiscal.iFilialForn, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
                If lErro <> SUCESSO Then gError 203306
            End If

        End If

        objSpedFiscal.sTipo = "C100"
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        lSeqPaiNF = objSpedFiscal.lSeq
        
        objSpedRegistroC100.sTipo = objSpedFiscal.sTipo
        objSpedRegistroC100.iCount = objSpedRegistroC100.iCount + 1
        
        'se a nota nao for cancelada
        If iCodSit <> 2 And iCodSit <> 4 Then
                
            dValorMercadorias = 0
            dValorDesconto = 0
            
            Set objTrib = New ClassTributacaoDoc
            objTrib.lNumIntDoc = tNFiscal.lNumIntDoc
            objTrib.iTipoDoc = TIPODOC_TRIB_NF
            lErro = NF_Le_BasePropria(objTrib)
            If lErro <> SUCESSO Then gError 203306
            
            tTributacaoNF.dICMSBase = objTrib.dICMSBase
            tTributacaoNF.dICMSValor = objTrib.dICMSValor
            tTributacaoNF.dICMSCredito = objTrib.dICMSCredito
            
            tTributacaoNF.dICMSSubstBase = objTrib.dICMSSubstBase
            tTributacaoNF.dICMSSubstValor = objTrib.dICMSSubstValor

            tTributacaoNF.dIPIBase = objTrib.dIPIBase
            tTributacaoNF.dIPIValor = objTrib.dIPIValor
            tTributacaoNF.dIPICredito = objTrib.dIPICredito
            
            lErro = Comando_Executar(alComando(6), "SELECT SUM(Quantidade *PrecoUnitario), SUM(ValorDesconto)  FROM ItensNFiscal  WHERE NumIntNF  = ?", _
                dValorMercadorias, dValorDesconto, tNFiscal.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 203972
        
            lErro = Comando_BuscarPrimeiro(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203973
        
            objSpedFiscal.sRegistro = "|C100|" & IIf(tTipoDocInfo.iTipo = 2, 1, 0) & "|" & IIf(tTipoDocInfo.iTipo = 3, 1, 0) & "|" & sCodParticipante & "|" & IIf(iNFe = 1 And tTipoDocInfo.iModeloArqICMS = 1, "55", Format(tTipoDocInfo.iModeloArqICMS, "00")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(iCodSit, "00") & "|" & Ajusta_Serie(Desconverte_Serie_Eletronica(tNFiscal.sSerie)) & "|" & tNFiscal.lNumNotaFiscal & "|" & schNFe & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tNFiscal.dtDataEmissao, "ddmmyyyy") & "|" & sData & "|" & Format(tNFiscal.dValorTotal, "Fixed") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & iIndPag & "|" & IIf(dValorDesconto = 0, 0, Format(dValorDesconto, "Fixed")) & "||" & Format(dValorMercadorias, "Fixed") & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & iFreteRespons & "|" & IIf(tNFiscal.dValorFrete = 0, 0, Format(tNFiscal.dValorFrete, "Fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tNFiscal.dValorSeguro = 0, 0, Format(tNFiscal.dValorSeguro, "Fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tNFiscal.dValorOutrasDespesas = 0, 0, Format(tNFiscal.dValorOutrasDespesas, "Fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSBase = 0, 0, Format(tTributacaoNF.dICMSBase, "Fixed")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSCredito = 0, 0, Format(tTributacaoNF.dICMSCredito, "Fixed")) & "|"
            
            If tTipoDocInfo.iModeloArqICMS <> 65 Then
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSSubstBase = 0, 0, Format(tTributacaoNF.dICMSSubstBase, "Fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSSubstValor = 0, 0, Format(tTributacaoNF.dICMSSubstValor, "Fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dIPICredito = 0, 0, Format(tTributacaoNF.dIPICredito, "Fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dPISCredito = 0, 0, Format(tTributacaoNF.dPISCredito, "Fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dCOFINSCredito = 0, 0, Format(tTributacaoNF.dCOFINSCredito, "Fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dPISRetido = 0, 0, Format(tTributacaoNF.dPISRetido, "Fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dCOFINSRetido = 0, 0, Format(tTributacaoNF.dCOFINSRetido, "Fixed")) & "|"
            Else
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "|||||||"
            End If
        Else
        
            objSpedFiscal.sRegistro = "|C100|" & IIf(tTipoDocInfo.iTipo = 2, 1, 0) & "|" & IIf(tTipoDocInfo.iTipo = 3, 1, 0) & "||" & IIf(iNFe = 1, "55", Format(tTipoDocInfo.iModeloArqICMS, "00")) & "|"
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(iCodSit, "00") & "|" & Ajusta_Serie(Desconverte_Serie_Eletronica(tNFiscal.sSerie)) & "|" & tNFiscal.lNumNotaFiscal & "|" & schNFe & "|||||||||||||||||||||"
        
        End If
        
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203334
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 203315
        
'        'C Informação complementar dos documentos fiscais quando das operações interestaduais destinadas a consumidor final não contribuinte EC 87/15 (código 55) C101
'        lErro = SpedFiscal_C101(objFisSpedAux, tNFiscal.lNumIntDoc, tNFiscal.sUFRemet, tNFiscal.sUFDest, IIf(tTipoDocInfo.iTipo = 2, 1, 0), iCodSit, lSeqPaiNF)
'        If lErro <> SUCESSO Then gError 203358

        If iCodSit <> 2 And iCodSit <> 4 And tTipoDocInfo.iModeloArqICMS <> 65 Then
                
            'C Informação complementar dos documentos fiscais quando das operações interestaduais destinadas a consumidor final não contribuinte EC 87/15 (código 55) C101
            lErro = SpedFiscal_C101(objFisSpedAux, tNFiscal.lNumIntDoc, tNFiscal.sUFRemet, tNFiscal.sUFDest, IIf(tTipoDocInfo.iTipo = 2, 1, 0), iCodSit, lSeqPaiNF)
            If lErro <> SUCESSO Then gError 203358
                    
                
            If Len(tNFiscal.sMensagemNota) > 0 Then
        
            
                '************ registro C110 ********************************
                
                'If iNFe <> 1 Then
                
                    Set objMensagem = New ClassMensagem
            
                    If Len(Trim(tNFiscal.sMensagemNota)) <> 0 Or Len(Trim(tNFiscal.sMensagemCorpoNota)) <> 0 Then
                        lErro = SpedFiscal_0450(objMensagem, objSpedFiscal, objSpedRegistro0450, tNFiscal, lSeqPaiNF, colMensagem)
                        If lErro <> SUCESSO Then gError 203316
                                        
                    End If

                    lErro = SpedFiscal_C110(objMensagem, objSpedFiscal, objSpedRegistroC110, lSeqPai, lSeqPaiNF)
                    If lErro <> SUCESSO Then gError 203317

                    '************ registro C111 ********************************
                    
                    lErro = SpedFiscal_C111(objSpedFiscal, objSpedRegistroC111, lSeqPai, tNFiscal)
                    If lErro <> SUCESSO Then gError 203318
                    
                    '************ registro C112 ********************************
                    
                    lErro = SpedFiscal_C112(objSpedFiscal, objSpedRegistroC112, lSeqPai, tNFiscal)
                    If lErro <> SUCESSO Then gError 203319
                    
                    '************ registro C113 ********************************
                    
                    lErro = SpedFiscal_C113(objSpedFiscal, objSpedRegistroC113, lSeqPai, tNFiscal, objSpedRegistro0150, objSpedRegistro0175, dtDataIni, dtDataFim)
                    If lErro <> SUCESSO Then gError 203320
                    
                    '************ registro C114 ********************************
                    
                    lErro = SpedFiscal_C114(objSpedFiscal, objSpedRegistroC114, lSeqPai, tNFiscal)
                    If lErro <> SUCESSO Then gError 203321
                    
            
                If iNFe <> 1 Then
            
                    '************ registro C115 ********************************
                    lErro = SpedFiscal_C115(objSpedFiscal, objSpedRegistroC115, lSeqPai, tNFiscal, tTipoDocInfo, objFilialEmpresa)
                    If lErro <> SUCESSO Then gError 203322
            
                End If
                
                
                
            End If
        
        
    
        End If
        
        'se a nota nao for cancelada
        If iCodSit <> 2 And iCodSit <> 4 Then

            If tTipoDocInfo.iModeloArqICMS <> 65 Then
            
                If iNFe <> 1 Then
    
                    '******************** registro C120 ************************************
                    lErro = SpedFiscal_C120(objSpedFiscal, objSpedRegistroC120, lSeqPaiNF, tNFiscal, tTipoDocInfo)
                    If lErro <> SUCESSO Then gError 203323
    
    
                    '******************** registro C130 ************************************
                    lErro = SpedFiscal_C130(objSpedFiscal, objSpedRegistroC130, lSeqPaiNF, tNFiscal, tTipoDocInfo)
                    If lErro <> SUCESSO Then gError 203324
            
                    
                    '******************** registro C140 e C141 ************************************
                    lErro = SpedFiscal_C140(objSpedFiscal, objSpedRegistroC140, objSpedRegistroC141, lSeqPaiNF, tNFiscal, tTipoDocInfo, iIndPag)
                    If lErro <> SUCESSO Then gError 203325
                    
                    
                    '******************** registro C160 ************************************
                    lErro = SpedFiscal_C160(objSpedFiscal, objSpedRegistroC160, lSeqPaiNF, tNFiscal, tTipoDocInfo, dtDataIni, dtDataFim, objSpedRegistro0150, objSpedRegistro0175)
                    If lErro <> SUCESSO Then gError 203326
                    
                End If
                    
                '****************************** registro C170 *****************************************
    '                lErro = SpedFiscal_C170(objSpedFiscal, objSpedRegistroC170, objSpedRegistroC172, objSpedRegistro0200, objSpedRegistro0205, objSpedRegistro0190, objSpedRegistro0220, lSeqPaiNF, tNFiscal, tTipoDocInfo, dtDataIni, dtDataFim, iNFe, dTotalRateioNF, colUM, colSpedE510, dValDebICMS_objSpedC190, dValCredICMS_objSpedC190, colSpedC190, objFilialEmpresa, tTributacaoNF, sData, dTotalRateioNFICMSBase, dTotalRateioNFICMSValor, dTotalRateioNFICMSSubstBase, dTotalRateioNFICMSSubstValor, dTotalRateioNFIPIBaseCalculo, dTotalRateioNFIPIValor)
                
                lErro = SpedFiscal_C170(objSpedFiscal, objSpedRegistroC170, objSpedRegistroC172, objSpedRegistro0200, objSpedRegistro0205, objSpedRegistro0190, objSpedRegistro0220, lSeqPaiNF, tNFiscal, tTipoDocInfo, dtDataIni, dtDataFim, iNFe, dTotalRateioNF, colUM, colSpedE510, dValDebICMS_objSpedC190, dValCredICMS_objSpedC190, colSpedC190, objFilialEmpresa, tTributacaoNF, sData, colSpedC190UF, objFisSpedAux)
                If lErro <> SUCESSO Then gError 203327
            
                lErro = SpedFiscal_C190(tTributacaoNF, objSpedFiscal, objSpedRegistroC190, lSeqPaiNF, tNFiscal, tTipoDocInfo, colSpedC190, iNFe, dTotalRateioNF, objFisSpedAux)
                If lErro <> SUCESSO Then gError 203328
            
            Else
            
                lErro = SpedFiscal_C170_NFCe(objSpedFiscal, objSpedRegistroC170, objSpedRegistroC172, objSpedRegistro0200, objSpedRegistro0205, objSpedRegistro0190, objSpedRegistro0220, lSeqPaiNF, tNFiscal, tTipoDocInfo, dtDataIni, dtDataFim, iNFe, dTotalRateioNF, colUM, colSpedE510, dValDebICMS_objSpedC190, dValCredICMS_objSpedC190, colSpedC190, objFilialEmpresa, tTributacaoNF, sData, colSpedC190UF, objFisSpedAux)
                If lErro <> SUCESSO Then gError 203327
            
                lErro = SpedFiscal_C190_NFCe(tTributacaoNF, objSpedFiscal, objSpedRegistroC190, lSeqPaiNF, tNFiscal, tTipoDocInfo, colSpedC190, iNFe, dTotalRateioNF, objFisSpedAux)
                If lErro <> SUCESSO Then gError 203328
            
            End If
            
            '********************** registro C190 *******************************
'                lErro = SpedFiscal_C190(tTributacaoNF, objSpedFiscal, objSpedRegistroC190, lSeqPaiNF, tNFiscal, tTipoDocInfo, colSpedC190, iNFe, dTotalRateioNF, dTotalRateioNFICMSBase, dTotalRateioNFICMSValor, dTotalRateioNFICMSSubstBase, dTotalRateioNFICMSSubstValor, dTotalRateioNFIPIBaseCalculo, dTotalRateioNFIPIValor)
            
        
        '********************** registro C195 *******************************
        
'            'se for nota fiscal (codigo 01) ou NFe
'            If tTipoDocInfo.iModeloArqICMS = 1 Or iNFe = 1 Then
'
'                sObs = String(STRING_OBSERVACAO, 0)
'
'                lErro = Comando_Executar(alComando(28), "SELECT ObservacaoLivFisc FROM LivRegES, LivRegESLinha WHERE TipoNumIntDocOrigem = 0 AND NumIntDocOrigem = ? AND LivRegES.NumIntDoc = NumIntDocRegES", _
'                sObs, tNFiscal.lNumIntDoc)
'                If lErro <> AD_SQL_SUCESSO Then gError 203008
'
'                lErro = Comando_BuscarPrimeiro(alComando(28))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203009
'
'                Do While lErro = AD_SQL_SUCESSO
'
'                    '********************** registro 0460 *******************************
'
'                    lErro = SpedFiscal_0460(sObs, objSpedFiscal, objSpedRegistro0460, sCodObs)
'                    If lErro <> SUCESSO Then gError 999999
'
'                    objSpedFiscal.sTipo = "C195"
'                    objSpedRegistroC195.sTipo = objSpedFiscal.sTipo
'
'                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
'                    objSpedRegistroC195.iCount = objSpedRegistroC195.iCount + 1
'
'                    objSpedFiscal.sRegistro = "|C195|" & sCodObs & "||"
'
'                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203123
'
'                    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
'                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
'                    If lErro <> AD_SQL_SUCESSO Then gError 203124
'
'                    lErro = Comando_BuscarProximo(alComando(28))
'                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203009
'
'                Loop
'
'            End If

        End If

    End If
    
 '   If objTela.dValorAtual = 540 Then MsgBox "540"
    
        If Not objTela Is Nothing Then

    lErro = objTela.Processa_Item
    If lErro <> SUCESSO Then gError 204133
    
End If

    SpedFiscal_Trata_ItemNF = SUCESSO
    
    Exit Function
    
Erro_SpedFiscal_Trata_ItemNF:

    SpedFiscal_Trata_ItemNF = gErr
    
    Select Case gErr
    
        Case 203288
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 203289
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 203290, 203294, 203296 To 203300, 203305, 203306, 203316 To 203328, 203337 To 203340, 203343, 203346 To 203368

        Case 203291, 203292
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISVERSAOLEIAUTE", gErr)

        Case 203293
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISVERSAOLEIAUTE_NAO_CADASTRADO", gErr)
        
        Case 203295
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, iFilialEmpresa)
        
        Case 203301, 203302, 203329, 204129, 204130
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case 203303, 203304
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFEFEDPROTNFE", gErr)
        
        Case 203307, 203308
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFEFEDRETCANCNFE", gErr)
        
        Case 203309, 203310
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAGTODOS", gErr)
        
        Case 203311, 203312
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTODOS", gErr)
        
        Case 203313, 203314
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAONF", gErr)
        
        Case 203315, 203336, 203345
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case 203330, 203331, 203342, 204131, 204132
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ECF1", gErr)
        
        Case 203332, 203333, 203341
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJAARQFISMESTRE", gErr)
        
        Case 203334, 203335, 203344
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
        
        Case 203369
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 204133, 204134, 204141, 204920
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203370)
     
    End Select
     
    Exit Function

End Function

Private Function SpedFiscal_D100_Trata_ItemNF(dtDataIni As Date, dtDataFim As Date, objSpedFiscal As ClassSpedFiscal, tNFiscal As typeNFiscal, objSpedRegistro0150 As ClassSpedRegistro, objSpedRegistro0175 As ClassSpedRegistro, tTipoDocInfo As typeTipoDocInfo, alComando() As Long, objSpedRegistro0450 As ClassSpedRegistro, colMensagem As Collection, objSpedRegistroD100 As ClassSpedRegistro, objSpedRegistroD190 As ClassSpedRegistro, dValCredICMS_objSpedD190 As Double) As Long

Dim lErro As Long
Dim schNFe As String
Dim iNFe As Integer
Dim sCodParticipante As String
Dim iCodSit As Integer
Dim iAmbiente As Integer
Dim iCondPagto As Integer
Dim iIndPag As Integer
Dim iFreteRespons As Integer
Dim tTributacaoNF As typeTributacaoNF
Dim lSeqPaiNF As Long
Dim dValorMercadorias As Double
Dim sData As String
Dim lSeqPai As Long
Dim objMensagem As ClassMensagem
Dim dTotalRateioNF As Double
Dim vbMsgRes As VbMsgBoxResult
Dim objTrib As ClassTributacaoDoc
Dim sIBGEOrig As String, sIBGEDest As String

On Error GoTo Erro_SpedFiscal_D100_Trata_ItemNF

    lErro = NF_Obter_Situacao(alComando(5), tNFiscal, tTipoDocInfo, iCodSit)
    If lErro <> SUCESSO Then gError 204920
    
'    If tTipoDocInfo.iComplementar = 1 Then
'        iCodSit = 6
'    Else
'        iCodSit = 0
'    End If
'
'    If tNFiscal.iStatus = STATUS_CANCELADO Then iCodSit = 2

'    If tNFiscal.dValorFrete <> 0 Then
'        If tNFiscal.iFreteRespons = 0 Or tNFiscal.iFreteRespons = 1 Then
'            iFreteRespons = 1
'        Else
'            iFreteRespons = 2
'        End If
'    Else
'        iFreteRespons = 9
'    End If
        
    Select Case tNFiscal.iFreteRespons
    
        Case 0, 1 'Emitente
            If tNFiscal.dtDataEmissao < #1/1/2012# Then
                iFreteRespons = 1
            Else
                iFreteRespons = 0
            End If
        Case 2 'Destinatário
            If tNFiscal.dtDataEmissao < #1/1/2012# Then
                iFreteRespons = 2
            Else
                iFreteRespons = 1
            End If
        Case 3 'Terceiros
            If tNFiscal.dtDataEmissao < #1/1/2012# Then
                iFreteRespons = 0
            Else
                iFreteRespons = 2
            End If
        Case 4 'Sem Frete
            iFreteRespons = 9

    End Select
    
    iNFe = SerieEletronica(tNFiscal.sSerie)

    'pesquisa a tributacaoNF
    lErro = Comando_Executar(alComando(2), "SELECT ICMSBase, ICMSValor, ICMSSubstBase, ICMSSubstValor, IPIValor, PISCredito, COFINSCredito, PISRetido, COFINSRetido, ISSAliquota, ICMSCredito FROM TributacaoNF WHERE NumIntDoc = ?", _
    tTributacaoNF.dICMSBase, tTributacaoNF.dICMSValor, tTributacaoNF.dICMSSubstBase, tTributacaoNF.dICMSSubstValor, tTributacaoNF.dIPIValor, tTributacaoNF.dPISCredito, tTributacaoNF.dCOFINSCredito, tTributacaoNF.dPISRetido, tTributacaoNF.dCOFINSRetido, tTributacaoNF.dISSAliquota, tTributacaoNF.dICMSCredito, tNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 203564

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203565
    
    If iCodSit <> 2 And iCodSit <> 4 Then
        'se é uma nota de cliente
        If tNFiscal.lCliente > 0 Then
            lErro = SpedFiscal_0150_Cliente(tNFiscal.lCliente, tNFiscal.iFilialCli, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
            If lErro <> SUCESSO Then gError 203562
        Else
    
            lErro = SpedFiscal_0150_Fornecedor(tNFiscal.lFornecedor, tNFiscal.iFilialForn, objSpedFiscal, objSpedRegistro0150, objSpedRegistro0175, sCodParticipante, dtDataIni, dtDataFim)
            If lErro <> SUCESSO Then gError 203563
        End If
    End If

    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    lSeqPaiNF = objSpedFiscal.lSeq
        
    'se a nota nao for cancelada
    If iCodSit <> 2 And iCodSit <> 4 Then
    
        Set objTrib = New ClassTributacaoDoc
        objTrib.lNumIntDoc = tNFiscal.lNumIntDoc
        objTrib.iTipoDoc = TIPODOC_TRIB_NF
        lErro = NF_Le_BasePropria(objTrib)
        If lErro <> SUCESSO Then gError 203566
        
        tTributacaoNF.dICMSBase = objTrib.dICMSBase
        tTributacaoNF.dICMSValor = objTrib.dICMSValor
        tTributacaoNF.dICMSCredito = objTrib.dICMSCredito

        tTributacaoNF.dIPIBase = objTrib.dIPIBase
        tTributacaoNF.dIPIValor = objTrib.dIPIValor
        tTributacaoNF.dIPICredito = objTrib.dIPICredito
    
        Set objMensagem = New ClassMensagem

        If Len(Trim(tNFiscal.sMensagemNota)) <> 0 Or Len(Trim(tNFiscal.sMensagemCorpoNota)) <> 0 Then
            lErro = SpedFiscal_0450(objMensagem, objSpedFiscal, objSpedRegistro0450, tNFiscal, lSeqPaiNF, colMensagem)
            If lErro <> SUCESSO Then gError 203566
        End If
    
        objSpedFiscal.sRegistro = "|D100|0|1|" & sCodParticipante & "|" & IIf(iNFe = 1, "57", Format(tTipoDocInfo.iModeloArqICMS, "00")) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(iCodSit, "00") & "|" & Ajusta_Serie(Desconverte_Serie_Eletronica(tNFiscal.sSerie)) & "||" & tNFiscal.lNumNotaFiscal & "|" & tNFiscal.sChvNFe & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tNFiscal.dtDataEmissao, "ddmmyyyy") & "|" & Format(tNFiscal.dtDataEntrada, "ddmmyyyy") & "|0||"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tNFiscal.dValorTotal, "Fixed") & "|" & IIf(tNFiscal.dValorDesconto = 0, 0, Format(tNFiscal.dValorDesconto, "Fixed")) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & iFreteRespons & "|" & Format(tNFiscal.dValorProdutos, "Fixed") & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSBase = 0, 0, Format(tTributacaoNF.dICMSBase, "Fixed")) & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(tTributacaoNF.dICMSCredito = 0, 0, Format(tTributacaoNF.dICMSCredito, "Fixed")) & "||"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objMensagem.iCodigo = 0, "", CStr(objMensagem.iCodigo)) & "||"
        
        
        If giVersao >= 12 Then
            '24 COD_MUN_OR
            'IG
            'Código do município de origem do serviço, conforme a
            'tabela IBGE (Preencher com 9999999, se Exterior)
            'N 007* - OC OC
            '25 COD_MUN_DE
            'ST
            'Código do município de destino, conforme a tabela IBGE
            '(Preencher com 9999999, se Exterior)
            'N 007* - OC OC
            
            lErro = CTe_Obtem_CodsIBGE(alComando(6), alComando(7), alComando(8), alComando(9), tNFiscal.lNumIntDoc, sIBGEOrig, sIBGEDest)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & sIBGEOrig & "|" & sIBGEDest & "|"
            
        End If
        
    Else
    
        objSpedFiscal.sRegistro = "|D100|0|1||" & Format(tTipoDocInfo.iModeloArqICMS, "00") & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(iCodSit, "00") & "|" & Ajusta_Serie(Desconverte_Serie_Eletronica(tNFiscal.sSerie)) & "||" & tNFiscal.lNumNotaFiscal & "|" & tNFiscal.sChvNFe & "||||||||||||||"
    
        If giVersao >= 12 Then objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "||"
    
    End If
    
    objSpedFiscal.sTipo = "D100"
    objSpedRegistroD100.sTipo = objSpedFiscal.sTipo
    objSpedRegistroD100.iCount = objSpedRegistroD100.iCount + 1
    
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 203567
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, lSeqPaiNF, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 203568


    'se a nota nao for cancelada
    If iCodSit <> 2 And iCodSit <> 4 Then

        '*****************  registro D190 ******************************************************
        lErro = SpedFiscal_D190(tNFiscal, objSpedFiscal, objSpedRegistroD190, dtDataIni, dtDataFim, dValCredICMS_objSpedD190, lSeqPaiNF)
        If lErro <> SUCESSO Then gError 203569

    End If

    SpedFiscal_D100_Trata_ItemNF = SUCESSO

    Exit Function
    
Erro_SpedFiscal_D100_Trata_ItemNF:

   SpedFiscal_D100_Trata_ItemNF = gErr

   Select Case gErr
    
        Case 203559
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203560, 203561, 203570
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
    
        Case 203562, 203563, 203566, 203569
    
        Case 203564, 203565
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAONF", gErr)
    
        Case 203567
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203568
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
            
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203571)
     
    End Select

End Function

Private Function Retorna_UF_Subst_Tributaria(objFilialEmpresa As AdmFiliais, tNFiscal As typeNFiscal, tTipoDocInfo As typeTipoDocInfo) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_Retorna_UF_Subst_Tributaria

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204921
    Next
    
    'se for nota de entrada
    If tTipoDocInfo.iTipo = 1 Or tTipoDocInfo.iTipo = 3 Then
    
    
            If tNFiscal.lCliente <> 0 Then
        
                tNFiscal.sUFSubstTrib = String(STRING_ESTADO, 0)
        
                lErro = Comando_Executar(alComando(1), "SELECT E.SiglaEstado FROM Enderecos AS E, FiliaisClientes AS F WHERE CodCliente =? AND CodFilial = ? AND F.Endereco = E.Codigo", tNFiscal.sUFSubstTrib, tNFiscal.lCliente, tNFiscal.iFilialCli)
                If lErro <> AD_SQL_SUCESSO Then gError 204922
            
                lErro = Comando_BuscarPrimeiro(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204923
        
            Else
            
                tNFiscal.sUFSubstTrib = String(STRING_ESTADO, 0)
            
                'Lê a Filial do Fornecedor que será lockada
                lErro = Comando_Executar(alComando(2), "SELECT E.SiglaEstado FROM Enderecos AS E, FiliaisFornecedores AS F WHERE CodFornecedor =? AND CodFilial = ? AND F.Endereco = E.Codigo", tNFiscal.sUFSubstTrib, tNFiscal.lFornecedor, tNFiscal.iFilialForn)
                If lErro <> AD_SQL_SUCESSO Then gError 204924
            
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204925
    
            End If
            
            tNFiscal.sUFRemet = tNFiscal.sUFSubstTrib
            tNFiscal.sUFDest = objFilialEmpresa.objEndereco.sSiglaEstado
    
    Else
    
        'se for nota de saida
    
        If tNFiscal.lCliente <> 0 Then
    
            tNFiscal.sUFSubstTrib = String(STRING_ESTADO, 0)
    
            lErro = Comando_Executar(alComando(1), "SELECT E.SiglaEstado FROM Enderecos AS E, FiliaisClientes AS F WHERE CodCliente =? AND CodFilial = ? AND F.Endereco = E.Codigo", tNFiscal.sUFSubstTrib, tNFiscal.lCliente, tNFiscal.iFilialCli)
            If lErro <> AD_SQL_SUCESSO Then gError 204926
        
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204927
    
        Else
        
            tNFiscal.sUFSubstTrib = String(STRING_ESTADO, 0)
        
            'Lê a Filial do Fornecedor que será lockada
            lErro = Comando_Executar(alComando(2), "SELECT E.SiglaEstado FROM Enderecos AS E, FiliaisFornecedores AS F WHERE CodFornecedor =? AND CodFilial = ? AND F.Endereco = E.Codigo", tNFiscal.sUFSubstTrib, tNFiscal.lFornecedor, tNFiscal.iFilialForn)
            If lErro <> AD_SQL_SUCESSO Then gError 204928
        
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204929

        End If

        tNFiscal.sUFDest = tNFiscal.sUFSubstTrib
        tNFiscal.sUFRemet = objFilialEmpresa.objEndereco.sSiglaEstado

    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Retorna_UF_Subst_Tributaria = SUCESSO

    Exit Function
    
Erro_Retorna_UF_Subst_Tributaria:

    Retorna_UF_Subst_Tributaria = gErr

    Select Case gErr
    
        Case 204921
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204922, 204923, 204926, 204927
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr, tNFiscal.lCliente, tNFiscal.iFilialCli)
    
        Case 204924, 204925, 204928, 204929
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES", gErr, tNFiscal.lFornecedor, tNFiscal.iFilialForn)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204930)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_E200(ByVal objFilialEmpresa As AdmFiliais, objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE200 As ClassSpedRegistro, ByVal objSpedRegistroE210 As ClassSpedRegistro, ByVal objSpedRegistroE220 As ClassSpedRegistro, ByVal objSpedRegistroE250 As ClassSpedRegistro, colSpedC190UF As Collection, colSpedC590UF As Collection, colSpedD590UF As Collection, dtDataIni As Date, dtDataFim As Date) As Long
  
Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim sInscricaoEstadual As String
Dim sUF As String
Dim sIE As String
Dim objSpedC190UF As ClassSpedC190UF
Dim objSpedC590UF As ClassSpedC590UF
Dim objSpedD590UF As ClassSpedD590UF
Dim colUF As New Collection
Dim vUF As Variant
Dim iAchou As Integer
Dim lSeqE200 As Long

On Error GoTo Erro_SpedFiscal_E200
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204946
    Next

    For Each objSpedC190UF In colSpedC190UF
            
'        MsgBox ("C190 = " & objSpedC190UF.sUF)

        If objSpedC190UF.dValorBCICMSST > 0 Or objSpedC190UF.dValorICMSST > 0 Then
            
            iAchou = 0
                
            For Each vUF In colUF
                If vUF = objSpedC190UF.sUF Then
                    iAchou = 1
                    Exit For
                End If
                
            Next
            
            If iAchou = 0 Then
            
                colUF.Add objSpedC190UF.sUF
                
            End If
            
        End If
        
    Next
        
    For Each objSpedC590UF In colSpedC590UF
            
        If objSpedC590UF.dValorBCICMSST > 0 Or objSpedC590UF.dValorICMSST > 0 Then
            
            iAchou = 0
                
    '        MsgBox ("C590 = " & objSpedC590UF.sUF)
                
            For Each vUF In colUF
                If vUF = objSpedC590UF.sUF Then
                    iAchou = 1
                    Exit For
                End If
                
            Next
            
            If iAchou = 0 Then
            
                colUF.Add objSpedC590UF.sUF
                
            End If
        
        End If
        
    Next
        
    
    
    For Each objSpedD590UF In colSpedD590UF
                       
        If objSpedD590UF.dValorBCICMSST > 0 Or objSpedD590UF.dValorICMSST > 0 Then
                           
    '        MsgBox ("D590 = " & objSpedD590UF.sUF)
                
            For Each vUF In colUF
                If vUF = objSpedD590UF.sUF Then
                    iAchou = 1
                    Exit For
                End If
                
            Next
            
            If iAchou = 0 Then
            
                colUF.Add objSpedD590UF.sUF
                
            End If
        
        End If
        
    Next


    sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    sUF = String(STRING_ESTADO, 0)
    
    lErro = Comando_Executar(alComando(3), "SELECT UF, InscricaoEstadual FROM ContribuinteSubstituto WHERE FilialEmpresa = ? ", _
    sUF, sInscricaoEstadual, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 204947

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204948
    
    Do While lErro = AD_SQL_SUCESSO
    
        iAchou = 0
            
            
'        MsgBox ("Contribuinte  = " & sUF)
            
        For Each vUF In colUF
            If vUF = sUF Then
                iAchou = 1
                Exit For
            End If
            
        Next
        
        If iAchou = 0 Then
        
            colUF.Add sUF
            
        End If
             
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204949
             
    Loop
    
    
    For Each vUF In colUF
    
        If vUF <> "EX" Then
    
            objSpedFiscal.sTipo = "E200"
            objSpedRegistroE200.sTipo = objSpedFiscal.sTipo
            objSpedRegistroE200.iCount = objSpedRegistroE200.iCount + 1
        
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            
            lSeqE200 = objSpedFiscal.lSeq
        
            objSpedFiscal.sRegistro = "|E200|" & vUF & "|" & Format(dtDataIni, "ddmmyyyy") & "|" & Format(dtDataFim, "ddmmyyyy") & "|"
        
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 204950
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
            If lErro <> AD_SQL_SUCESSO Then gError 204951
            
            sUF = vUF
            
            lErro = SpedFiscal_E210(objSpedFiscal, objSpedRegistroE210, objSpedRegistroE220, objSpedRegistroE250, objFilialEmpresa, dtDataIni, dtDataFim, colSpedC190UF, colSpedC590UF, colSpedD590UF, lSeqE200, sUF)
            If lErro <> SUCESSO Then gError 204952
            
        End If
        
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E200 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E200:

   SpedFiscal_E200 = gErr

   Select Case gErr
    
        Case 204946
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204947 To 204949
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTRIBUINTESUBSTITUTO", gErr)
    
        Case 204950
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204951
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
            
        Case 204952
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204952)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Private Function SpedFiscal_E210(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE210 As ClassSpedRegistro, ByVal objSpedRegistroE220 As ClassSpedRegistro, ByVal objSpedRegistroE250 As ClassSpedRegistro, ByVal objFilialEmpresa As AdmFiliais, ByVal dtDataIni As Date, ByVal dtDataFim As Date, colSpedC190UF As Collection, colSpedC590UF As Collection, colSpedD590UF As Collection, ByVal lSeqE200 As Long, ByVal sUF As String) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim dSaldoCredorInicial As Double
Dim dSaldoApurado As Double
Dim dSaldoCredTransp As Double
Dim dValorDevolvido As Double
Dim dValorRessarc As Double
Dim dValorOutCred As Double
Dim objSpedC190UF As ClassSpedC190UF
Dim objSpedC590UF As ClassSpedC590UF
Dim objSpedD590UF As ClassSpedD590UF
Dim dValorRetencao As Double
Dim dValorOutDeb As Double
Dim dValorAjCreditos As Double
Dim dValorEstCreditos As Double
Dim dValorAjDebitos As Double
Dim dValorEstDebitos As Double
Dim dSaldoDevedorAntes As Double
Dim dValorDeducoes As Double
Dim dValorICMSSTRecolher As Double
Dim dSaldoTransportar As Double
Dim lSeqPai As Long
Dim iIndMov As Integer


On Error GoTo Erro_SpedFiscal_E210
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204939
    Next


    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedRegistroE210.iCount = objSpedRegistroE210.iCount + 1
    lSeqPai = objSpedFiscal.lSeq
    
    '*************** registro E220 **********************
    lErro = SpedFiscal_E220(objSpedFiscal, objSpedRegistroE220, objFilialEmpresa, dtDataIni, dtDataFim, lSeqPai, dValorAjDebitos, dValorEstCreditos, dValorAjCreditos, dValorEstDebitos)
    If lErro <> SUCESSO Then gError 204940
    
    'Verifica se o Item de apuração ICMS está cadastrada
    lErro = Comando_Executar(alComando(2), "SELECT SaldoCredorInicial FROM RegApuracaoICMSST WHERE  DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ? AND UF=?", _
    dSaldoCredorInicial, dtDataIni, dtDataFim, objFilialEmpresa.iCodFilial, sUF)
    If lErro <> AD_SQL_SUCESSO Then gError 204941
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204942
    
    If lErro <> AD_SQL_SUCESSO Then dSaldoCredorInicial = 0
    
    For Each objSpedC190UF In colSpedC190UF
    
        If objSpedC190UF.sUF = sUF Then
        
            If objSpedC190UF.sCFOP = "1410" Or objSpedC190UF.sCFOP = "1411" Or objSpedC190UF.sCFOP = "1660" Or objSpedC190UF.sCFOP = "1661" Or objSpedC190UF.sCFOP = "1662" Or _
                objSpedC190UF.sCFOP = "2410" Or objSpedC190UF.sCFOP = "2411" Or objSpedC190UF.sCFOP = "2660" Or objSpedC190UF.sCFOP = "2661" Or objSpedC190UF.sCFOP = "2662" Then
                    dValorDevolvido = dValorDevolvido + objSpedC190UF.dValorICMSST
            End If
            
        End If
    
    Next
    
    For Each objSpedC190UF In colSpedC190UF
    
        If objSpedC190UF.sUF = sUF Then
        
            If objSpedC190UF.sCFOP = "1603" Or objSpedC190UF.sCFOP = "2603" Then
                    dValorRessarc = dValorRessarc + objSpedC190UF.dValorICMSST
            End If
            
        End If
    
    Next
    
    dValorOutCred = 0
    For Each objSpedC190UF In colSpedC190UF
    
        If objSpedC190UF.sUF = sUF And (left(objSpedC190UF.sCFOP, 1) = "1" Or left(objSpedC190UF.sCFOP, 1) = "2") Then
        
            If objSpedC190UF.sCFOP <> "1410" And objSpedC190UF.sCFOP <> "1411" And objSpedC190UF.sCFOP <> "1660" And objSpedC190UF.sCFOP <> "1661" And objSpedC190UF.sCFOP <> "1662" And _
                objSpedC190UF.sCFOP <> "2410" And objSpedC190UF.sCFOP <> "2411" And objSpedC190UF.sCFOP <> "2660" And objSpedC190UF.sCFOP <> "2661" And objSpedC190UF.sCFOP <> "2662" Then
                    dValorOutCred = dValorOutCred + objSpedC190UF.dValorICMSST
            End If
            
        End If
    
    Next
    
    dValorOutCred = dValorOutCred + dValorAjCreditos + dValorEstDebitos
    
    
    For Each objSpedC190UF In colSpedC190UF
    
        If objSpedC190UF.sUF = sUF And (left(objSpedC190UF.sCFOP, 1) = "5" Or left(objSpedC190UF.sCFOP, 1) = "6") Then
            dValorRetencao = dValorRetencao + objSpedC190UF.dValorICMSST
        End If
    
    Next
    
    For Each objSpedC590UF In colSpedC590UF
    
        If objSpedC590UF.sUF = sUF And (left(objSpedC590UF.sCFOP, 1) = "5" Or left(objSpedC590UF.sCFOP, 1) = "6") Then
            dValorRetencao = dValorRetencao + objSpedC590UF.dValorICMSST
        End If
    
    Next
    
    For Each objSpedD590UF In colSpedD590UF
    
        If objSpedD590UF.sUF = sUF And (left(objSpedD590UF.sCFOP, 1) = "5" Or left(objSpedD590UF.sCFOP, 1) = "6") Then
            dValorRetencao = dValorRetencao + objSpedD590UF.dValorICMSST
        End If
    
    Next
    
    dValorOutDeb = dValorOutDeb + dValorAjDebitos + dValorEstCreditos
    
    
    dSaldoDevedorAntes = dValorRetencao + dValorOutDeb + dValorAjDebitos - (dSaldoCredorInicial + dValorDevolvido + dValorRessarc + dValorOutCred + dValorAjCreditos)
    
    dValorICMSSTRecolher = dSaldoDevedorAntes - dValorDeducoes
    
    If dValorICMSSTRecolher < 0 Then
        dSaldoTransportar = Abs(dValorICMSSTRecolher)
        dValorICMSSTRecolher = 0
    Else
        dSaldoTransportar = 0
    End If
    
    
    If dValorRetencao + dValorOutDeb + dValorAjDebitos + (dSaldoCredorInicial + dValorDevolvido + dValorRessarc + dValorOutCred + dValorAjCreditos) > 0 Then
        iIndMov = 1
    Else
        iIndMov = 0
    End If
    
    
    If dSaldoDevedorAntes < 0 Then dSaldoDevedorAntes = 0
    
    objSpedFiscal.sTipo = "E210"
    objSpedRegistroE210.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.sRegistro = "|E210|" & iIndMov & "|" & Format(dSaldoCredorInicial, "fixed") & "|" & Format(dValorDevolvido, "fixed") & "|" & Format(dValorRessarc, "fixed") & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dValorOutCred, "fixed") & "|" & Format(dValorAjCreditos, "fixed") & "|" & Format(dValorRetencao, "fixed") & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dValorOutDeb, "fixed") & "|" & Format(dValorAjDebitos, "fixed") & "|" & Format(dSaldoDevedorAntes, "fixed") & "|"
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dValorDeducoes, "fixed") & "|" & Format(dValorICMSSTRecolher, "fixed") & "|" & Format(dSaldoTransportar, "fixed") & "|0|"
    
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204943

    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai) VALUES (?,?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, lSeqPai, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqE200)
    If lErro <> AD_SQL_SUCESSO Then gError 204944

    lErro = SpedFiscal_E250(objSpedFiscal, objSpedRegistroE250, objFilialEmpresa, dtDataIni, dtDataFim, lSeqPai, dValorICMSSTRecolher, sUF)
    If lErro <> SUCESSO Then gError 204955

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E210 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E210:

   SpedFiscal_E210 = gErr

   Select Case gErr
    
        Case 204939
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204940, 204955
    
        Case 204941, 204942
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMSST", gErr)
    
        Case 204943
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204944
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204945)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Function SpedFiscal_E220(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE220 As ClassSpedRegistro, ByVal objFilialEmpresa As AdmFiliais, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal lSeqPai As Long, dTotAjDebitos As Double, dTotEstCreditos As Double, dTotAjCreditos As Double, dTotEstDebitos As Double) As Long
        
Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim tRegApuracaoItens As typeRegApuracaoItens
Dim iRegApuracaoICMS As Integer
Dim iRegApuracaoICMSAtual As Integer
Dim iSequencial As Integer
Dim sCodAjApur As String
Dim iTipoReg As Integer

On Error GoTo Erro_SpedFiscal_E220
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 204932
    Next
    
    iRegApuracaoICMSAtual = -1
        
    tRegApuracaoItens.sDescricao = String(STRING_DESCRICAO_APURACAO, 0)
        
    'Verifica se o Item de apuração ICMS está cadastrada
    lErro = Comando_Executar(alComando(2), "SELECT Valor, R.Descricao, RegApuracaoICMS, TipoReg FROM RegApuracaoICMSSTItem AS R, TiposRegApuracaoICMS WHERE TipoReg = Codigo AND Data >= ? AND Data <= ? AND FilialEmpresa = ? AND (RegApuracaoICMS = 0 OR RegApuracaoICMS = 1 OR RegApuracaoICMS = 2 OR RegApuracaoICMS = 3 OR RegApuracaoICMS = -1)  ORDER BY RegApuracaoICMS ", _
    tRegApuracaoItens.dValor, tRegApuracaoItens.sDescricao, iRegApuracaoICMS, iTipoReg, dtDataIni, dtDataFim, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 204933
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204934

    Do While lErro = AD_SQL_SUCESSO
        
        If iRegApuracaoICMSAtual <> iRegApuracaoICMS Then
            iRegApuracaoICMSAtual = iRegApuracaoICMS
            iSequencial = 1
        Else
            iSequencial = iSequencial + 1
        End If
        
        If iRegApuracaoICMS = ICMS_OUTROS_DEBITOS Then
            dTotAjDebitos = dTotAjDebitos + tRegApuracaoItens.dValor
        ElseIf iRegApuracaoICMS = ICMS_ESTORNO_CREDITOS Then
            dTotEstCreditos = dTotEstCreditos + tRegApuracaoItens.dValor
        ElseIf iRegApuracaoICMS = ICMS_OUTROS_CREDITOS Then
            dTotAjCreditos = dTotAjCreditos + tRegApuracaoItens.dValor
        ElseIf iRegApuracaoICMS = ICMS_ESTORNO_DEBITOS Then
            dTotEstDebitos = dTotEstDebitos + tRegApuracaoItens.dValor
        ElseIf iRegApuracaoICMS = -1 And iTipoReg = 12 Then
            dTotAjCreditos = dTotAjCreditos + tRegApuracaoItens.dValor
        End If
        
        sCodAjApur = objFilialEmpresa.objEndereco.sSiglaEstado & "1" & iRegApuracaoICMS & Format(iSequencial, "0000")
        
        '******************* registro E111 *******************
    
        objSpedFiscal.sTipo = "E220"
        objSpedRegistroE220.sTipo = objSpedFiscal.sTipo
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroE220.iCount = objSpedRegistroE220.iCount + 1
        
        objSpedFiscal.sRegistro = "|E220|" & sCodAjApur & "|" & SFF_Texto(tRegApuracaoItens.sDescricao) & "|" & Format(tRegApuracaoItens.dValor, "fixed") & "|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 204935
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
        If lErro <> AD_SQL_SUCESSO Then gError 204936

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204937

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E220 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E220:

   SpedFiscal_E220 = gErr

   Select Case gErr
    
        Case 204932
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 204933, 204934, 204937
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMSSTITEM", gErr)
    
        Case 204935
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 204936
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204938)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_E250(ByVal objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroE250 As ClassSpedRegistro, ByVal objFilialEmpresa As AdmFiliais, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal lSeqPai As Long, ByVal dValorICMSSTRecolher As Double, ByVal sUF As String) As Long
        
Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim tRegApuracaoItens As typeRegApuracaoItens
Dim iSequencial As Integer
Dim sCodAjApur As String
Dim tGuiasICMS As typeGuiasICMS
Dim dTotalRecolher As Double
Dim iFilialEmpresa As Integer


On Error GoTo Erro_SpedFiscal_E250
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203609
    Next
    
    iFilialEmpresa = objFilialEmpresa.iCodFilial

    tGuiasICMS.sCodReceita = String(STRING_GUIASICMS_CODRECEITA, 0)
    tGuiasICMS.sCodObrigRecolher = String(STRING_GUIASICMS_CODOBRIGRECOLHER, 0)

    'Verifica se tem GuiaICMSST cadastrada
    lErro = Comando_ExecutarPos(alComando(3), "SELECT CodReceita, CodObrigRecolher, Vencimento, Valor FROM GuiasICMSST WHERE ApuracaoDe >= ? And ApuracaoDe <= ? AND ApuracaoAte >= ? And ApuracaoAte <= ? AND FilialEmpresa = ? AND UF = ? ", 0, _
    tGuiasICMS.sCodReceita, tGuiasICMS.sCodObrigRecolher, tGuiasICMS.dtDataVencimento, tGuiasICMS.dValor, dtDataIni, dtDataFim, dtDataIni, dtDataFim, iFilialEmpresa, sUF)
    If lErro <> AD_SQL_SUCESSO Then gError 203615
    
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203616

    Do While lErro = AD_SQL_SUCESSO
        
        dTotalRecolher = dTotalRecolher + tGuiasICMS.dValor
        
        objSpedFiscal.sTipo = "E250"
        objSpedRegistroE250.sTipo = objSpedFiscal.sTipo
            
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedRegistroE250.iCount = objSpedRegistroE250.iCount + 1
        
        objSpedFiscal.sRegistro = "|E250|" & SFF_Texto(tGuiasICMS.sCodObrigRecolher) & "|" & Format(tGuiasICMS.dValor, "Fixed") & "|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(tGuiasICMS.dtDataVencimento, "ddmmyyyy") & "|" & SFF_Texto(tGuiasICMS.sCodReceita) & "|||||" & Format(dtDataIni, "mmyyyy") & "|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 203617
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
        If lErro <> AD_SQL_SUCESSO Then gError 203618

        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203619

    Loop


    'os valores tem que ser iguais...senao erro...mostrar os totais na mensagem
 '   If Abs(dValorICMSSTRecolher - dTotalRecolher) > DELTA_VALORMONETARIO Then gError 203620

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E250 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E250:

   SpedFiscal_E250 = gErr

   Select Case gErr
    
        Case 203609
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203610, 203611, 203614
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GNRICMS", gErr)
    
        Case 203612, 203617
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203613, 203618
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
    
        Case 203615, 203616, 203619
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GUIAICMS", gErr)
    
        Case 203620
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ICMS_RECOLHER_DIFERE_GUIAS", gErr, dValorICMSSTRecolher, dTotalRecolher)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203621)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_E300(ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer, objSpedFiscal As ClassSpedFiscal
Dim vUF As Variant

On Error GoTo Erro_SpedFiscal_E300
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201223
    Next
    
    For Each vUF In objFisSpedAux.colC101UF
    
        Set objSpedFiscal = objFisSpedAux.objSpedFiscal
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objFisSpedAux.objSpedRegistroE300.iCount = objFisSpedAux.objSpedRegistroE300.iCount + 1
           
        objSpedFiscal.sTipo = "E300"
        objFisSpedAux.objSpedRegistroE300.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.sRegistro = "|E300|" '01 REG
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & vUF & "|" '02 UF
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(objFisSpedAux.dtDataIni, "ddmmyyyy") & "|" '03 DT_INI
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(objFisSpedAux.dtDataFim, "ddmmyyyy") & "|" '04 DT_FIN
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 201224
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 201225
        
        lErro = SpedFiscal_E310(objFisSpedAux, objSpedFiscal.lSeq, vUF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E300 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E300:

   SpedFiscal_E300 = gErr

   Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 201223
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201224
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201225
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204078)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_E310(ByVal objFisSpedAux As ClassFISSpedAux, ByVal lSeqPai As Long, ByVal sUF As String) As Long
'REGISTRO E310: APURAÇÃO DO FUNDO DE COMBATE À POBREZA E DO ICMS
'DIFERENCIAL DE ALÍQUOTA – UF ORIGEM/DESTINO EC 87/15. (VÁLIDO A PARTIR
'DE 01/01/2017)
'Este registro tem por objetivo informar valores relativos à apuração do ICMS – Diferencial de Alíquota e Fundo
'de Combate à Pobreza - FCP, por UF origem/destino, mesmo nos casos de períodos sem movimento. Registro obrigatório,
'se existir o registro E300.
Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer, objSpedFiscal As ClassSpedFiscal
Dim indMov As Integer, objC101 As ClassSpedC101
Dim dVL_SLD_CRED_ANT_DIFAL As Double, dVL_TOT_DEBITOS_DIFAL As Double, dVL_OUT_DEB_DIFAL As Double, dVL_TOT_CREDITOS_DIFAL As Double, dVL_OUT_CRED_DIFAL As Double
Dim dVL_SLD_DEV_ANT_DIFAL As Double, dVL_DEDUCOES_DIFAL As Double, dVL_RECOL_DIFAL As Double, dVL_SLD_CRED_TRANSPORTAR_DIFAL As Double
Dim dDEB_ESP_DIFAL As Double, dVL_SLD_CRED_ANT_FCP As Double, dVL_TOT_DEB_FCP As Double, dVL_OUT_DEB_FCP As Double, dVL_TOT_CRED_FCP As Double
Dim dVL_OUT_CRED_FCP As Double, dVL_SLD_DEV_ANT_FCP As Double, dVL_DEDUCOES_FCP As Double, dVL_RECOL_FCP As Double, dVL_SLD_CRED_TRANSPORTAR_FCP As Double, dDEB_ESP_FCP As Double

On Error GoTo Erro_SpedFiscal_E310
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201223
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objFisSpedAux.objSpedRegistroE310.iCount = objFisSpedAux.objSpedRegistroE310.iCount + 1
       
    objSpedFiscal.sTipo = "E310"
    objFisSpedAux.objSpedRegistroE310.sTipo = objSpedFiscal.sTipo
    
    '01 REG Texto fixo contendo "E310"
    '02 IND_MOV_FCP_DIFAL Indicador de movimento: 0 – Sem operações 1 – Com operações
    '03 VL_SLD_CRED_ANT_DIFAL Valor do "Saldo credor de período anterior – ICMS Diferencial de Alíquota da UF de Origem/Destino"
    '04 VL_TOT_DEBITOS_DIFAL Valor total dos débitos por "Saídas e prestações com débito do ICMS referente ao diferencial de alíquota devido à UF de Origem/Destino"
    '05 VL_OUT_DEB_DIFAL Valor total dos ajustes "Outros débitos ICMS Diferencial de Alíquota da UF de Origem/Destino" e “Estorno de créditos ICMS Diferencial de Alíquota da UF de Origem/Destino”
    '06 VL_TOT_CREDITOS_DIFAL Valor total dos créditos do ICMS referente ao diferencial de alíquota devido à UF de Origem/Destino
    '07 VL_OUT_CRED_DIFAL Valor total de Ajustes "Outros créditos ICMS Diferencial de Alíquota da UF de Origem/Destino" e “Estorno de débitos ICMS Diferencial de Alíquota da UF de Origem/Destino”
    '08 VL_SLD_DEV_ANT_DIFAL Valor total de “Saldo devedor ICMS Diferencial de Alíquota da UF de Origem/Destino antes das deduções”
    '09 VL_DEDUÇÕES_DIFAL Valor total dos ajustes "Deduções ICMS Diferencial de Alíquota da UF de
    '10 VL_RECOL_DIFAL Valor recolhido ou a recolher referente ao ICMS Diferencial de Alíquota da UF de Origem/Destino (08-09)
    '11 VL_SLD_CRED_TRANSPORTAR_DIFAL Saldo credor a transportar para o período seguinte referente ao ICMS Diferencial de Alíquota da UF de Origem/Destino
    '12 DEB_ESP_DIFAL Valores recolhidos ou a recolher, extraapuração - ICMS Diferencial de Alíquota da UF de Origem/Destino.
    '13 VL_SLD_CRED_ANT_FCP Valor do "Saldo credor de período anterior – FCP"
    '14 VL_TOT_DEB_FCP Valor total dos débitos FCP por "Saídas e prestações”
    '15 VL_OUT_DEB_FCP Valor total dos ajustes "Outros débitos FCP " e “Estorno de créditos FCP”"
    '16 VL_TOT_CRED_FCP Valor total dos créditos FCP por Entradas N - 02 O
    '17 VL_OUT_CRED_FCP Valor total de Ajustes "Outros créditos FCP " e “Estorno de débitos FCP”"
    '18 VL_SLD_DEV_ANT_FCP Valor total de Saldo devedor FCP antes das deduções
    '19 VL_DEDUÇÕES_FCP Valor total das deduções "FCP" N - 02 O
    '20 VL_RECOL_FCP Valor recolhido ou a recolher referente ao FCP (18–19)
    '21 VL_SLD_CRED_TRANSPORTAR_FCP Saldo credor a transportar para o período seguinte referente ao FCP
    '22 DEB_ESP_FCP Valores recolhidos ou a recolher, extraapuração - FCP.
    indMov = 0
    dVL_SLD_CRED_ANT_DIFAL = 0
    dVL_OUT_DEB_DIFAL = 0
    dVL_OUT_CRED_DIFAL = 0
    For Each objC101 In objFisSpedAux.colC101
        If objC101.sUFDest = sUF Then
            indMov = 1
            If objC101.iCodSit <> 2 And objC101.iCodSit <> 7 Then
                If objC101.iIndOper = 1 Then
                    dVL_TOT_DEBITOS_DIFAL = dVL_TOT_DEBITOS_DIFAL + objC101.dVlrICMSDest
                    If objC101.sUFDest <> objC101.sUFRemet Then dVL_TOT_DEB_FCP = dVL_TOT_DEB_FCP + objC101.dVlrFCPDest
                Else
                    dVL_TOT_CREDITOS_DIFAL = dVL_TOT_CREDITOS_DIFAL + objC101.dVlrICMSDest
                    If objC101.sUFDest <> objC101.sUFRemet Then dVL_TOT_CRED_FCP = dVL_TOT_CRED_FCP + objC101.dVlrFCPDest
                End If
            Else
                If objC101.iIndOper = 1 Then
                    dDEB_ESP_DIFAL = dDEB_ESP_DIFAL + objC101.dVlrICMSDest
                    dDEB_ESP_FCP = dDEB_ESP_FCP + objC101.dVlrFCPDest
                End If
            End If
        End If
        If objC101.sUFRemet = sUF Then
            indMov = 1
            If objC101.iCodSit <> 2 And objC101.iCodSit <> 7 Then
                If objC101.iIndOper = 1 Then
                    dVL_TOT_DEBITOS_DIFAL = dVL_TOT_DEBITOS_DIFAL + objC101.dVlrICMSRemet
                Else
                    dVL_TOT_CREDITOS_DIFAL = dVL_TOT_CREDITOS_DIFAL + objC101.dVlrICMSRemet
                End If
            Else
                If objC101.iIndOper = 1 Then
                    dDEB_ESP_DIFAL = dDEB_ESP_DIFAL + objC101.dVlrICMSRemet
                End If
            End If
        End If
    Next
    
    dVL_SLD_DEV_ANT_DIFAL = (dVL_TOT_DEBITOS_DIFAL + dVL_OUT_DEB_DIFAL) - (dVL_SLD_CRED_ANT_DIFAL + dVL_TOT_CREDITOS_DIFAL + dVL_OUT_CRED_DIFAL)
    If dVL_SLD_DEV_ANT_DIFAL < 0 Then dVL_SLD_DEV_ANT_DIFAL = 0
    
    dVL_DEDUCOES_DIFAL = 0 'o valor informado deve corresponder ao somatório do campo VL_AJ_APUR do registro E311, por UF, quando o terceiro caractere for igual a ‘2’ e o quarto caractere do campo COD_AJ_APUR for igual a ‘4’.
    
    dVL_RECOL_DIFAL = dVL_SLD_DEV_ANT_DIFAL - dVL_DEDUCOES_DIFAL
    If dVL_RECOL_DIFAL < 0 Then dVL_RECOL_DIFAL = 0
    
    dVL_SLD_CRED_TRANSPORTAR_DIFAL = (dVL_SLD_CRED_ANT_DIFAL + dVL_TOT_CREDITOS_DIFAL + dVL_OUT_CRED_DIFAL) - (dVL_TOT_DEBITOS_DIFAL + dVL_OUT_DEB_DIFAL)
    If dVL_SLD_CRED_TRANSPORTAR_DIFAL < 0 Then dVL_SLD_CRED_TRANSPORTAR_DIFAL = 0
    
    dVL_SLD_CRED_ANT_FCP = 0 'Valor do campo VL_SLD_CRED_TRANSPORTAR_FCP do período de apuração anterior.

    dVL_OUT_DEB_FCP = 0 'o valor informado deve corresponder ao somatório do campo VL_AJ_APUR do registro E311, quando o terceiro caractere for igual a ‘3’ e o quarto for igual a ‘0’ ou ‘1’, ambos do campo COD_AJ_APUR do registro E311.

    dVL_OUT_CRED_FCP = 0 'o valor informado deve corresponder ao somatório do campo VL_AJ_APUR dos registros E311, quando o terceiro caractere for igual a ‘3’ e o quarto caractere do campo COD_AJ_APUR for igual a ‘2’ ou ‘3’.

    dVL_SLD_DEV_ANT_FCP = (dVL_TOT_DEB_FCP + dVL_OUT_DEB_FCP) - (dVL_SLD_CRED_ANT_FCP + dVL_TOT_CRED_FCP + dVL_OUT_CRED_FCP)
    If dVL_SLD_DEV_ANT_FCP < 0 Then dVL_SLD_DEV_ANT_FCP = 0
    
    dVL_DEDUCOES_FCP = 0 'o valor informado deve corresponder ao somatório do campo VL_AJ_APUR do registro E311, por UF, quando o terceiro caractere for igual a ‘3’ e o quarto caractere do campo COD_AJ_APUR for igual a ‘4’.

    dVL_RECOL_FCP = dVL_SLD_DEV_ANT_FCP - dVL_DEDUCOES_FCP
    If dVL_RECOL_FCP < 0 Then dVL_RECOL_FCP = 0
    
    dVL_SLD_CRED_TRANSPORTAR_FCP = (dVL_SLD_CRED_ANT_FCP + dVL_TOT_CRED_FCP + dVL_OUT_CRED_FCP) - (dVL_TOT_DEB_FCP + dVL_OUT_DEB_FCP)
    If dVL_SLD_CRED_TRANSPORTAR_FCP < 0 Then dVL_SLD_CRED_TRANSPORTAR_FCP = 0
    
    dDEB_ESP_FCP = 0
    
    objSpedFiscal.sRegistro = "|E310|" '01 REG
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & CStr(indMov) & "|" '02 IND_MOV_FCP_DIFAL
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_SLD_CRED_ANT_DIFAL) & "|" '03 VL_SLD_CRED_ANT_DIFAL
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_TOT_DEBITOS_DIFAL) & "|" '04 VL_TOT_DEBITOS_DIFAL
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_OUT_DEB_DIFAL) & "|" '05 VL_OUT_DEB_DIFAL
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_TOT_CREDITOS_DIFAL) & "|" '06 VL_TOT_CREDITOS_DIFAL
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_OUT_CRED_DIFAL) & "|" '07 VL_OUT_CRED_DIFAL
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_SLD_DEV_ANT_DIFAL) & "|" '08 VL_SLD_DEV_ANT_DIFAL
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_DEDUCOES_DIFAL) & "|" '09 VL_DEDUÇÕES_DIFAL
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_RECOL_DIFAL) & "|" '10 VL_RECOL_DIFAL
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_SLD_CRED_TRANSPORTAR_DIFAL) & "|" '11 VL_SLD_CRED_TRANSPORTAR_DIFAL
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dDEB_ESP_DIFAL) & "|" '12 DEB_ESP_DIFAL
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_SLD_CRED_ANT_FCP) & "|" '13 VL_SLD_CRED_ANT_FCP
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_TOT_DEB_FCP) & "|" '14 VL_TOT_DEB_FCP
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_OUT_DEB_FCP) & "|" '15 VL_OUT_DEB_FCP
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_TOT_CRED_FCP) & "|" '16 VL_TOT_CRED_FCP
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_OUT_CRED_FCP) & "|" '17 VL_OUT_CRED_FCP
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_SLD_DEV_ANT_FCP) & "|" '18 VL_SLD_DEV_ANT_FCP
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_DEDUCOES_FCP) & "|" '19 VL_DEDUÇÕES_FCP
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_RECOL_FCP) & "|" '20 VL_RECOL_FCP
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dVL_SLD_CRED_TRANSPORTAR_FCP) & "|" '21 VL_SLD_CRED_TRANSPORTAR_FCP
    objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Formata_Val2Dec(dDEB_ESP_FCP) & "|" '22 DEB_ESP_FCP

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 201224
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
    If lErro <> AD_SQL_SUCESSO Then gError 201225
    
    'lErro = SpedFiscal_E316(objFisSpedAux, objSpedFiscal.lSeq, vUF) NÃO TEM GUIA DE RECOLHIMENTO PARA ICMS INTERESTADUAL NO SISTEMA
    'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_E310 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_E310:

   SpedFiscal_E310 = gErr

   Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 201223
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201224
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201225
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204078)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function Desacentua_SpedFiscal(ByVal sTextoIn As String) As String
Dim i As Integer, sTextoOut As String, sAux As String

    sTextoOut = ""
    For i = 1 To Len(sTextoIn)
    
        sAux = Mid(sTextoIn, i, 1)
        Select Case sAux
        
            Case "§"
                sTextoOut = sTextoOut & "p"
            
            Case "|"
                sTextoOut = sTextoOut & "/"
            
            Case vbNewLine, vbCr, vbCrLf, vbLf
            
            Case Else
                sTextoOut = sTextoOut & sAux
                
        End Select
        
    Next

    Desacentua_SpedFiscal = SFF_Texto(sTextoOut)
    
End Function

Private Function Formata_Val2Dec(ByVal dValor As Double) As String
'formata valor com 2 decimais

    If dValor = 0 Then
        Formata_Val2Dec = "0"
    Else
        Formata_Val2Dec = Format(dValor, "Fixed")
    End If
    
End Function

Private Sub Formata_String_Numero2(ByVal sStringRecebe As String, sStringRetorna As String)
    sStringRetorna = ""
    Call Formata_String_Numero(sStringRecebe, sStringRetorna)
End Sub

Private Function SpedFiscal_G001(ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer, lSeqG001 As Long
Dim iIndMov As Integer, objSpedFiscal As ClassSpedFiscal

On Error GoTo Erro_SpedFiscal_G001
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201223
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    lSeqG001 = objSpedFiscal.lSeq
    objFisSpedAux.objSpedRegistroG001.iCount = objFisSpedAux.objSpedRegistroG001.iCount + 1

    lErro = SpedFiscal_G110(objFisSpedAux)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iIndMov = IIf(objFisSpedAux.objSpedRegistroG110.iCount = 0, 1, 0)
    
    objSpedFiscal.sTipo = "G001"
    objFisSpedAux.objSpedRegistroG001.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.sRegistro = "|G001|" & iIndMov & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 201224
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, lSeqG001, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 201225

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_G001 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_G001:

   SpedFiscal_G001 = gErr

   Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 201223
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201224
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201225
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204078)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_G990(ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lTotal As Long, objSpedFiscal As ClassSpedFiscal


On Error GoTo Erro_SpedFiscal_G990
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201226
    Next

    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
    
    objSpedFiscal.sTipo = "G990"
    objFisSpedAux.objSpedRegistroG990.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objFisSpedAux.objSpedRegistroG990.iCount = objFisSpedAux.objSpedRegistroG990.iCount + 1

    With objFisSpedAux
        lTotal = .objSpedRegistroG001.iCount + .objSpedRegistroG110.iCount + .objSpedRegistroG125.iCount + .objSpedRegistroG126.iCount + .objSpedRegistroG130.iCount + .objSpedRegistroG140.iCount + .objSpedRegistroG990.iCount
    End With

    objSpedFiscal.sRegistro = "|G990|" & lTotal & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 201227
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 201228


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_G990 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_G990:

   SpedFiscal_G990 = gErr

   Select Case gErr
    
        Case 201226
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201227
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201228
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204086)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_G110(ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 2) As Long, lSeq110 As Long
Dim iIndice As Integer, dSomaParc As Double, dSomaOutrosCreditos As Double
Dim iIndMov As Integer, dtDataIniApuracao As Date, dtDataFimApuracao As Date
Dim lCodG110 As Long, dValorTribExp As Double, dValorTotalSaidas As Double, dSaldoInicialICMS As Double
Dim objSpedFiscal As ClassSpedFiscal

On Error GoTo Erro_SpedFiscal_G110
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201210
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
    
    lErro = Comando_Executar(alComando(2), "SELECT Codigo, DataInicial, DataFinal, SaldoInicialICMS, ValorTribExp, ValorTotalSaidas FROM SpedFisRG110 WHERE FilialEmpresa = ? AND Ano = ? AND Mes = ? ORDER BY Codigo, DataInicial", _
        lCodG110, dtDataIniApuracao, dtDataFimApuracao, dSaldoInicialICMS, dValorTribExp, dValorTotalSaidas, objSpedFiscal.iFilialEmpresa, objFisSpedAux.iAno, objFisSpedAux.iMes)
    If lErro <> AD_SQL_SUCESSO Then gError 201205
        
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201206
    
    Do While lErro <> AD_SQL_SEM_DADOS
            
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        lSeq110 = objSpedFiscal.lSeq
        
        If lCodG110 = 0 Then lCodG110 = 1
        
        lErro = SpedFiscal_G125(objFisSpedAux, lCodG110, lSeq110)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dSomaParc = 0
        dSomaOutrosCreditos = 0
        
        objSpedFiscal.sTipo = "G110"
        objFisSpedAux.objSpedRegistroG110.sTipo = objSpedFiscal.sTipo
        objFisSpedAux.objSpedRegistroG110.iCount = objFisSpedAux.objSpedRegistroG110.iCount + 1
            
        objSpedFiscal.sRegistro = "|G110|" & Format(dtDataIniApuracao, "ddmmyyyy") & "|" & Format(dtDataFimApuracao, "ddmmyyyy") & "|" & Formata_Val2Dec(dSaldoInicialICMS) & "|" & Formata_Val2Dec(dSomaParc) & "|" & _
            Formata_Val2Dec(dValorTribExp) & "|" & Formata_Val2Dec(dValorTotalSaidas) & "|" & IIf(dValorTotalSaidas = 0 Or dValorTribExp = 0, "0", Format(dValorTribExp / dValorTotalSaidas, "0.00000000")) & "|" & _
            IIf(dSomaParc = 0 Or dValorTotalSaidas = 0 Or dValorTribExp = 0, "0", Formata_Val2Dec(dSomaParc * dValorTribExp / dValorTotalSaidas)) & "|" & Formata_Val2Dec(dSomaOutrosCreditos) & "|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 201207
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, lSeq110, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 201208

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201209
    
    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_G110 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_G110:

   SpedFiscal_G110 = gErr

   Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
                
        Case 201210
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201207
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201205, 201206, 201209
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_LEITURA_REGISTRO_BLOCO_G", gErr)
    
        Case 201208
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204078)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_0300(ByVal objFisSpedAux As ClassFISSpedAux, ByVal sCodigoBem As String) As Long

Dim lErro As Long, iIndice As Integer, alComando(1 To 3) As Long, lSeq300 As Long, objSpedFiscal As ClassSpedFiscal
Dim sDescricao As String, sTipo As String, sContaContabil As String, iNumParcelas As Integer, sCodigoBemPrincipal As String, sCcl As String, sFuncao As String, iVidaUtil As Integer
Dim lNumIntSpedAux As Long, sContaMascarada As String, sCclMascarada As String

On Error GoTo Erro_SpedFiscal_0300

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201241
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
        
    'verificar se já existe registro 0300 para este bem
    lErro = Comando_Executar(alComando(3), "SELECT NumIntSped FROM SpedFiscal WHERE NumIntSped = ? AND Registro LIKE ?", lNumIntSpedAux, objSpedFiscal.lNumIntSped, "|0300|" & sCodigoBem & "|%")
    If lErro <> AD_SQL_SUCESSO Then gError 201249
        
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201250
    
    If lErro <> AD_SQL_SUCESSO Then
    
        sDescricao = String(255, 0)
        sTipo = String(255, 0)
        sContaContabil = String(255, 0)
        sCodigoBemPrincipal = String(255, 0)
        sCcl = String(255, 0)
        sFuncao = String(255, 0)
        
        lErro = Comando_Executar(alComando(2), "SELECT Descricao, Tipo, ContaContabil, NumParcelas, CodigoBemPrincipal, Ccl, Funcao, VidaUtil FROM SpedFisRG03000305 WHERE FilialEmpresa = ? AND Ano = ? AND Mes = ? AND CodigoBem = ?", _
            sDescricao, sTipo, sContaContabil, iNumParcelas, sCodigoBemPrincipal, sCcl, sFuncao, iVidaUtil, objSpedFiscal.iFilialEmpresa, objFisSpedAux.iAno, objFisSpedAux.iMes, sCodigoBem)
        If lErro <> AD_SQL_SUCESSO Then gError 201242
            
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201243
        
        If lErro <> AD_SQL_SEM_DADOS Then
                
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            lSeq300 = objSpedFiscal.lSeq
            objSpedFiscal.sTipo = "0300"
            objFisSpedAux.objSpedRegistro0300.sTipo = objSpedFiscal.sTipo
            objFisSpedAux.objSpedRegistro0300.iCount = objFisSpedAux.objSpedRegistro0300.iCount + 1
            
            Call Mascara_MascararConta(sContaContabil, sContaMascarada)
           
            objSpedFiscal.sRegistro = "|0300|" & sCodigoBem & "|" & sTipo & "|" & SFF_Texto(sDescricao) & "|" & SFF_Texto(sCodigoBemPrincipal) & "|" & sContaMascarada & "|" & IIf(iNumParcelas = 0, "0", Format(iNumParcelas, "000")) & "|"
        
            If Len(objSpedFiscal.sRegistro) > 500 Then gError 201244
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, lSeq300, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 201245
    
            'criar reg 0500, se necessario
            If Len(Trim(sContaContabil)) <> 0 Then
                lErro = SpedFiscal_0500(objFisSpedAux, sContaContabil)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
            
            If sTipo = "1" Then
            
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                
                objSpedFiscal.sTipo = "0305"
                objFisSpedAux.objSpedRegistro0305.sTipo = objSpedFiscal.sTipo
                objFisSpedAux.objSpedRegistro0305.iCount = objFisSpedAux.objSpedRegistro0305.iCount + 1
                
                Call Mascara_MascararCcl(sCcl, sCclMascarada)
                
                objSpedFiscal.sRegistro = "|0305|" & sCclMascarada & "|" & SFF_Texto(sFuncao) & "|" & IIf(iVidaUtil = 0, "0", Format(iVidaUtil, "000")) & "|"
            
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 201246
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeq300)
                If lErro <> AD_SQL_SUCESSO Then gError 201247
        
                'criar reg 0600, se necessario
                If Len(Trim(sCcl)) <> 0 Then
                    lErro = SpedFiscal_0600(objFisSpedAux, sCcl)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                End If
                
            End If
            
        End If

    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    SpedFiscal_0300 = SUCESSO
    
    Exit Function
    
Erro_SpedFiscal_0300:

    SpedFiscal_0300 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 201242, 201243, 201248, 201249, 201250
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_LEITURA_REGISTRO_BLOCO_G", gErr)
    
        Case 201241
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201244, 201246
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201245, 201247
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201240)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function SpedFiscal_G125(ByVal objFisSpedAux As ClassFISSpedAux, ByVal lCodG110 As Long, ByVal lSeq110 As Long) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 2) As Long, objSpedFiscal As ClassSpedFiscal
Dim iIndice As Integer, lSeq125 As Long, sCodBemAnt As String
Dim sCodigoBem As String, sTipoMovto As String, iSeq As Integer, dtDataMovto As Date, dValorICMSEntrada As Double, dValorICMSSTEntrada As Double, dValorICMSFreteCT As Double, dValorICMSDifAliquota As Double, iNumParcela As Integer, dValorICMSPassAprop As Double

On Error GoTo Erro_SpedFiscal_G125
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201211
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
    
    sCodigoBem = String(STRING_CIAP_CODBEM, 0)
    sTipoMovto = String(STRING_CIAP_G125_TIPOMOV, 0)
    sCodBemAnt = ""
    
    lErro = Comando_Executar(alComando(2), "SELECT CodigoBem, TipoMovto, Seq, DataMovto, ValorICMSEntrada, ValorICMSSTEntrada, ValorICMSFreteCT, ValorICMSDifAliquota, NumParcela, ValorICMSPassAprop FROM SpedFisRG125 WHERE FilialEmpresa = ? AND Ano = ? AND Mes = ? AND (CodG110 = ? OR CodG110 = 0) ORDER BY CodigoBem, Seq", _
        sCodigoBem, sTipoMovto, iSeq, dtDataMovto, dValorICMSEntrada, dValorICMSSTEntrada, dValorICMSFreteCT, dValorICMSDifAliquota, iNumParcela, dValorICMSPassAprop, objSpedFiscal.iFilialEmpresa, objFisSpedAux.iAno, objFisSpedAux.iMes, lCodG110)
    If lErro <> AD_SQL_SUCESSO Then gError 201212
        
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201213
    
    Do While lErro <> AD_SQL_SEM_DADOS
            
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        lSeq125 = objSpedFiscal.lSeq
        
        If sCodBemAnt <> sCodigoBem Then
        
            lErro = SpedFiscal_0300(objFisSpedAux, sCodigoBem)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            sCodBemAnt = sCodigoBem
            
        End If
        
        lErro = SpedFiscal_G126(objFisSpedAux, lCodG110, sCodigoBem, sTipoMovto, lSeq125)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = SpedFiscal_G130(objFisSpedAux, lCodG110, sCodigoBem, sTipoMovto, lSeq125)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objSpedFiscal.sTipo = "G125"
        objFisSpedAux.objSpedRegistroG125.sTipo = objSpedFiscal.sTipo
        objFisSpedAux.objSpedRegistroG125.iCount = objFisSpedAux.objSpedRegistroG125.iCount + 1
        
        objSpedFiscal.sRegistro = "|G125|" & sCodigoBem & "|" & Format(dtDataMovto, "ddmmyyyy") & "|" & SFF_Texto(sTipoMovto) & "|" & Formata_Val2Dec(dValorICMSEntrada) & "|" & Formata_Val2Dec(dValorICMSSTEntrada) & "|" & Formata_Val2Dec(dValorICMSFreteCT) & "|" & Formata_Val2Dec(dValorICMSDifAliquota) & "|" & IIf(iNumParcela = 0, "0", Format(iNumParcela, "000")) & "|" & Formata_Val2Dec(dValorICMSPassAprop) & "|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 201214
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, lSeq125, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeq110)
        If lErro <> AD_SQL_SUCESSO Then gError 201215

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201216
    
    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_G125 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_G125:

   SpedFiscal_G125 = gErr

   Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 201212, 201213, 201216
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_LEITURA_REGISTRO_BLOCO_G", gErr)
    
        Case 201211
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201214
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201215
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204078)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_G126(ByVal objFisSpedAux As ClassFISSpedAux, ByVal lCodG110 As Long, ByVal sCodigoBem As String, ByVal sTipoMovto As String, ByVal lSeq125 As Long) As Long

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro, objSpedFiscal As ClassSpedFiscal
Dim alComando(1 To 2) As Long
Dim iIndice As Integer, dtDataInicial As Date, dtDataFinal As Date, iNumParcela As Integer, dValorICMSPassAprop As Double, dValorTribExp As Double, dValorTotal As Double, dValorICMSAprop As Double

On Error GoTo Erro_SpedFiscal_G126
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201217
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
    
    lErro = Comando_Executar(alComando(2), "SELECT DataInicial, DataFinal, NumParcela, ValorICMSPassAprop, ValorTribExp, ValorTotal, ValorICMSAprop FROM SpedFisRG126 WHERE FilialEmpresa = ? AND Ano = ? AND Mes = ? AND (CodG110 = ? OR CodG110 = 0) AND CodigoBem = ? AND TipoMovto = ? ORDER BY Seq", _
        dtDataInicial, dtDataFinal, iNumParcela, dValorICMSPassAprop, dValorTribExp, dValorTotal, dValorICMSAprop, objSpedFiscal.iFilialEmpresa, objFisSpedAux.iAno, objFisSpedAux.iMes, lCodG110, sCodigoBem, sTipoMovto)
    If lErro <> AD_SQL_SUCESSO Then gError 201218
        
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201219
    
    Do While lErro <> AD_SQL_SEM_DADOS
            
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        
        objSpedFiscal.sTipo = "G126"
        objFisSpedAux.objSpedRegistroG126.sTipo = objSpedFiscal.sTipo
        objFisSpedAux.objSpedRegistroG126.iCount = objFisSpedAux.objSpedRegistroG126.iCount + 1
        
        objSpedFiscal.sRegistro = "|G126|" & Format(dtDataInicial, "ddmmyyyy") & "|" & Format(dtDataFinal, "ddmmyyyy") & "|" & iNumParcela & "|" & Formata_Val2Dec(dValorICMSPassAprop) & "|" & Formata_Val2Dec(dValorTribExp) & "|" & Formata_Val2Dec(dValorTotal) & "|" & IIf(dValorTotal = 0 Or dValorTribExp = 0, "0", Format(dValorTribExp / dValorTotal, "0.00000000")) & "|" & Formata_Val2Dec(dValorICMSAprop) & "|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 201220
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeq125)
        If lErro <> AD_SQL_SUCESSO Then gError 201221

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201222
    
    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_G126 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_G126:

   SpedFiscal_G126 = gErr

   Select Case gErr
    
        Case 201218, 201219, 201222
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_LEITURA_REGISTRO_BLOCO_G", gErr)
    
        Case 201217
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201220
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201221
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204078)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_G130(ByVal objFisSpedAux As ClassFISSpedAux, ByVal lCodG110 As Long, ByVal sCodigoBem As String, ByVal sTipoMovto As String, ByVal lSeq125 As Long) As Long

Dim lErro As Long, lSeq130 As Long
Dim objSpedRegistro As ClassSpedRegistro, objSpedFiscal As ClassSpedFiscal
Dim alComando(1 To 2) As Long, iSeq As Integer, sCodParticipante As String
Dim iIndice As Integer, iEmitente As Integer, iTipoCliForn As Integer, lCliForn As Long, iFilial As Integer, iModelo As Integer, sSerie As String, lnumdoc As Long, dtDataEmissaoDoc As Date, schNFe As String, iNumItem As Integer, sProduto As String
Dim iEmitenteAnt As Integer, iTipoCliFornAnt As Integer, lCliFornAnt As Long, iFilialAnt As Integer, iModeloAnt As Integer, sSerieAnt As String, lNumDocAnt As Long, dtDataEmissaoDocAnt As Date, schNFeAnt As String
Dim objFilialCliente As ClassFilialCliente, objFilialForn As ClassFilialFornecedor, sErro As String

On Error GoTo Erro_SpedFiscal_G130
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201229
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
    
    sSerie = String(STRING_SERIE, 0)
    schNFe = String(STRING_NFE_CHNFE, 0)
    sProduto = String(STRING_PRODUTO, 0)
    
    lErro = Comando_Executar(alComando(2), "SELECT Emitente, TipoCliForn, CliForn, Filial, Modelo, Serie, NumDoc, DataEmissaoDoc, chNFe, NumItem, Produto FROM SpedFisRG130G140 WHERE FilialEmpresa = ? AND Ano = ? AND Mes = ? AND (CodG110 = ? OR CodG110 = 0) AND CodigoBem = ? AND TipoMovto = ? ORDER BY Emitente, TipoCliForn, CliForn, Filial, Modelo, Serie, NumDoc, DataEmissaoDoc, chNFe, NumItem", _
        iEmitente, iTipoCliForn, lCliForn, iFilial, iModelo, sSerie, lnumdoc, dtDataEmissaoDoc, schNFe, iNumItem, sProduto, objSpedFiscal.iFilialEmpresa, objFisSpedAux.iAno, objFisSpedAux.iMes, lCodG110, sCodigoBem, sTipoMovto)
    If lErro <> AD_SQL_SUCESSO Then gError 201230
        
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201231
    
    iEmitenteAnt = -1 'para forçar a quebra logo no 1o registro
    
    Do While lErro <> AD_SQL_SEM_DADOS
            
        If iEmitente <> iEmitenteAnt Or iTipoCliForn <> iTipoCliFornAnt Or lCliForn <> lCliFornAnt Or iFilial <> iFilialAnt Or iModelo <> iModeloAnt Or sSerie <> sSerieAnt Or lnumdoc <> lNumDocAnt Or dtDataEmissaoDoc <> dtDataEmissaoDocAnt Or schNFe <> schNFeAnt Then
        
            If iTipoCliForn = 1 Then
                sErro = "ERRO_FILIALCLIENTE_NAO_CADASTRADA"
            Else
                sErro = "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA"
            End If
            lErro = 1
            iIndice = 0
            Do While lErro <> SUCESSO And iIndice < 2
            
                iIndice = iIndice + 1
            
                If iTipoCliForn = 1 Then
                
                    'Se estiver marcado que é cliente mas não estiver cadastrado verifica se não é fornecedor
                    Set objFilialCliente = New ClassFilialCliente
                
                    objFilialCliente.lCodCliente = lCliForn
                    objFilialCliente.iCodFilial = iFilial
                            
                    'Le os Dados da Filial Cliente
                    lErro = CF("FilialCliente_Le", objFilialCliente)
                    If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
                    
                    If lErro <> SUCESSO Then iTipoCliForn = 2
                    
                Else
                    
                    Set objFilialForn = New ClassFilialFornecedor
                
                    objFilialForn.lCodFornecedor = lCliForn
                    objFilialForn.iCodFilial = iFilial
                
                    lErro = CF("FilialFornecedor_Le", objFilialForn)
                    If lErro <> SUCESSO And lErro <> 12929 Then gError ERRO_SEM_MENSAGEM
                    
                    If lErro <> SUCESSO Then iTipoCliForn = 1
                    
                End If
                
            Loop
            If lErro <> SUCESSO Then gError 216620
       
            If iTipoCliForn = 1 Then
            
                lErro = SpedFiscal_0150_Cliente(lCliForn, iFilial, objSpedFiscal, objFisSpedAux.objSpedRegistro0150, objFisSpedAux.objSpedRegistro0175, sCodParticipante, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            Else
    
                lErro = SpedFiscal_0150_Fornecedor(lCliForn, iFilial, objSpedFiscal, objFisSpedAux.objSpedRegistro0150, objFisSpedAux.objSpedRegistro0175, sCodParticipante, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            End If
            
            objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
            lSeq130 = objSpedFiscal.lSeq
            objSpedFiscal.sTipo = "G130"
            objFisSpedAux.objSpedRegistroG130.sTipo = objSpedFiscal.sTipo
            objFisSpedAux.objSpedRegistroG130.iCount = objFisSpedAux.objSpedRegistroG130.iCount + 1
            
            objSpedFiscal.sRegistro = "|G130|" & CStr(iEmitente) & "|" & sCodParticipante & "|" & Format(iModelo, "00") & "|" & sSerie & "|" & Format(lnumdoc, "000000000") & "|" & schNFe & "|" & Format(dtDataEmissaoDoc, "ddmmyyyy") & "|"
        
            If giVersao >= 14 Then objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "|"

            If Len(objSpedFiscal.sRegistro) > 500 Then gError 201234
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeq125)
            If lErro <> AD_SQL_SUCESSO Then gError 201235

            iEmitenteAnt = iEmitente
            iTipoCliFornAnt = iTipoCliForn
            lCliFornAnt = lCliForn
            iFilialAnt = iFilial
            iModeloAnt = iModelo
            sSerieAnt = sSerie
            lNumDocAnt = lnumdoc
            dtDataEmissaoDocAnt = dtDataEmissaoDoc
            schNFeAnt = schNFe
        
        End If
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedFiscal.sTipo = "G140"
        objFisSpedAux.objSpedRegistroG140.sTipo = objSpedFiscal.sTipo
        objFisSpedAux.objSpedRegistroG140.iCount = objFisSpedAux.objSpedRegistroG140.iCount + 1
        
        objSpedFiscal.sRegistro = "|G140|" & Format(iNumItem, "000") & "|" & SFF_Texto(sProduto) & "|"

        If giVersao >= 14 Then objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & "||||||"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 201236
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeq130)
        If lErro <> AD_SQL_SUCESSO Then gError 201237
        
        lErro = SpedFiscal_0200(sProduto, objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, 0, objFisSpedAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201239
    
    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_G130 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_G130:

    SpedFiscal_G130 = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 201230, 201231, 201239
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_LEITURA_REGISTRO_BLOCO_G", gErr)
    
        Case 201229
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201234, 201236
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201235, 201237
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
            
        Case 216620
            Call Rotina_Erro(vbOKOnly, sErro, gErr, iFilial, lCliForn)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204078)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SpedFiscal_0500(ByVal objFisSpedAux As ClassFISSpedAux, ByVal sContaContabil As String) As Long

Dim lErro As Long, iIndice As Integer, alComando(1 To 3) As Long, objSpedFiscal As ClassSpedFiscal
Dim tConta As typePlanoConta
Dim sTipo As String, sContaMascarada As String, dtDataAtualizacao As Date
Dim lNumIntSpedAux As Long

On Error GoTo Erro_SpedFiscal_0500

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201251
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
        
    Call Mascara_MascararConta(sContaContabil, sContaMascarada)
        
    'verificar se já existe registro 0500 para esta conta contabil
    lErro = Comando_Executar(alComando(3), "SELECT NumIntSped FROM SpedFiscal WHERE NumIntSped = ? AND Registro LIKE ?", lNumIntSpedAux, objSpedFiscal.lNumIntSped, "|0500|%|%|%|%|" & sContaMascarada & "|%")
    If lErro <> AD_SQL_SUCESSO Then gError 201252
        
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201253
    
    If lErro <> AD_SQL_SUCESSO Then
    
        tConta.sDescConta = String(STRING_CONTA_DESCRICAO, 0)
        
        lErro = Comando_Executar(alComando(1), "SELECT NivelConta, DescConta, TipoConta, NaturezaSped FROM PlanoConta WHERE Conta = ?", _
            tConta.iNivelConta, tConta.sDescConta, tConta.iTipoConta, tConta.iNaturezaSped, sContaContabil)
        If lErro <> AD_SQL_SUCESSO Then gError 201254
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201255
        If lErro <> AD_SQL_SUCESSO Then gError 201256
                
        If tConta.iNaturezaSped = 0 Then gError 201257 'Tem que estar preenchido
        
        If tConta.iTipoConta = CONTA_SINTETICA Then
            sTipo = "S"
        Else
            sTipo = "A"
        End If
        
        lErro = Comando_Executar(alComando(2), "SELECT MAX(DataAtualizacao) FROM PlanoContaHistorico WHERE Codigo = ? AND DataAtualizacao <= ?", _
                                                dtDataAtualizacao, sContaContabil, objFisSpedAux.dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 201258
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201259
        If lErro <> AD_SQL_SUCESSO Then dtDataAtualizacao = objFisSpedAux.dtDataIni
                
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedFiscal.sTipo = "0500"
        objFisSpedAux.objSpedRegistro0500.sTipo = objSpedFiscal.sTipo
        objFisSpedAux.objSpedRegistro0500.iCount = objFisSpedAux.objSpedRegistro0500.iCount + 1
        
        objSpedFiscal.sRegistro = "|0500|" & Format(dtDataAtualizacao, "ddmmyyyy") & "|" & Format(tConta.iNaturezaSped, "00") & "|" & sTipo & "|" & Format(tConta.iNivelConta, "00000") & "|" & sContaMascarada & "|" & left(tConta.sDescConta, 60) & "|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 201260
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, 0)
        If lErro <> AD_SQL_SUCESSO Then gError 201261

    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    SpedFiscal_0500 = SUCESSO
    
    Exit Function
    
Erro_SpedFiscal_0500:

    SpedFiscal_0500 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 201252, 201253, 201254, 201255, 201258, 201259
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_LEITURA_REGISTRO_BLOCO_G", gErr)
    
        Case 201256
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_CTACONTABIL_INEXISTENTE", gErr, sContaContabil)
        
        Case 201257
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_CTACONTABIL_SEM_NATUREZA", gErr, sContaContabil)
        
        Case 201251
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201260
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201261
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201240)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function SpedFiscal_0600(ByVal objFisSpedAux As ClassFISSpedAux, ByVal sCcl As String) As Long

Dim lErro As Long, iIndice As Integer, alComando(1 To 3) As Long, objSpedFiscal As ClassSpedFiscal
Dim tCcl As typeCcl
Dim sCclMascarado As String, dtDataAtualizacao As Date, lNumIntSpedAux As Long

On Error GoTo Erro_SpedFiscal_0600

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201263
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
        
    If giSetupUsoCcl = CCL_NAO_USA Then
        
        sCclMascarado = sCcl
        dtDataAtualizacao = objFisSpedAux.dtDataIni
        
        Select Case sCcl
            Case "1"
                tCcl.sDescCcl = "área operacional"
            Case "2", "5"
                tCcl.sDescCcl = "área administrativa"
            Case "3"
                tCcl.sDescCcl = "área produtiva"
            Case "4"
                tCcl.sDescCcl = "área de apoio à produção"
        End Select
        
    Else
    
        Call Mascara_MascararCcl(sCcl, sCclMascarado)
        
    End If
        
    'verificar se já existe registro 0600 para este centro de custo
    lErro = Comando_Executar(alComando(3), "SELECT NumIntSped FROM SpedFiscal WHERE NumIntSped = ? AND Registro LIKE ?", lNumIntSpedAux, objSpedFiscal.lNumIntSped, "|0600|%|" & sCclMascarado & "|%")
    If lErro <> AD_SQL_SUCESSO Then gError 201264
        
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201265
    
    If lErro <> AD_SQL_SUCESSO Then
    
        If giSetupUsoCcl <> CCL_NAO_USA Then
        
            tCcl.sDescCcl = String(STRING_CCL_DESCRICAO, 0)
            
            lErro = Comando_Executar(alComando(1), "SELECT DescCcl FROM Ccl WHERE Ccl = ?", _
                tCcl.sDescCcl, sCcl)
            If lErro <> AD_SQL_SUCESSO Then gError 201266
        
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201267
            If lErro <> AD_SQL_SUCESSO Then gError 201268
                    
            lErro = Comando_Executar(alComando(2), "SELECT MAX(DataAtualizacao) FROM CclHistorico WHERE CodigoCcl = ? AND DataAtualizacao <= ?", _
                                                    dtDataAtualizacao, sCcl, objFisSpedAux.dtDataFim)
            If lErro <> AD_SQL_SUCESSO Then gError 201269
        
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201270
            If lErro <> AD_SQL_SUCESSO Then dtDataAtualizacao = objFisSpedAux.dtDataIni
                
        End If
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objSpedFiscal.sTipo = "0600"
        objFisSpedAux.objSpedRegistro0600.sTipo = objSpedFiscal.sTipo
        objFisSpedAux.objSpedRegistro0600.iCount = objFisSpedAux.objSpedRegistro0600.iCount + 1
        
        objSpedFiscal.sRegistro = "|0600|" & Format(dtDataAtualizacao, "ddmmyyyy") & "|" & sCclMascarado & "|" & SFF_Texto(tCcl.sDescCcl, 60) & "|"
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 201271
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
            objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, 0)
        If lErro <> AD_SQL_SUCESSO Then gError 201272

    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    SpedFiscal_0600 = SUCESSO
    
    Exit Function
    
Erro_SpedFiscal_0600:

    SpedFiscal_0600 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 201264, 201265, 201266, 201267, 201269, 201270
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_LEITURA_REGISTRO_BLOCO_G", gErr)
    
        Case 201268
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_CCL_INEXISTENTE", gErr, sCcl)
        
        Case 201263
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 201271
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 201272
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201262)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Ajusta_Serie(sSerie As String) As String
    'Não pode ter mais de 3 caracteres
    sSerie = Trim(sSerie)
    If Len(sSerie) > 3 Then
        sSerie = Replace(Replace(Replace(Replace(Replace(sSerie, ".", ""), "-", ""), "_", ""), "/", ""), "\", "")
    End If
    If Len(sSerie) > 3 Then
        sSerie = left(sSerie, 3)
    End If
    Ajusta_Serie = sSerie
End Function

Private Function NF_Le_BasePropria(ByVal objTrib As ClassTributacaoDoc) As Long
'Retorna zerando os dados de ICMS nos casos específicos (normalmente em entradas sem crédito de ICMS)

Dim lErro As Long, alComando(1 To 1) As Long
Dim lNumIntDocItem As Long, iIndice As Integer
Dim objTribItem As ClassTributacaoDocItem
Dim dICMSBase As Double, dICMSValor As Double, dICMSCredito As Double
Dim dIPIBase As Double, dIPIValor As Double, dIPICredito As Double
Dim iTipoTribCST_ICMS As Integer, iIPICSTEntrada As Integer, iIPICSTSaida As Integer
Dim dICMSSTBase As Double, dICMSSTValor As Double
Dim iTipoTribCST_ICMSAux As Integer, iOrigemMercadoria As Integer

On Error GoTo Erro_NF_Le_BasePropria

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211138
    Next

    lErro = Comando_Executar(alComando(1), "SELECT OrigemMercadoria, NumIntDocItem, TipoTribCST, IPI.CSTEntrada, IPI.CSTSaida FROM TributacaoDocItem, tiposTribICMS AS ICMS, TiposTribIPI AS IPI  WHERE TipoDoc = ? AND NumIntDoc  = ? AND ICMS.Tipo = ICMSTipo AND IPI.Tipo = IPITipo", _
    iOrigemMercadoria, lNumIntDocItem, iTipoTribCST_ICMSAux, iIPICSTEntrada, iIPICSTSaida, objTrib.iTipoDoc, objTrib.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 211139
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211140

    Do While lErro <> AD_SQL_SEM_DADOS
    
        iTipoTribCST_ICMS = iTipoTribCST_ICMSAux
    
        Set objTribItem = New ClassTributacaoDocItem
        objTribItem.lNumIntDocItem = lNumIntDocItem
        objTribItem.iTipoDoc = objTrib.iTipoDoc
        lErro = ItemNF_Le_BasePropria(objTribItem, iTipoTribCST_ICMS, iIPICSTEntrada)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        iTipoTribCST_ICMS = iTipoTribCST_ICMS + (100 * iOrigemMercadoria)
        
        dICMSBase = dICMSBase + (objTribItem.dICMSBase * (1 - objTribItem.dICMSPercRedBase))
        dICMSValor = dICMSValor + objTribItem.dICMSValor
        dICMSCredito = dICMSCredito + objTribItem.dICMSCredito
        
        dICMSSTBase = dICMSSTBase + objTribItem.dICMSSubstBase
        dICMSSTValor = dICMSSTValor + objTribItem.dICMSSubstValor

        dIPIBase = dIPIBase + objTribItem.dIPIBaseCalculo
        dIPIValor = dIPIValor + objTribItem.dIPIValor
        dIPICredito = dIPICredito + objTribItem.dIPICredito
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211141
    
    Loop
    
    objTrib.dICMSBase = dICMSBase
    objTrib.dICMSValor = dICMSValor
    objTrib.dICMSCredito = dICMSCredito
    
    objTrib.dICMSSubstBase = dICMSSTBase
    objTrib.dICMSSubstValor = dICMSSTValor
            
    objTrib.dIPIBase = dIPIBase
    objTrib.dIPIValor = dIPIValor
    objTrib.dIPICredito = dIPICredito
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    NF_Le_BasePropria = SUCESSO
    
    Exit Function
    
Erro_NF_Le_BasePropria:

    NF_Le_BasePropria = gErr

    Select Case gErr
    
        Case 211138
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 211139 To 211141
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAODOCITEM", gErr)
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211142)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function ItemNF_Le_BasePropria(ByVal objTribItem As ClassTributacaoDocItem, iTipoTribCST_ICMS As Integer, iIPICSTEntrada As Integer) As Long
'iCST_ICMS e iCST_IPI tem que estarem preenchidos
'Retorna zerando os dados de ICMS nos casos específicos (normalmente em entradas sem crédito de ICMS)
Dim lErro As Long

On Error GoTo Erro_ItemNF_Le_BasePropria

    lErro = CF("TributacaoDocItem_Le", objTribItem)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If iTipoTribCST_ICMS = 0 And objTribItem.dICMSValor <> 0 And objTribItem.dICMSCredito = 0 Then
        iTipoTribCST_ICMS = 90
    End If

    If iTipoTribCST_ICMS = 90 And objTribItem.dICMSCredito = 0 Then
       
        objTribItem.dICMSBase = 0
        objTribItem.dICMSAliquota = 0
        objTribItem.dICMSValor = 0
        objTribItem.dICMSCredito = 0
        
        objTribItem.dICMSSubstBase = 0
        objTribItem.dICMSSubstAliquota = 0
        objTribItem.dICMSSubstValor = 0
        objTribItem.dICMSSubstPercMVA = 0
        objTribItem.dICMSSubstPercRedBase = 0
        
    End If
    
    'Se é entrada e tem ST
    If (left(objTribItem.sNaturezaOp, 1) = "1" Or left(objTribItem.sNaturezaOp, 1) = "2") And objTribItem.dICMSSubstValor > DELTA_VALORMONETARIO Then
        'Se não é devolução
        If objTribItem.sNaturezaOp <> "1410" And objTribItem.sNaturezaOp <> "1411" And objTribItem.sNaturezaOp <> "1660" And objTribItem.sNaturezaOp <> "1661" And objTribItem.sNaturezaOp <> "1662" And _
            objTribItem.sNaturezaOp <> "2410" And objTribItem.sNaturezaOp <> "2411" And objTribItem.sNaturezaOp <> "2660" And objTribItem.sNaturezaOp <> "2661" And objTribItem.sNaturezaOp <> "2662" Then
            
            objTribItem.dICMSBase = 0
            objTribItem.dICMSAliquota = 0
            objTribItem.dICMSValor = 0
            objTribItem.dICMSCredito = 0
            
            'Entrada sem ser devolução não vai gerar créditos de ST
            objTribItem.dICMSSubstBase = 0
            objTribItem.dICMSSubstAliquota = 0
            objTribItem.dICMSSubstValor = 0
            objTribItem.dICMSSubstPercMVA = 0
            objTribItem.dICMSSubstPercRedBase = 0
            'C170, C190, E220 -> Repercute nesses registros
            'EM ValorOper em C190 o ST está sendo considerado
        End If
    End If

    If iIPICSTEntrada = 0 And objTribItem.dIPIValor <> 0 And objTribItem.dIPICredito = 0 Then
        iIPICSTEntrada = 49
    End If

    If iIPICSTEntrada = 49 And objTribItem.dIPICredito = 0 Then
        objTribItem.dIPIBaseCalculo = 0
        objTribItem.dIPIAliquota = 0
        objTribItem.dIPIValor = 0
        objTribItem.dIPICredito = 0
    End If

    ItemNF_Le_BasePropria = SUCESSO
    
    Exit Function
    
Erro_ItemNF_Le_BasePropria:

    ItemNF_Le_BasePropria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211137)

    End Select
    
    Exit Function
    
End Function

Private Function NF_Obter_Situacao(ByVal lComando As Long, tNFiscal As typeNFiscal, tTipoDocInfo As typeTipoDocInfo, iCodSit As Integer) As Long

Dim lErro As Long
Dim iAux As Integer

On Error GoTo Erro_NF_Obter_Situacao
    '
    'Código da Situação do Documento Descrição da Situação do Documento  Data de Início  Data de Fim
    '00  Documento regular   01012009
    '01  Documento regular extemporâneo  01012009
    '02  Documento cancelado 01012009
    '03  Documento cancelado extemporâneo    01012009
    '04  NFe ou CT-e denegada    01012009
    '05  NFe ou CT-e Numeração inutilizada   01012009
    '06  Documento Fiscal Complementar   01012009
    '07  Documento Fiscal Complementar extemporâneo  01012009
    '08  Documento Fiscal emitido com base em Regime Especial ou Norma Específica    01012009

    iCodSit = 0 'Documento regular
    If tNFiscal.iStatus = STATUS_CANCELADO Then
        iCodSit = 2 'Documento cancelado
        'Busca Denegado senão vai ser cancelada mesmo
        If ISSerieEletronica(tNFiscal.sSerie) Then

            'se for uma nota eletronica denegada no ambiente de producao
            lErro = Comando_Executar(lComando, "SELECT tpAmb FROM NFeFedProtNFe WHERE NumIntNF = ? AND cStat IN ('110', '301','302') AND tpAmb = 0", iAux, tNFiscal.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 209381
        
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209382
        
            'nota fiscal denegada
            If lErro = AD_SQL_SUCESSO Then iCodSit = 4 'NFe ou CT-e denegada
                
        End If
    Else
        If tTipoDocInfo.iComplementar = MARCADO Then iCodSit = 6 'Documento Fiscal Complementar
    End If
            
    NF_Obter_Situacao = SUCESSO
    
    Exit Function
    
Erro_NF_Obter_Situacao:

    NF_Obter_Situacao = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211137)

    End Select
    
    Exit Function
    
End Function

Private Function SpedFiscal_Obtem_DadosInicias(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim lNumIntSped As Long
Dim iCodVersao As Integer
Dim lTransacao As Long
Dim iIndice As Integer
Dim objFilialEmpresa As New AdmFiliais, lContador As Long

On Error GoTo Erro_SpedFiscal_Obtem_DadosInicias

    'Abertura de Comando
    For iIndice = 0 To objEFD.NUM_MAX_COMANDOS
        objEFD.alComando(iIndice) = Comando_Abrir()
        If objEFD.alComando(iIndice) = 0 Then gError 209366
    Next
    
'    'Obtem o total de registros que vão atualizar a barra de progressão
'    lErro = Registros_Totaliza(objEFD)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    Call objEFD.objTela.Inicia_Processo(objEFD.lNumTotalPrevisto)
'
    lErro = CF("Config_ObterNumInt", "FISConfig", "NUM_PROX_SPEDFISCALPIS", lNumIntSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    'Obtém a versão do layout
'    lErro = Comando_Executar(objEFD.ProxComando, "SELECT Codigo From SpedFisPisVersaoLeiaute WHERE DataInicio <= ? ORDER BY DataInicio DESC ", iCodVersao, objEFD.dtDataFim)
'    If lErro <> AD_SQL_SUCESSO Then gError 209367
'
'    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209368
'
'    'versao nao cadastrada
'    If lErro <> AD_SQL_SUCESSO Then gError 209369
'
'    objFilialEmpresa.iCodFilial = giFilialEmpresa 'FILIAL_MATRIZ 'iFilialEmpresa
    objFilialEmpresa.iCodFilial = objEFD.iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
    
    objEFD.lNumIntSped = lNumIntSped
    'objEFD.iCodVersao = iCodVersao
    
    objEFD.objSpedFiscal.iFilialEmpresa = objEFD.iFilialEmpresa
    objEFD.objSpedFiscal.lNumIntSped = objEFD.lNumIntSped
    
    Set objEFD.objFilialEmpresa = objFilialEmpresa
    
'    If objEFD.iCodVersao < 3 Or objEFD.objFilialEmpresa.iLucroPresumido <> MARCADO Or objEFD.objFilialEmpresa.iCOFINSNaoCumulativo = MARCADO Or objEFD.objFilialEmpresa.iPisCofinsRegCumTipo <> 1 Then
'        objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA
'    Else
'        objEFD.iRegimeApuracao = objEFD.TRATA_REG_CAIXA
'    End If
    
    'Grava SpedFiscalDocs e SpedFiscalDocItens
    lErro = SpedFiscal_Grava_Tabelas(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    SpedFiscal_Obtem_DadosInicias = SUCESSO

    Exit Function
    
Erro_SpedFiscal_Obtem_DadosInicias:

   SpedFiscal_Obtem_DadosInicias = gErr

   Select Case gErr
   
        Case 209366
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 209367, 209368
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISPISVERSAOLEIAUTE", gErr)
            
        Case 209369
            Call Rotina_Erro(vbOKOnly, "ERRO_EFD_PIS_COFINS_SEM_VERSAO", gErr, Format(objEFD.dtDataFim, "dd/mm/yyyy"))
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209370)
     
    End Select
    
End Function

'Private Function Registros_Totaliza(ByVal objEFD As ClassEFDPisCofinsSel) As Long
''Verifica os registros com grande processamento para ajustar a barra de progressão
'
'Dim lErro As Long, iAno As Integer
'Dim lContadorItensNF As Long
'Dim sSQLNF As String
'
'On Error GoTo Erro_Registros_Totaliza
'
'    Call objEFD.objTela.SPED_Status("Totalizando Registros", "NFiscal")
'
'    sSQLNF = "SELECT COUNT (*) "
'    sSQLNF = sSQLNF & "FROM NFiscal AS N, TiposDocInfo AS T, TributacaoDoc AS TD, ItensNFiscal AS I "
'    sSQLNF = sSQLNF & "WHERE I.NumIntNF = N.NumIntDoc AND N.TipoNFiscal = T.Codigo AND TD.TipoDoc = 0 AND TD.NumIntDoc = N.NumIntDoc AND "
'    sSQLNF = sSQLNF & "((T.Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) Or ( "
'    sSQLNF = sSQLNF & "(T.Tipo = 1 Or T.Tipo = 3) AND DataEntrada >= ? AND DataEntrada <= ?)) AND N.FilialEmpresa < 50 "
'
'    lErro = Comando_Executar(objEFD.ProxComando, sSQLNF, lContadorItensNF, objEFD.dtDataIni, objEFD.dtDataFim, objEFD.dtDataIni, objEFD.dtDataFim)
'    If lErro <> AD_SQL_SUCESSO Then gError 209371
'
'    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209372
'
'    objEFD.lNumTotalItensNF = lContadorItensNF
'
'    'É uma previsão mais ou menos que vai ser ajustada no decorrer do processo
'    objEFD.lNumTotalPrevisto = 3 * lContadorItensNF '(Ler, Processar e Gravar)
'
'    Registros_Totaliza = SUCESSO
'
'    Exit Function
'
'Erro_Registros_Totaliza:
'
'   Registros_Totaliza = gErr
'
'   Select Case gErr
'
'        Case 209371, 209372
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209373)
'
'    End Select
'
'End Function

Private Sub Registro_Adiciona_Count(ByVal sTipo As String, ByVal colSpedRegistro As Collection, ByVal iCount As Integer)

Dim objSpedRegistro As New ClassSpedRegistro
Dim bAchou As Boolean

    If iCount > 0 Then

        bAchou = False
        
        For Each objSpedRegistro In colSpedRegistro
            If objSpedRegistro.sTipo = sTipo Then
                objSpedRegistro.lCount = objSpedRegistro.lCount + iCount
                bAchou = True
                Exit For
            End If
        Next
        
        If bAchou = False Then
        
            Set objSpedRegistro = New ClassSpedRegistro
        
            objSpedRegistro.sTipo = sTipo
            objSpedRegistro.lCount = iCount
        
            colSpedRegistro.Add objSpedRegistro
        
        End If
        
    End If

End Sub

Private Function Registro_Insere_BD(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim objSpedFiscal As ClassSpedFiscal
Dim sAux As String, iSeq As Long, sReg As String, sErro As String

On Error GoTo Erro_Registro_Insere_BD

    sErro = "01"
    Set objSpedFiscal = objEFD.objSpedFiscal

    sErro = "02"
    If left(objSpedFiscal.sTipo, 2) <> "99" Then Call Registro_Adiciona_Count(objEFD.objSpedFiscal.sTipo, objEFD.colSpedRegistro, 1)

    sErro = "03"
    lErro = Comando_Executar(objEFD.lComandoInsert, "INSERT INTO SpedFiscalPis( NumIntSped, Seq, FilialEmpresa, Tipo,  SeqPai ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.lSeqPai)
    sErro = "04"
    If lErro <> AD_SQL_SUCESSO Then gError 209374
    
    sErro = "05"
    sAux = objSpedFiscal.sRegistro

    sErro = "06"
    iSeq = 0
    sErro = "07"
    Do While Len(Trim(sAux)) > 0
    
    sErro = "08"
        iSeq = iSeq + 1
        
    sErro = "09"
        sReg = left(sAux, STRING_MAXIMO)
    
    sErro = "10"
        If Len(Trim(sAux)) > STRING_MAXIMO Then
    sErro = "11"
            sAux = Mid(sAux, STRING_MAXIMO + 1)
        Else
    sErro = "12"
            sAux = ""
        End If
        
    sErro = "13"
        lErro = Comando_Executar(objEFD.lComandoInsertReg, "INSERT INTO SpedFiscalPisReg(NumIntSped, Seq, SeqReg, Registro ) VALUES (?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, iSeq, sReg)
    sErro = "14"
        If lErro <> AD_SQL_SUCESSO Then gError 209375
    
    sErro = "15"
    Loop
    
    'sErro = "16"
    'lErro = objEFD.objTela.Processa_Item(True)
    'sErro = "17"
    'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sErro = "18"
    objEFD.lNumTotalProcessados = objEFD.lNumTotalProcessados + 1 'Total de processamentos (3 por registro)
    sErro = "19"
    objEFD.lNumTotalReal = objEFD.lNumTotalReal + 1 'Total real de registros processados
    sErro = "20"

    Registro_Insere_BD = SUCESSO

    Exit Function

Erro_Registro_Insere_BD:

   Registro_Insere_BD = gErr

   Select Case gErr

        Case 209374, 209375
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209376)

    End Select

    Call Rotina_Erro(vbOKOnly, "Localização do Erro: " & sErro, gErr)

End Function

Private Function SpedFiscal_Grava_Tabelas(ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 22) As Long, lComandoAux As Long
Dim tNFiscal As typeNFiscal, tTipoDocInfo As typeTipoDocInfo, tTributacaoDoc As typeTributacaoDoc
Dim tItemNF As typeItemNF, tTribDocItem As typeTributacaoDocItem, iIPICSTE As Integer, iIPICSTS As Integer, iICMSIncide As Integer, iISSIncide As Integer
Dim sSQLNF As String, sSQLItemNF As String
Dim objDoc As ClassSpedDocFiscais, objItem As ClassSpedDocFiscaisItens
Dim bNaoIncideICMS As Boolean, iNatureza As Integer, iCondPagto As Integer
Dim sCodVerificacaoNFe As String, sCgc As String, sChave As String
Dim tNFOrig As typeNFiscal, tTribDocOrig As typeTributacaoDoc, bExisteNFOrig As Boolean, bNFOrigPropEmp As Boolean
Dim lNumNFInut As Long, lNumNFDeInut As Long, lNumNFAteInut As Long
Dim sSerieNFInut As String, dtDataNFInut As Date, iFilialEmpNFInut As Integer, sCGCNFInut As String
Dim sNumDe As String, sNumAte As String
Dim dValorRecebido As Double, sSQLCR As String, dValorNoCR As Double
Dim iPasso As Integer, sCodAtividadeTab511 As String, dAliquotaContrPrev As Double
Dim sSQLNFBase, sSQLNFRec As String

On Error GoTo Erro_SpedFiscal_Grava_Tabelas

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next
   
    sSQLNFBase = "SELECT F.CGC, N.FilialEmpresa, N.NumIntDoc, N.RPS, N.CodVerificacaoNFe, N.ChvNFe, N.NFe, N.TipoNFiscal, T.Tipo, N.Cliente, N.FilialCli, "
    sSQLNFBase = sSQLNFBase & "N.Fornecedor, N.FilialForn, N.Status, Complementar, N.Serie, N.NumNotaFiscal, "
    sSQLNFBase = sSQLNFBase & "N.DataEmissao , N.DataEntrada, N.DataSaida, N.ValorTotal, N.ClasseDocCPR, N.NumIntDocCPR, N.ValorDesconto, "
    sSQLNFBase = sSQLNFBase & "N.ValorProdutos , N.ValorFrete, N.FreteRespons, N.ValorSeguro, N.ValorOutrasDespesas, "
    sSQLNFBase = sSQLNFBase & "N.FilialEntrega , N.CodTransportadora, N.Placa, N.PlacaUF, N.VolumeQuant, "
    sSQLNFBase = sSQLNFBase & "N.PesoLiq , N.PesoBruto, T.TipoMovtoEstoque, T.NomeTelaNFiscal, "
    sSQLNFBase = sSQLNFBase & "TD.TipoTributacao , TD.IPIBase, TD.IPIValor, TD.ICMSBase, TD.ICMSValor, "
    sSQLNFBase = sSQLNFBase & "TD.ICMSSubstBase, TD.ICMSSubstValor, TD.ISSBase, TD.ISSValor, TD.PisRetido, TD.CofinsRetido, "
    sSQLNFBase = sSQLNFBase & "TD.NaturezaOPInterna, N.NumIntNotaOriginal, T.Faturamento,  TD.ISSRetido "
    sSQLNFBase = sSQLNFBase & "FROM NFiscal AS N, TiposDocInfo AS T, TributacaoDoc AS TD, FiliaisEmpresa F "
    sSQLNFBase = sSQLNFBase & "WHERE F.FilialEmpresa = N.FilialEmpresa AND N.TipoNFiscal = T.Codigo AND TD.TipoDoc = 0 AND TD.NumIntDoc = N.NumIntDoc AND F.CGC LIKE ? "
   
    'For iPasso = 1 To 2
    For iPasso = 1 To 1
    
'        If objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA And iPasso = 1 Then
        
            'REGIME DE COMPETÊNCIA
            'NOTAS FISCAIS COM ENTRADA E SAÍDA NO PERÍODO ... BLOCOS A, C e D
            sSQLNF = sSQLNFBase & " AND "
'            sSQLNF = sSQLNF & "(TD.PisCredito <> 0 OR TD.CofinsCredito <> 0 OR T.Faturamento <> 0) AND "
'            sSQLNF = sSQLNF & "((Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) Or ( "
'            sSQLNF = sSQLNF & "(Tipo = 1 Or Tipo = 3) AND DataEntrada >= ? AND DataEntrada <= ?)) AND F.FilialEmpresa < 50 "
            sSQLNF = sSQLNF & "(T.Faturamento <> 0 AND TD.ISSValor > 0) AND ((Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) OR ((Tipo = 1 Or Tipo = 3) AND DataEntrada >= ? AND DataEntrada <= ?)) AND F.FilialEmpresa < 50 "

            sSQLNF = sSQLNF & " ORDER BY N.NumIntDoc" 'para garantir que reexecute passando pelos registros na mesma ordem...
        
'        ElseIf (objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA And iPasso = 2) Or (objEFD.iRegimeApuracao = objEFD.TRATA_REG_CAIXA And iPasso = 1) Then
'
'            'NOTAS FISCAIS PAGAS NO PERÍODO ... BLOCO F PARA REGIME DE CAIXA E PODE SER USADO PARA O REGIME DE COMPETÊNCIA PARA CONTRIBUIÇÃO PREVIDENCIÁRIA POR RECEITA BRUTA
'            sSQLNF = sSQLNFBase & " AND "
'            sSQLNF = sSQLNF & "(TD.PisCredito <> 0 OR TD.CofinsCredito <> 0 OR T.Faturamento <> 0) AND F.FilialEmpresa < 50 AND "
'
'            'Pega as NFs se existir um recebimento no período
'            'Se for adiantamento pega pela data de faturamento senão pela baixa (exceto baixa por perda)
'            sSQLNF = sSQLNF & "N.NumIntDoc IN ( "
'            sSQLNF = sSQLNF & "SELECT N.NumIntDoc "
'            sSQLNF = sSQLNF & "FROM NFiscal AS N, TitulosRecTodos AS TR, ParcelasRecTodas As P, BaixasRec AS B, BaixasParcRec As BP "
'            sSQLNF = sSQLNF & "WHERE N.Status <> 7 AND N.ClasseDocCPR = 2 AND N.NumIntDocCPR = TR.NumIntDoc AND TR.Status <> 5 AND "
'            sSQLNF = sSQLNF & "TR.NumIntDoc = P.NumIntTitulo AND BP.NumIntParcela = P.NumIntDoc AND B.NumIntBaixa = BP.NumIntBaixa AND "
'            sSQLNF = sSQLNF & "B.Status <> 5 AND BP.Status <> 5 AND B.Motivo NOT IN (4,6) AND ((B.Motivo <> 5 AND B.Data BETWEEN ? AND ?) OR (B.Motivo = 5 AND N.DataEmissao BETWEEN ? AND ?))) "
'
'        ElseIf objEFD.iRegimeApuracao = objEFD.TRATA_REG_CAIXA And iPasso = 2 Then
'
'            'REGIME DE CAIXA
'            'NOTAS FISCAIS GERADORAS DE RECEITA EMITIDAS NO PERÍODO .. REGISTRO 1900
'
'            'Além das NFs que geraram receita no período e vão gerar os registros F500 será necessário
'            'informar as NFs emitidas no próprio período para o registro 1900
'            'Para o regime de competência não vai gerar o 1900 porque o arquivo é detalhado nos blocos A, C e D e
'            'não consolidado no bloco F
'            sSQLNF = sSQLNFBase & " AND "
'            sSQLNF = sSQLNF & "T.Faturamento = 1 AND Status <> 7 AND (Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) AND (Tipo=2 AND DataEmissao >= ? AND DataEmissao <= ?) AND F.FilialEmpresa < 50 "
'            sSQLNF = sSQLNF & " ORDER BY N.NumIntDoc" 'para garantir que reexecute passando pelos registros na mesma ordem...
'
'        End If
    
        If iPasso = 1 Then
            lComandoAux = alComando(2)
        Else
            lComandoAux = alComando(20)
        End If
        
        tNFiscal.sSerie = String(STRING_MAXIMO, 0)
        tNFiscal.sPlaca = String(STRING_MAXIMO, 0)
        tNFiscal.sPlacaUF = String(STRING_MAXIMO, 0)
        tTipoDocInfo.sNomeTelaNFiscal = String(STRING_MAXIMO, 0)
        tTributacaoDoc.sNaturezaOpInterna = String(STRING_MAXIMO, 0)
        tNFiscal.sCodVerificacaoNFe = String(STRING_MAXIMO, 0)
        tNFiscal.sChvNFe = String(STRING_MAXIMO, 0)
        sCgc = String(STRING_CGC, 0)
            
        lErro = Comando_Executar(lComandoAux, sSQLNF, _
                                                sCgc, tNFiscal.iFilialEmpresa, tNFiscal.lNumIntDoc, tNFiscal.iRecibo, tNFiscal.sCodVerificacaoNFe, tNFiscal.sChvNFe, tNFiscal.iNFe, tNFiscal.iTipoNFiscal, tTipoDocInfo.iTipo, tNFiscal.lCliente, tNFiscal.iFilialCli, tNFiscal.lFornecedor, tNFiscal.iFilialForn, tNFiscal.iStatus, tTipoDocInfo.iComplementar, tNFiscal.sSerie, tNFiscal.lNumNotaFiscal, tNFiscal.dtDataEmissao, tNFiscal.dtDataEntrada, tNFiscal.dtDataSaida, tNFiscal.dValorTotal, _
                                                tNFiscal.iClasseDocCPR, tNFiscal.lNumIntDocCPR, tNFiscal.dValorDesconto, tNFiscal.dValorProdutos, tNFiscal.dValorFrete, tNFiscal.iFreteRespons, tNFiscal.dValorSeguro, tNFiscal.dValorOutrasDespesas, tNFiscal.iFilialEntrega, tNFiscal.iCodTransportadora, tNFiscal.sPlaca, tNFiscal.sPlacaUF, _
                                                tNFiscal.lVolumeQuant, tNFiscal.dPesoLiq, tNFiscal.dPesoBruto, tTipoDocInfo.iTipoMovtoEstoque, tTipoDocInfo.sNomeTelaNFiscal, tTributacaoDoc.iTipoTributacao, tTributacaoDoc.dIPIBase, tTributacaoDoc.dIPIValor, tTributacaoDoc.dICMSBase, tTributacaoDoc.dICMSValor, _
                                                tTributacaoDoc.dICMSSubstBase, tTributacaoDoc.dICMSSubstValor, tTributacaoDoc.dISSBase, tTributacaoDoc.dISSValor, tTributacaoDoc.dPISRetido, tTributacaoDoc.dCOFINSRetido, tTributacaoDoc.sNaturezaOpInterna, tNFiscal.lNumIntNotaOriginal, tTipoDocInfo.iFaturamento, tTributacaoDoc.dISSRetido, _
                                                left(objEFD.objFilialEmpresa.sCgc, 8) & "%", objEFD.dtDataIni, objEFD.dtDataFim, objEFD.dtDataIni, objEFD.dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 209377
                               
        lErro = Comando_BuscarPrimeiro(lComandoAux)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209378
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objDoc = New ClassSpedDocFiscais
            
            objDoc.iFilialEmpresa = tNFiscal.iFilialEmpresa
            objDoc.iTipoDoc = SPED_DOC_NF
            objDoc.lNumIntDoc = tNFiscal.lNumIntDoc
            objDoc.lNumIntSped = objEFD.lNumIntSped
            
            sChave = "NF ID: " & CStr(tNFiscal.lNumIntDoc)
            
            'Lê LivRegES
            lErro = Comando_Executar(alComando(9), "SELECT Modelo FROM LivRegES WHERE TipoNumIntDocOrigem = 0 AND NumIntDocOrigem = ? ", tTipoDocInfo.iModeloArqICMS, tNFiscal.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 209379
        
            lErro = Comando_BuscarPrimeiro(alComando(9))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209380
        
            'Se não tiver registro em LivRegES não tem ICMS
            If lErro <> AD_SQL_SUCESSO Then
                tTipoDocInfo.iModeloArqICMS = 0
                bNaoIncideICMS = True
            Else
                If tTipoDocInfo.iModeloArqICMS >= 9000 Then
                    If tNFiscal.iNFe = MARCADO Then
                        tTipoDocInfo.iModeloArqICMS = 55
                    Else
                        tTipoDocInfo.iModeloArqICMS = 1
                    End If
                End If
                bNaoIncideICMS = False
            End If
            
            If right(UCase(tNFiscal.sSerie), 2) <> "-E" Then
                objDoc.sSerie = tNFiscal.sSerie
            Else
                objDoc.sSerie = left(tNFiscal.sSerie, Len(tNFiscal.sSerie) - 2)
            End If
            objDoc.lNumDocumento = tNFiscal.lNumNotaFiscal
            objDoc.dtDataEmissao = tNFiscal.dtDataEmissao
            
            If tTipoDocInfo.iTipo = 2 Then
                objDoc.sOperacao = 1 'Saída
                objDoc.dtDataES = tNFiscal.dtDataSaida
            Else
                objDoc.sOperacao = 0 'Entrada;
                objDoc.dtDataES = tNFiscal.dtDataEntrada
            End If
            
            If tTipoDocInfo.iTipo = 2 Or tTipoDocInfo.iTipo = 1 Then
                objDoc.sEmitente = "0" 'Emissão própria
            Else
                objDoc.sEmitente = "1" 'Terceiros
            End If
                   
            'Código Descrição
            '00 Documento regular
            '01 Escrituração extemporânea de documento regular
            '02 Documento cancelado
            '03 Escrituração extemporânea de documento cancelado
            '04 NF-e ou CT-e – denegado
            '05 NF-e ou CT-e - Numeração inutilizada
            '06 Documento Fiscal Complementar
            '07 Escrituração extemporânea de documento complementar
            '08 Documento Fiscal emitido com base em Regime Especial ou Norma Específica
            objDoc.iSituacao = 0
            If tNFiscal.iStatus = STATUS_CANCELADO Then
                objDoc.iSituacao = 2
                'Busca Denegado senão vai ser cancelada mesmo
                If tNFiscal.iNFe = 1 Then
        
                    'se for uma nota eletronica denegada no ambiente de producao
                    lErro = Comando_Executar(alComando(4), "SELECT tpAmb FROM NFeFedProtNFe WHERE NumIntNF = ? AND cStat IN ('110', '301','302') AND tpAmb = 0", tNFiscal.iCanal, tNFiscal.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 209381
                
                    lErro = Comando_BuscarPrimeiro(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209382
                
                    'nota fiscal denegada
                    If lErro = AD_SQL_SUCESSO Then objDoc.iSituacao = 4
                        
                End If
            Else
                If tTipoDocInfo.iComplementar = MARCADO Then
                    objDoc.iSituacao = 6
'                Else
'                    'Busca Denegado senão vai ser regular
'                    If tNFiscal.iNFe = 1 Then
'
'                        'se for uma nota eletronica denegada no ambiente de producao
'                        lErro = Comando_Executar(alComando(4), "SELECT tpAmb FROM NFeFedProtNFe WHERE NumIntNF = ? AND cStat IN ('110', '301','302') AND tpAmb = 0", tNFiscal.iCanal, tNFiscal.lNumIntDoc)
'                        If lErro <> AD_SQL_SUCESSO Then gError 209381
'
'                        lErro = Comando_BuscarPrimeiro(alComando(4))
'                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209382
'
'                        'nota fiscal denegada
'                        If lErro = AD_SQL_SUCESSO Then objDoc.iSituacao = 4
'
'                    End If
                End If
            End If
            
            Select Case tNFiscal.iFreteRespons
            
                Case 1 'Emitente
                    If objDoc.dtDataEmissao < #1/1/2012# Then
                        objDoc.sFrete = "1"
                    Else
                        objDoc.sFrete = "0"
                    End If
                Case 2 'Destinatário
                    If objDoc.dtDataEmissao < #1/1/2012# Then
                        objDoc.sFrete = "2"
                    Else
                        objDoc.sFrete = "1"
                    End If
                Case 3 'Terceiros
                    If objDoc.dtDataEmissao < #1/1/2012# Then
                        objDoc.sFrete = "0"
                    Else
                        objDoc.sFrete = "2"
                    End If
                Case 4 'Sem Frete
                    If objDoc.dtDataEmissao < #1/1/2012# Then
                        objDoc.sFrete = "9"
                    Else
                        objDoc.sFrete = "9"
                    End If
            
            End Select
            
            If tTipoDocInfo.iModeloArqICMS = 65 Then objDoc.sFrete = "9"
            
            objDoc.sModelo = Format(tTipoDocInfo.iModeloArqICMS, "00")
            
            objDoc.dValorBaseICMS = tTributacaoDoc.dICMSBase
            objDoc.dValorBaseICMSST = tTributacaoDoc.dICMSSubstBase
            objDoc.dValorBaseIPI = tTributacaoDoc.dIPIBase
            objDoc.dValorBaseISS = tTributacaoDoc.dISSBase
            objDoc.dValorDespesas = tNFiscal.dValorOutrasDespesas
            objDoc.dValorFrete = tNFiscal.dValorFrete
            objDoc.dValorICMS = tTributacaoDoc.dICMSValor
            objDoc.dValorICMSST = tTributacaoDoc.dICMSSubstValor
            objDoc.dValorIPI = tTributacaoDoc.dIPIValor
            objDoc.dValorISS = tTributacaoDoc.dISSValor
            objDoc.dValorISSRet = tTributacaoDoc.dISSRetido
            objDoc.dValorSeguro = tNFiscal.dValorSeguro
            objDoc.dValorTotal = tNFiscal.dValorTotal
            objDoc.dValorPisRetido = tTributacaoDoc.dPISRetido
            objDoc.dValorCofinsRetido = tTributacaoDoc.dCOFINSRetido
            objDoc.sCgc = sCgc
    
            objDoc.iFilialCli = tNFiscal.iFilialCli
            objDoc.iFilialForn = tNFiscal.iFilialForn
            objDoc.lCliente = tNFiscal.lCliente
            objDoc.lFornecedor = tNFiscal.lFornecedor
            
            'Pegar a chave
            If tNFiscal.iNFe = MARCADO Then
            
                objDoc.sChaveEletronica = tNFiscal.sChvNFe
            
                sCodVerificacaoNFe = String(STRING_MAXIMO, 0)
            
                'Pega a chave de autorização
                lErro = Comando_Executar(alComando(7), "SELECT ChNFe FROM NFeFedProtNFe WHERE NumIntNF = ? AND cStat IN (?,?) AND tpAmb = 0 ORDER BY cStat", sCodVerificacaoNFe, tNFiscal.lNumIntDoc, 100, 302)
                If lErro <> AD_SQL_SUCESSO Then gError 209387
            
                lErro = Comando_BuscarPrimeiro(alComando(7))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209388
            
                If lErro = AD_SQL_SUCESSO Then objDoc.sChaveEletronica = sCodVerificacaoNFe
                
            ElseIf tNFiscal.iRecibo = MARCADO And Len(Trim(tNFiscal.sCodVerificacaoNFe)) = 0 Then
            
                sCodVerificacaoNFe = String(STRING_MAXIMO, 0)
            
                'Pega a chave de autorização
                lErro = Comando_Executar(alComando(8), "SELECT CodigoVerificacao FROM RPSWebProt WHERE NumIntNF = ? AND Ambiente = 1 AND CodigoVerificacao <> ''", sCodVerificacaoNFe, tNFiscal.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 209389
            
                lErro = Comando_BuscarPrimeiro(alComando(8))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209390
            
                If lErro = AD_SQL_SUCESSO Then objDoc.sChaveEletronica = sCodVerificacaoNFe
                
            Else
                objDoc.sChaveEletronica = tNFiscal.sCodVerificacaoNFe 'Pode ser RPS com envio via arquivo
            End If
            
            'Pegar a condição de pagamento
            iCondPagto = 0
            Select Case tNFiscal.iClasseDocCPR
                Case CPR_TITULO_RECEBER
                
                    lErro = Comando_Executar(alComando(5), "SELECT CondicaoPagto FROM TitulosRecTodos WHERE NumIntDoc = ?", iCondPagto, tNFiscal.lNumIntDocCPR)
                    If lErro <> AD_SQL_SUCESSO Then gError 209391
                
                    lErro = Comando_BuscarPrimeiro(alComando(5))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209392
                    
                Case CPR_TITULO_PAGAR
            
                    lErro = Comando_Executar(alComando(6), "SELECT CondicaoPagto FROM TitulosPagTodos WHERE NumIntDoc = ?", iCondPagto, tNFiscal.lNumIntDocCPR)
                    If lErro <> AD_SQL_SUCESSO Then gError 209393
                
                    lErro = Comando_BuscarPrimeiro(alComando(6))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209394
                    
                Case CPR_NF_PAGAR
                
                    lErro = Comando_Executar(alComando(16), "SELECT T.CondicaoPagto FROM TitulosPagTodos AS T, NFsPag_Todas As N WHERE T.NumIntDoc = N.NumIntTitPag AND N.NumIntDoc = ?", iCondPagto, tNFiscal.lNumIntDocCPR)
                    If lErro <> AD_SQL_SUCESSO Then gError 211985
                
                    lErro = Comando_BuscarPrimeiro(alComando(16))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211986
                    
            End Select
            
            'Indicador do tipo de pagamento:
            '0- À vista;
            '1- A prazo;
            '9- Sem pagamento.
            Select Case iCondPagto
                Case 0
                    objDoc.sPagamento = "9"
                Case 1
                    objDoc.sPagamento = "0"
                Case Else
                    objDoc.sPagamento = "1"
            End Select
            
            sSQLItemNF = "SELECT I.NumIntDoc, I.Item, I.Produto, I.UnidadeMed, I.Quantidade, I.PrecoUnitario, I.ValorDesconto, I.DescricaoItem,  "
            sSQLItemNF = sSQLItemNF & "T.NaturezaOP, T.TipoTributacao, T.IPITipo, T.IPIBaseCalculo, T.IPIALiquota, T.IPIValor,T.IPICredito, T.ICMSTipo, T.CST, "
            sSQLItemNF = sSQLItemNF & "T.ICMSBase, T.ICMSPercRedBase, T.ICMSAliquota, T.ICMSValor, T.ICMSCredito, T.ICMSSubstBase, T.ICMSSubstAliquota, T.ICMSSubstValor, "
            sSQLItemNF = sSQLItemNF & "T.PISCredito, T.CofinsCredito, T.ICMSSubstPercRedBase, T.ICMSSubstPercMVA, T.PISTipo, T.PISBase, T.PISAliquota, T.PISAliquotaValor, T.PISQtde, T.PISValor, "
            sSQLItemNF = sSQLItemNF & "T.PISSTBase, T.PISSTAliquota, T.PISSTAliquotaValor, T.PISSTQtde, T.PISSTValor, T.CofinsTipo, T.CofinsBase, T.CofinsAliquota, T.CofinsAliquotaValor, T.CofinsQtde, T.CofinsValor, "
            sSQLItemNF = sSQLItemNF & "T.CofinsSTBase, T.CofinsSTAliquota, T.CofinsSTAliquotaValor, T.CofinsSTQtde, T.CofinsSTValor, T.ISSBase, T.ISSAliquota, T.ISSValor, T.IPIEnquadramentoCodigo, "
            sSQLItemNF = sSQLItemNF & "T.ValorFreteItem , T.ValorSeguroItem, T.ValorOutrasDespesasItem, T.OrigemMercadoria, TT.CSTEntrada, TT.CSTSaida, TM.ICMSIncide, TM.ISSIncide, P.Natureza, T.NatBCCred, T.IPICodProduto, T.ISSCidadeIBGE, T.ISSCodServ "
            sSQLItemNF = sSQLItemNF & "FROM ItensNFiscal AS I, TributacaoDocItem AS T, TiposTribIPI As TT, TiposDeTributacaoMovto As TM, Produtos AS P "
            sSQLItemNF = sSQLItemNF & "WHERE P.Codigo = I.Produto AND TM.Tipo = T.TipoTributacao AND TT.Tipo = T.IPITipo AND T.Tipodoc = 0 AND T.NumIntDocItem = I.NumIntDoc AND I.NumIntNF = ? ORDER BY I.Item "
            
            tItemNF.sProduto = String(STRING_MAXIMO, 0)
            tItemNF.sUnidadeMed = String(STRING_MAXIMO, 0)
            tItemNF.sDescricaoItem = String(STRING_MAXIMO, 0)
            tTribDocItem.sNaturezaOp = String(STRING_MAXIMO, 0)
            tTribDocItem.sCST = String(STRING_MAXIMO, 0)
            tTribDocItem.sIPIEnquadramentoCodigo = String(STRING_MAXIMO, 0)
            tTribDocItem.sNatBCCred = String(STRING_MAXIMO, 0)
            tTribDocItem.sIPICodProduto = String(STRING_MAXIMO, 0)
            tTribDocItem.sISSCidadeIBGE = String(STRING_MAXIMO, 0)
            tTribDocItem.sISSCodServ = String(STRING_MAXIMO, 0)
            
            lErro = Comando_Executar(alComando(3), sSQLItemNF, _
                                                    tItemNF.lNumIntDoc, tItemNF.iItem, tItemNF.sProduto, tItemNF.sUnidadeMed, tItemNF.dQuantidade, tItemNF.dPrecoUnitario, tItemNF.dValorDesconto, tItemNF.sDescricaoItem, _
                                                    tTribDocItem.sNaturezaOp, tTribDocItem.iTipoTributacao, tTribDocItem.iIPITipo, tTribDocItem.dIPIBaseCalculo, tTribDocItem.dIPIAliquota, tTribDocItem.dIPIValor, tTribDocItem.dIPICredito, tTribDocItem.iICMSTipo, tTribDocItem.sCST, _
                                                    tTribDocItem.dICMSBase, tTribDocItem.dICMSPercRedBase, tTribDocItem.dICMSAliquota, tTribDocItem.dICMSValor, tTribDocItem.dICMSCredito, tTribDocItem.dICMSSubstBase, tTribDocItem.dICMSSubstAliquota, tTribDocItem.dICMSSubstValor, _
                                                    tTribDocItem.dPISCredito, tTribDocItem.dCOFINSCredito, tTribDocItem.dICMSSubstPercRedBase, tTribDocItem.dICMSSubstPercMVA, tTribDocItem.iPISTipo, tTribDocItem.dPISBase, tTribDocItem.dPISAliquota, tTribDocItem.dPISAliquotaValor, tTribDocItem.dPISQtde, tTribDocItem.dPISValor, _
                                                    tTribDocItem.dPISSTBase, tTribDocItem.dPISSTAliquota, tTribDocItem.dPISSTAliquotaValor, tTribDocItem.dPISSTQtde, tTribDocItem.dPISSTValor, tTribDocItem.iCOFINSTipo, tTribDocItem.dCOFINSBase, tTribDocItem.dCOFINSAliquota, tTribDocItem.dCOFINSAliquotaValor, tTribDocItem.dCOFINSQtde, tTribDocItem.dCOFINSValor, _
                                                    tTribDocItem.dCOFINSSTBase, tTribDocItem.dCOFINSSTAliquota, tTribDocItem.dCOFINSSTAliquotaValor, tTribDocItem.dCOFINSSTQtde, tTribDocItem.dCOFINSSTValor, tTribDocItem.dISSBase, tTribDocItem.dISSAliquota, tTribDocItem.dISSValor, tTribDocItem.sIPIEnquadramentoCodigo, _
                                                    tTribDocItem.dValorFreteItem, tTribDocItem.dValorSeguroItem, tTribDocItem.dValorOutrasDespesasItem, tTribDocItem.iOrigemMercadoria, iIPICSTE, iIPICSTS, iICMSIncide, iISSIncide, iNatureza, tTribDocItem.sNatBCCred, tTribDocItem.sIPICodProduto, tTribDocItem.sISSCidadeIBGE, tTribDocItem.sISSCodServ, _
                                                    tNFiscal.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 209395
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209396
        
            Do While lErro <> AD_SQL_SEM_DADOS
            
                sChave = "NF ID: " & CStr(tNFiscal.lNumIntDoc) & " Item: " & CStr(tItemNF.iItem)
            
                Set objItem = New ClassSpedDocFiscaisItens
                objDoc.colItens.Add objItem
                
                objItem.dPrecoUnitario = tItemNF.dPrecoUnitario
                objItem.dCOFINSAliquota = tTribDocItem.dCOFINSAliquota
                objItem.dCofinsAliquotaQtd = tTribDocItem.dCOFINSAliquotaValor
                objItem.dCofinsBCQtd = tTribDocItem.dCOFINSQtde
                objItem.dICMSAliquota = tTribDocItem.dICMSAliquota
                objItem.dICMSSTAliquota = tTribDocItem.dICMSSubstAliquota
                objItem.dAliquotaISS = tTribDocItem.dISSAliquota
                objItem.dIPIAliquota = tTribDocItem.dIPIAliquota
                objItem.dPISAliquota = tTribDocItem.dPISAliquota
                objItem.dPISAliquotaQtd = tTribDocItem.dPISAliquotaValor
                objItem.dPISBCQtd = tTribDocItem.dPISQtde
                objItem.dQuantidade = tItemNF.dQuantidade
                objItem.dValorBaseCofins = tTribDocItem.dCOFINSBase
                objItem.dValorBaseICMS = tTribDocItem.dICMSBase
                objItem.dValorBaseICMSST = tTribDocItem.dICMSSubstBase
                objItem.dValorBaseIPI = tTribDocItem.dIPIBaseCalculo
                objItem.dValorBasePIS = tTribDocItem.dPISBase
                objItem.dValorBaseISS = tTribDocItem.dISSBase
                objItem.dValorCofins = tTribDocItem.dCOFINSValor
                'objItem.dValorCofinsRetido
                objItem.dValorCofinsST = tTribDocItem.dCOFINSSTValor
                objItem.dValorDesconto = tItemNF.dValorDesconto
                objItem.dValorDespesas = tTribDocItem.dValorOutrasDespesasItem
                objItem.dValorFrete = tTribDocItem.dValorFreteItem
                objItem.dValorICMS = tTribDocItem.dICMSValor
                objItem.dValorICMSST = tTribDocItem.dICMSSubstValor
                objItem.dValorIPI = tTribDocItem.dIPIValor
                objItem.dValorISS = tTribDocItem.dISSValor
                'objItem.dValorNaoTributado
                objItem.dValorPIS = tTribDocItem.dPISValor
                'objItem.dValorPisRetido
                objItem.dValorPisST = tTribDocItem.dPISSTValor
                objItem.dValorSeguro = tTribDocItem.dValorSeguroItem
                objItem.iItem = tItemNF.iItem
                objItem.sCFOP = tTribDocItem.sNaturezaOp
                objItem.sNatBCCred = tTribDocItem.sNatBCCred
                objItem.sNCM = tTribDocItem.sIPICodProduto
                objItem.sCodISSServ = tTribDocItem.sISSCodServ
                
                '"A majoração da alíquota em referência, pela Medida provisória nº 563/2012, foi tão somente no cálculo da Cofins na importação, especificada no art. 8º da Lei nº 10.865/04.
                'Referido aumento não se estendeu aos créditos na importação, que continuam tendo o seu valor determinado de acordo com o disposto no art. 15 da referida Lei, o qual estabelece em seu § 3º que para o cálculo dos créditos na importação aplicam-se as alíquotas dispostas no caput do art. 2º das Leis nº 10.637/02 (PIS) e nº 10.833/03 (Cofins).
                'O objetivo do aumento em 1% foi o de desestimular a importação dos produtos listados e sujeitos à Contribuição Previdenciária sobre a Receita Bruta (Lei nº 12.546). Caso o aumento fosse repassado para os créditos, teria os seus efeitos anulados.
                'Desta forma:
                'Considerando que na EFD Contribuições não se escritura as operações referentes ao PIS/Pasep - Importação e à Cofins - Importação,  e sim e tão somente, as operações referentes aos créditos das referidas importações;
                'Considerando que na apuração dos créditos sobre a importação, as alíquotas aplicáveis (1,65% e 7,6%) não sofreram alterações.
                'Resta esclarecido que o aumento em 1% na alíquota da Cofins-Importação, não é passível de escrituração da EFD-Contribuições.
                'OBS; Os créditos nas importações, previstos no art. 15 da Lei nº 10.865/04, não correspondem exatamente aos valores pagos na importação e sim, corresponde aos valores determinados conforme as disposições do referido artigo - Base de Cálculo x Alíquota = Valor do Crédito."
                'se é uma importação com Cofins de 8,6% ajusta pois o 1% a mais não constitui crédito
                If left(objItem.sCFOP, 1) = "3" And Abs(objItem.dCOFINSAliquota - 0.086) < QTDE_ESTOQUE_DELTA Then
                    objItem.dCOFINSAliquota = 0.076
                    objItem.dValorCofins = Arredonda_Moeda(objItem.dValorBaseCofins * objItem.dCOFINSAliquota)
                End If
                
                'Se paga contribuição previdenciária pela receita bruta
                If objEFD.objFilialEmpresa.iContribPrevPorRecBruta = MARCADO Then
                    objItem.dAliquotaContrPrev = objEFD.objFilialEmpresa.dAliquotaPadraoContribPrev
                    If Len(Trim(objItem.sNCM)) > 0 Then
                    
                        sCodAtividadeTab511 = String(STRING_MAXIMO, 0)
                    
                        lErro = Comando_Executar(alComando(21), "SELECT CodAtividadeTab511, AliquotaContrPrev FROM CodAtivContrPrev WHERE DataIniEscrituracao <= ? AND DataFimEscrituracao >= ? AND NCM = LEFT(?,LEN(NCM)) AND ExNCM NOT LIKE ?", sCodAtividadeTab511, dAliquotaContrPrev, objDoc.dtDataEmissao, objDoc.dtDataEmissao, objItem.sNCM, "%" & objItem.sNCM & "%")
                        If lErro <> AD_SQL_SUCESSO Then gError 211987
                        
                        lErro = Comando_BuscarPrimeiro(alComando(21))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211988
                    
                        If lErro <> AD_SQL_SEM_DADOS Then
                            objItem.sCodAtividadeTab511 = sCodAtividadeTab511
                            objItem.dAliquotaContrPrev = dAliquotaContrPrev
                        Else
                            objItem.sCodAtividadeTab511 = "99999999" 'Código Genérico - Contribuição Previdenciária sobre Receitas
                        End If
                    
                    Else
                        If objDoc.sModelo = "00" Or iISSIncide = MARCADO Then
                            objItem.sCodAtividadeTab511 = "00009999" 'Código Genérico - Serviços
                        End If
                    End If
                End If
                
                If tTipoDocInfo.iTipo = 2 Or tTribDocItem.iCOFINSTipo >= 50 Then
                    objItem.sCofinsCST = Format(tTribDocItem.iCOFINSTipo, "00")
                Else
                    Select Case tTribDocItem.iCOFINSTipo
                        Case 1
                            objItem.sCofinsCST = "56"
                        Case 6
                            objItem.sCofinsCST = "73"
                        Case 8
                            objItem.sCofinsCST = "74"
                        Case 9
                            objItem.sCofinsCST = "72"
                        Case 49
                            objItem.sCofinsCST = "98"
                        Case Else
                            objItem.sCofinsCST = Format(tTribDocItem.iCOFINSTipo, "00")
                    End Select
                End If
                
                objItem.sDescricao = tItemNF.sDescricaoItem
                objItem.sICMSCST = tTribDocItem.sCST
                objItem.sIPICodEnq = Format(tTribDocItem.sIPIEnquadramentoCodigo, "000")
                
                If tTipoDocInfo.iTipo = 2 Then
                    objItem.sIPICST = Format(iIPICSTS, "00") 'Saída
                Else
                    objItem.sIPICST = Format(iIPICSTE, "00") 'Entrada
                End If
                            
                If tTipoDocInfo.iTipo = 2 Or tTribDocItem.iPISTipo >= 50 Then
                    objItem.sPISCST = Format(tTribDocItem.iPISTipo, "00")
                Else
                    Select Case tTribDocItem.iPISTipo
                        Case 1
                            objItem.sPISCST = "56"
                        Case 6
                            objItem.sPISCST = "73"
                        Case 8
                            objItem.sPISCST = "74"
                        Case 9
                            objItem.sPISCST = "72"
                        Case 49
                            objItem.sPISCST = "98"
                        Case Else
                            objItem.sPISCST = Format(tTribDocItem.iPISTipo, "00")
                    End Select
                End If
                
                objItem.sProduto = tItemNF.sProduto
                objItem.sUM = tItemNF.sUnidadeMed
                
                'Movimentação física do ITEM/PRODUTO:
                '0. SIM
                '1. NÃO
                If tTipoDocInfo.iTipoMovtoEstoque = 0 Then
                    objItem.sMovFisica = "1"
                Else
                    objItem.sMovFisica = "0"
                End If
                
                objItem.dPrecoTotal = objItem.dQuantidade * objItem.dPrecoUnitario
                If iNatureza = NATUREZA_PROD_SERVICO Then
                    objItem.dValorServico = objItem.dPrecoTotal
                    objItem.sMovFisica = "1"
                Else
                    objDoc.dValorMercadoria = objDoc.dValorMercadoria + objItem.dPrecoTotal
                End If
                
                If Len(Trim(tTribDocItem.sISSCidadeIBGE)) > 0 Then objDoc.sCodMunicServ = tTribDocItem.sISSCidadeIBGE
                
                objDoc.dValorBaseCofins = objDoc.dValorBaseCofins + objItem.dValorBaseCofins
                objDoc.dValorBasePIS = objDoc.dValorBasePIS + objItem.dValorBasePIS
                objDoc.dValorCofins = objDoc.dValorCofins + objItem.dValorCofins
                objDoc.dValorCofinsRetido = objDoc.dValorCofinsRetido + objItem.dValorCofinsRetido
                objDoc.dValorCofinsST = objDoc.dValorCofinsST + objItem.dValorCofinsST
                objDoc.dValorDesconto = objDoc.dValorDesconto + objItem.dValorDesconto
                objDoc.dValorPIS = objDoc.dValorPIS + objItem.dValorPIS
                objDoc.dValorPisRetido = objDoc.dValorPisRetido + objItem.dValorPisRetido
                objDoc.dValorPisST = objDoc.dValorPisST + objItem.dValorPisST
                objDoc.dValorServico = objDoc.dValorServico + objItem.dValorServico
                objDoc.dValorNaoTributado = objDoc.dValorNaoTributado + objItem.dValorNaoTributado
            
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209397
                        
            Loop
            
            'Tabela de modelo de documentos fiscais e registros correspondentes na EFD-PIS/Cofins:
            'Página 7/8 do manual
            'Guia Prático EFD-PIS/COFINS – Versão 1.00
            'Atualização: 31 de janeiro de 2011
            If objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA And iPasso = 1 Then
                If objDoc.sModelo = "03" Or objDoc.sModelo = "01" Or objDoc.sModelo = "1B" Or objDoc.sModelo = "04" Or objDoc.sModelo = "55" Or objDoc.sModelo = "65" Then
                    
                    If tTipoDocInfo.iFaturamento = 2 And objEFD.objFilialEmpresa.iCOFINSNaoCumulativo = DESMARCADO Then
                        'Se é uma devolução no regime cumulativo
                        objDoc.sBloco = "M"
                        objDoc.sRegistro = "M220"
                    Else
                        If iISSIncide = MARCADO And objEFD.objFilialEmpresa.objEndereco.sSiglaEstado = "DF" Then 'Or objDoc.sModelo = "00" Then
                            objDoc.sBloco = "B"
                            If tTipoDocInfo.iTipo = 2 Then
                                objDoc.sRegistro = "B020"
                            Else
                                objDoc.sRegistro = "B440"
                            End If
                        Else
                            objDoc.sBloco = "C"
                            objDoc.sRegistro = "C100"
                        End If
                    End If
                ElseIf objDoc.sModelo = "02" Then
                    objDoc.sBloco = "C"
                    objDoc.sRegistro = "C380"
                ElseIf objDoc.sModelo = "06" Or objDoc.sModelo = "28" Or objDoc.sModelo = "29" Then
                    objDoc.sBloco = "C"
                    If tTipoDocInfo.iTipo = 2 Then
                        objDoc.sRegistro = "C600" 'Fornecimento:
                    Else
                        objDoc.sRegistro = "C500" 'Aquisição:
                    End If
                ElseIf objDoc.sModelo = "07" Or objDoc.sModelo = "08" Or objDoc.sModelo = "09" Or objDoc.sModelo = "10" Or objDoc.sModelo = "11" Or objDoc.sModelo = "26" Or objDoc.sModelo = "27" Or objDoc.sModelo = "57" Or objDoc.sModelo = "8B" Then
                    objDoc.sBloco = "D"
                    If tTipoDocInfo.iTipo = 2 Then
                        objDoc.sRegistro = "D200" 'Fornecimento:
                    Else
                        objDoc.sRegistro = "D100" 'Aquisição:
                    End If
                ElseIf objDoc.sModelo = "21" Or objDoc.sModelo = "22" Then
                    objDoc.sBloco = "D"
                    If tTipoDocInfo.iTipo = 2 Then
                        objDoc.sRegistro = "D600" 'Fornecimento:
                    Else
                        objDoc.sRegistro = "D500" 'Aquisição:
                    End If
                End If
            ElseIf objEFD.iRegimeApuracao = objEFD.TRATA_REG_COMPETENCIA And iPasso = 2 Then
                objDoc.sBloco = "" 'Deixa vazio pois vai servir só para apurar a Receita, mas não vai gerar registro no bloco F
                objDoc.sRegistro = "F500" 'COMPOSIÇÃO DA RECEITA ESCRITURADA NO PERÍODO
            ElseIf objEFD.iRegimeApuracao = objEFD.TRATA_REG_CAIXA And iPasso = 1 Then
                objDoc.sBloco = "F"
                objDoc.sRegistro = "F500" 'COMPOSIÇÃO DA RECEITA ESCRITURADA NO PERÍODO
            ElseIf objEFD.iRegimeApuracao = objEFD.TRATA_REG_CAIXA And iPasso = 2 Then
                objDoc.sBloco = "1"
                objDoc.sRegistro = "1900" 'CONSOLIDAÇÃO DOS DOCUMENTOS EMITIDOS NO PERÍODO GERADORES DE RECEITA
                If objDoc.sModelo = "00" Or iISSIncide = MARCADO Then
                    objDoc.sModelo = "98" 'Nota Fiscal de Prestação de Serviços (ISSQN)
                End If
            End If
            
            'Indicador do tipo de pagamento:
            '0- À vista;
            '1- A prazo;
            '9- Sem pagamento
            'C100
            'Obs.: A partir de 01/07/2012 passará a ser:
            'Indicador do tipo de pagamento:
            '0- À vista;
            '1- A prazo;
            '2 - Outros
            If objEFD.iCodVersao >= 3 Then ' versão posterior a 01/07/2012
                If objDoc.sPagamento = "9" And objDoc.sBloco <> "A" Then
                    objDoc.sPagamento = "2"
                End If
            End If
            
            If tNFiscal.iClasseDocCPR = CPR_TITULO_RECEBER Then
            
                'Acumula o total recebido por cada nf (depois vai ratear por CST/Alíquota)
                sSQLCR = "SELECT SUM(BP.ValorRecebido) "
                sSQLCR = sSQLCR & "FROM NFiscal AS N, TitulosRecTodos AS TR, ParcelasRecTodas As P, BaixasRec AS B, BaixasParcRec As BP "
                sSQLCR = sSQLCR & "WHERE N.Status <> 7 AND N.ClasseDocCPR = 2 AND N.NumIntDocCPR = TR.NumIntDoc AND TR.Status <> 5 AND "
                sSQLCR = sSQLCR & "TR.NumIntDoc = P.NumIntTitulo AND BP.NumIntParcela = P.NumIntDoc AND B.NumIntBaixa = BP.NumIntBaixa AND "
                sSQLCR = sSQLCR & "B.Status <> 5 AND BP.Status <> 5 AND B.Motivo NOT IN (4,6) AND ((B.Motivo <> 5 AND B.Data BETWEEN ? AND ?) OR (B.Motivo = 5 AND N.DataEmissao BETWEEN ? AND ?)) AND N.NumIntDoc = ? "
    
                lErro = Comando_Executar(alComando(18), sSQLCR, _
                                                        dValorRecebido, objEFD.dtDataIni, objEFD.dtDataFim, objEFD.dtDataIni, objEFD.dtDataFim, objDoc.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 211989
                
                lErro = Comando_BuscarPrimeiro(alComando(18))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211990
    
                objDoc.dValorRecebido = dValorRecebido
                
                lErro = Comando_Executar(alComando(19), "SELECT TR.Valor FROM NFiscal AS N, TitulosRecTodos AS TR WHERE N.Status <> 7 AND N.ClasseDocCPR = 2 AND N.NumIntDocCPR = TR.NumIntDoc AND TR.Status <> 5 AND N.NumIntDoc = ?", dValorNoCR, objDoc.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 211991
                
                lErro = Comando_BuscarPrimeiro(alComando(19))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211992
                
                If lErro = AD_SQL_SUCESSO Then objDoc.dValorNoCR = dValorNoCR
                
            End If
            
            If objDoc.sRegistro = "D100" Then
            
                'Indicador da Natureza do Frete Contratado, referente a:
                '0 - Operações de vendas, com ônus suportado pelo estabelecimento vendedor;
                '1 - Operações de vendas, com ônus suportado pelo adquirente;
                '2 - Operações de compras (bens para revenda, matérias-prima e outros produtos, geradores de crédito);
                '3 - Operações de compras (bens para revenda, matérias-prima e outros produtos, não geradores de crédito);
                '4 - Transferência de produtos acabados entre estabelecimentos da pessoa jurídica;
                '5 - Transferência de produtos em elaboração entre estabelecimentos da pessoa jurídica
                '9 - Outras.
                
                bExisteNFOrig = False
                bNFOrigPropEmp = False
                
                If tNFiscal.lNumIntNotaOriginal <> 0 Then
                
                    tNFOrig.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
                
                    lErro = Comando_Executar(alComando(13), "SELECT N.Cliente, N.FilialCli, N.Fornecedor, N.FilialForn, N.ValorFrete, N.NaturezaOP, T.PisCredito, T.CofinsCredito FROM NFiscal AS N, TributacaoDoc AS T WHERE T.TipoDoc = 0 AND T.NumIntDoc = N.NumIntDoc AND N.NumIntDoc = ?", tNFOrig.lCliente, tNFOrig.iFilialCli, tNFOrig.lFornecedor, tNFOrig.iFilialForn, tNFOrig.dValorFrete, tNFOrig.sNaturezaOp, tTribDocOrig.dPISCredito, tTribDocOrig.dCOFINSCredito, tNFiscal.lNumIntNotaOriginal)
                    If lErro <> AD_SQL_SUCESSO Then gError 209393
                
                    lErro = Comando_BuscarPrimeiro(alComando(13))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209394
                    
                    If lErro = AD_SQL_SUCESSO Then
                        
                        bExisteNFOrig = True
                        
                        If tNFOrig.lCliente <> 0 Then
                        
                            sCgc = String(STRING_CGC, 0)
                        
                            lErro = Comando_Executar(alComando(14), "SELECT CGC FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", sCgc, tNFOrig.lCliente, tNFOrig.iFilialCli)
                            If lErro <> AD_SQL_SUCESSO Then gError 209393
                        
                            lErro = Comando_BuscarPrimeiro(alComando(14))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209394
                            
                            If left(sCgc, 8) = left(objEFD.objFilialEmpresa.sCgc, 8) Then bNFOrigPropEmp = True
                        
                        Else
                            
                            sCgc = String(STRING_CGC, 0)
                        
                            lErro = Comando_Executar(alComando(15), "SELECT CGC FROM FiliaisFornecedores WHERE CodFornecedor = ? AND CodFilial = ?", sCgc, tNFOrig.lFornecedor, tNFOrig.iFilialForn)
                            If lErro <> AD_SQL_SUCESSO Then gError 209393
                        
                            lErro = Comando_BuscarPrimeiro(alComando(15))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209394
                            
                            If left(sCgc, 8) = left(objEFD.objFilialEmpresa.sCgc, 8) Then bNFOrigPropEmp = True
                        
                        End If
                        
                    End If
                
                End If
    
                If Not bExisteNFOrig Then
                    '9 - Outras.
                    objDoc.sIndNatFrtPis = "9"
                    objDoc.sIndNatFrtCofins = "9"
                Else
                    If bNFOrigPropEmp Then 'Se a NF original é da própria empresa então é uma transferência
                        Select Case tNFOrig.sNaturezaOp
                            Case ""
                                '5 - Transferência de produtos em elaboração entre estabelecimentos da pessoa jurídica
                                objDoc.sIndNatFrtPis = "5"
                                objDoc.sIndNatFrtCofins = "5"
                            Case Else
                                '4 - Transferência de produtos acabados entre estabelecimentos da pessoa jurídica;
                                objDoc.sIndNatFrtPis = "4"
                                objDoc.sIndNatFrtCofins = "4"
                        End Select
                    Else
                        If tNFiscal.iTipoNFiscal = DOCINFO_CFECT Then 'Compra
                            '2 - Operações de compras (bens para revenda, matérias-prima e outros produtos, geradores de crédito);
                            '3 - Operações de compras (bens para revenda, matérias-prima e outros produtos, não geradores de crédito);
                            If tTribDocOrig.dPISCredito > 0 Then
                                objDoc.sIndNatFrtPis = "2"
                            Else
                                objDoc.sIndNatFrtPis = "3"
                            End If
                            If tTribDocOrig.dCOFINSCredito > 0 Then
                                objDoc.sIndNatFrtCofins = "2"
                            Else
                                objDoc.sIndNatFrtCofins = "3"
                            End If
                            
                        ElseIf tNFiscal.iTipoNFiscal = DOCINFO_CFEV Then 'Venda
                            '0 - Operações de vendas, com ônus suportado pelo estabelecimento vendedor;
                            '1 - Operações de vendas, com ônus suportado pelo adquirente;
                            If tNFOrig.dValorFrete > 0 Then
                                objDoc.sIndNatFrtPis = "1"
                                objDoc.sIndNatFrtCofins = "1"
                            Else
                                objDoc.sIndNatFrtPis = "0"
                                objDoc.sIndNatFrtCofins = "0"
                            End If
                        Else
                            '9 - Outras.
                            objDoc.sIndNatFrtPis = "9"
                            objDoc.sIndNatFrtCofins = "9"
                        End If
                    End If
                
                End If
    
            End If
            
            For Each objItem In objDoc.colItens
            
                'Preenche a Natureza caso seja um CST com direito a crédito
                If Len(Trim(objItem.sNatBCCred)) = 0 And ((objItem.sCofinsCST >= "50" And objItem.sCofinsCST <= "56") Or (objItem.sPISCST >= "50" And objItem.sPISCST <= "56")) Then
                
                    'Código Descrição
                    '01  Aquisição de bens para revenda
                    '02  Aquisição de bens utilizados como insumo
                    '03  Aquisição de serviços utilizados como insumo
                    '04  Energia elétrica e térmica, inclusive sob a forma de vapor
                    '05  Aluguéis de prédios
                    '06  Aluguéis de máquinas e equipamentos
                    '07  Armazenagem de mercadoria e frete na operação de venda
                    '08  Contraprestações de arrendamento mercantil
                    '09  Máquinas, equipamentos e outros bens incorporados ao ativo imobilizado (crédito sobre encargos de depreciação).
                    '10  Máquinas, equipamentos e outros bens incorporados ao ativo imobilizado (crédito com base no valor de aquisição).
                    '11  Amortização e Depreciação de edificações e benfeitorias em imóveis
                    '12  Devolução de Vendas Sujeitas à Incidência Não-Cumulativa
                    '13  Outras Operações com Direito a Crédito
                    '14  Atividade de Transporte de Cargas - Subcontratação
                    '15  Atividade Imobiliária - Custo Incorrido de Unidade Imobiliária
                    '16  Atividade Imobiliária - Custo Orçado de unidade não concluída
                    '17  Atividade de Prestação de Serviços de Limpeza, Conservação e Manutenção - vale-transporte, vale-refeição ou vale-alimentação, fardamento ou uniforme.
                    '18  Estoque de abertura de bens
                
                    Select Case objDoc.sRegistro
                
                        Case "A100"
                            objItem.sNatBCCred = "03"
                
                        Case "C500"
                            objItem.sNatBCCred = "04"
                
                        Case "D100"
                            objItem.sNatBCCred = "14"
                
                        Case "D500"
                            objItem.sNatBCCred = "03"
                
                    End Select
                
                End If
            
            Next
            
            If (objDoc.sRegistro = "F500" Or objDoc.sRegistro = "1900") And objDoc.dValorNoCR <> 0 Then
                lErro = SpedFiscal_Trata_ValorRecebido(objDoc, objEFD)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
            
            lErro = SpedFiscal_Docs_Grava(alComando(0), alComando(1), alComando(22), objDoc, objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            lErro = Comando_BuscarProximo(lComandoAux)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209398
        
        Loop
    Next
    
    'NFS INUTILIZADAS
    sSerieNFInut = String(STRING_SERIE, 0)
    sCGCNFInut = String(STRING_CGC, 0)
    sNumDe = String(STRING_MAXIMO, 0)
    sNumAte = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(17), "SELECT F.CGC, F.FilialEmpresa, N.Data, N.Serie, N.nNFIni, N.nNFFim  FROM NFeFedRetInutNFe AS N, FiliaisEmpresa AS F WHERE N.FilialEMpresa = F.FilialEMpresa AND N.Data BETWEEN ? AND ? AND N.tpAmb = ? AND N.cStat = ? AND N.FilialEmpresa < 50 ORDER BY Data, nNFIni ", sCGCNFInut, iFilialEmpNFInut, dtDataNFInut, sSerieNFInut, sNumDe, sNumAte, objEFD.dtDataIni, objEFD.dtDataFim, 0, "102")
    If lErro <> AD_SQL_SUCESSO Then gError 209377
    
    lErro = Comando_BuscarPrimeiro(alComando(17))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209378
    
    Do While lErro <> AD_SQL_SEM_DADOS

        lNumNFDeInut = StrParaLong(sNumDe)
        lNumNFAteInut = StrParaLong(sNumAte)

        For lNumNFInut = lNumNFDeInut To lNumNFAteInut
    
            Set objDoc = New ClassSpedDocFiscais
            
            objDoc.iFilialEmpresa = tNFiscal.iFilialEmpresa
            objDoc.iTipoDoc = SPED_DOC_NF_INUT
            objDoc.lNumIntDoc = lNumNFInut
            objDoc.lNumIntSped = objEFD.lNumIntSped
            objDoc.lNumDocumento = lNumNFInut
            If right(UCase(sSerieNFInut), 2) <> "-E" Then
                objDoc.sSerie = sSerieNFInut
            Else
                objDoc.sSerie = left(sSerieNFInut, Len(sSerieNFInut) - 2)
            End If
            objDoc.sModelo = "55"
            objDoc.sOperacao = 1 'Saída
            objDoc.sEmitente = "0" 'Emissão própria
            objDoc.iSituacao = 5 '05 NF-e ou CT-e - Numeração inutilizada
            objDoc.iFilialEmpresa = iFilialEmpNFInut
            objDoc.sCgc = sCGCNFInut
            objDoc.sBloco = "C"
            objDoc.sRegistro = "C100"
        
            lErro = SpedFiscal_Docs_Grava(alComando(0), alComando(1), alComando(22), objDoc, objEFD)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        Next
        
        lErro = Comando_BuscarProximo(alComando(17))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209378

    Loop
    'CUPOM FISCAL
    
    'CONHECIMENTO DE TRANSPORTE
    
    'OUTROS (Bloco F)

    SpedFiscal_Grava_Tabelas = SUCESSO

    Exit Function

Erro_SpedFiscal_Grava_Tabelas:

    SpedFiscal_Grava_Tabelas = gErr

    Select Case gErr
   
        Case ERRO_SEM_MENSAGEM
        
        Case 209377, 209378, 209398, 211987 To 211992
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)

        Case 209379, 209380
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)

        Case 209381, 209382, 209387, 209388
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFEFEDPROTNFE", gErr)

        Case 209383, 209384
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr, tNFiscal.lCliente, tNFiscal.iFilialCli)

        Case 209385, 209386
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES", gErr)

        Case 209389, 209390
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RPSWEBPROT", gErr)

        Case 209391, 209392
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECTODOS", gErr)

        Case 209393, 209394, 211985, 211986
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAGTODOS", gErr)

        Case 209395 To 209397
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209399)

    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "SpedFiscal_Grava_Tabelas", sChave)
    
    Exit Function

End Function

Private Function SpedFiscal_Docs_Grava(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal objDoc As ClassSpedDocFiscais, ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long, lAux As Long
Dim objItem As ClassSpedDocFiscaisItens

On Error GoTo Erro_SpedFiscal_Docs_Grava

    lErro = Comando_Executar(lComando1, "INSERT INTO SpedDocFiscais( NumIntSped, TipoDoc, NumIntDoc, Bloco, " & _
                    "Registro, FilialEmpresa, Cliente, FilialCli, Fornecedor, " & _
                    "FilialForn, DataEmissao, DataES, ValorTotal, Operacao, " & _
                    "Emitente, Modelo, Situacao, Serie, NumDocumento, " & _
                    "ChaveEletronica, Frete, Pagamento, ValorDesconto, ValorFrete, " & _
                    "ValorSeguro, ValorDespesas, ValorServico, ValorNaoTributado, ValorBaseICMS, " & _
                    "ValorICMS, ValorBaseICMSST, ValorICMSST, ValorBaseIPI, ValorIPI, " & _
                    "ValorBasePIS, ValorPIS, ValorBaseCofins, ValorCofins, ValorPisRetido, " & _
                    "ValorPisST, ValorCofinsRetido, ValorCofinsST, ValorISS, CGC, CodMensagem, ValorMercadoria, IndNatFrtPis, IndNatFrtCofins, ValorRecebido, ValorNoCR, ValorRecContrPrev, CodMunicServ, ValorBaseISS, ValorBaseISSRet, ValorISSRet) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    objDoc.lNumIntSped, objDoc.iTipoDoc, objDoc.lNumIntDoc, objDoc.sBloco, _
                    objDoc.sRegistro, objDoc.iFilialEmpresa, objDoc.lCliente, objDoc.iFilialCli, objDoc.lFornecedor, _
                    objDoc.iFilialForn, objDoc.dtDataEmissao, objDoc.dtDataES, objDoc.dValorTotal, objDoc.sOperacao, _
                    objDoc.sEmitente, objDoc.sModelo, objDoc.iSituacao, objDoc.sSerie, objDoc.lNumDocumento, _
                    objDoc.sChaveEletronica, objDoc.sFrete, objDoc.sPagamento, objDoc.dValorDesconto, objDoc.dValorFrete, _
                    objDoc.dValorSeguro, objDoc.dValorDespesas, objDoc.dValorServico, objDoc.dValorNaoTributado, objDoc.dValorBaseICMS, _
                    objDoc.dValorICMS, objDoc.dValorBaseICMSST, objDoc.dValorICMSST, objDoc.dValorBaseIPI, objDoc.dValorIPI, _
                    objDoc.dValorBasePIS, objDoc.dValorPIS, objDoc.dValorBaseCofins, objDoc.dValorCofins, objDoc.dValorPisRetido, _
                    objDoc.dValorPisST, objDoc.dValorCofinsRetido, objDoc.dValorCofinsST, objDoc.dValorISS, objDoc.sCgc, objDoc.lCodMensagem, objDoc.dValorMercadoria, objDoc.sIndNatFrtPis, objDoc.sIndNatFrtCofins, objDoc.dValorRecebido, objDoc.dValorNoCR, objDoc.dValorRecContrPrev, objDoc.sCodMunicServ, objDoc.dValorBaseISS, objDoc.dValorBaseISSRet, objDoc.dValorISSRet)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209400
    
    'Verifica se os itens já foram inseridos
    lErro = Comando_Executar(lComando3, "SELECT NumIntSped FROM SpedDocFiscaisItens WHERE NumIntSped = ? AND TipoDoc = ? AND NumIntDoc = ?", _
    lAux, objDoc.lNumIntSped, objDoc.iTipoDoc, objDoc.lNumIntDoc)
    If lErro <> SUCESSO Then gError 209400

    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209400
    
    'Para não causar repetições de NFs de registros de emissão e recebimento ocorridos no mesmo período
    If lErro <> AD_SQL_SUCESSO Then
        
        For Each objItem In objDoc.colItens
            
            objItem.lNumIntSped = objDoc.lNumIntSped
            objItem.lNumIntDoc = objDoc.lNumIntDoc
            objItem.iTipoDoc = objDoc.iTipoDoc
        
            lErro = Comando_Executar(lComando2, "INSERT INTO SpedDocFiscaisItens( NumIntSped, TipoDoc, NumIntDoc, Item, " & _
                    "Registro, Produto, Descricao, Quantidade, UM, " & _
                    "MovFisica, ICMSCST, CFOP, ICMSAliquota, ICMSSTAliquota, " & _
                    "IPICST, IPICodEnq, IPIALiquota, PISCST, PISAliquota, " & _
                    "PISBCQtd, PISAliquotaQtd, CofinsCST, CofinsAliquota, CofinsBCQtd, " & _
                    "CofinsAliquotaQtd, ValorDesconto, ValorFrete, ValorSeguro, ValorDespesas, " & _
                    "ValorServico, ValorNaoTributado, ValorBaseICMS, ValorICMS, ValorBaseICMSST, " & _
                    "ValorICMSST, ValorBaseIPI, ValorIPI, ValorBasePIS, ValorPIS, " & _
                    "ValorBaseCofins, ValorCofins, ValorPisRetido, ValorPisST, ValorCofinsRetido, " & _
                    "ValorCofinsST, ValorISS, PrecoUnitario, PrecoTotal, NatBCCred, ValorRecebido, ValorNoCR, NCM, CodAtividadeTab511, AliquotaContrPrev, ValorRecContrPrev, CodISSServ, ValorBaseISS, AliquotaISS) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    objItem.lNumIntSped, objItem.iTipoDoc, objItem.lNumIntDoc, objItem.iItem, _
                    objItem.sRegistro, objItem.sProduto, objItem.sDescricao, objItem.dQuantidade, objItem.sUM, _
                    objItem.sMovFisica, objItem.sICMSCST, objItem.sCFOP, objItem.dICMSAliquota, objItem.dICMSSTAliquota, _
                    objItem.sIPICST, objItem.sIPICodEnq, objItem.dIPIAliquota, objItem.sPISCST, objItem.dPISAliquota, _
                    objItem.dPISBCQtd, objItem.dPISAliquotaQtd, objItem.sCofinsCST, objItem.dCOFINSAliquota, objItem.dCofinsBCQtd, _
                    objItem.dCofinsAliquotaQtd, objItem.dValorDesconto, objItem.dValorFrete, objItem.dValorSeguro, objItem.dValorDespesas, _
                    objItem.dValorServico, objItem.dValorNaoTributado, objItem.dValorBaseICMS, objItem.dValorICMS, objItem.dValorBaseICMSST, _
                    objItem.dValorICMSST, objItem.dValorBaseIPI, objItem.dValorIPI, objItem.dValorBasePIS, objItem.dValorPIS, _
                    objItem.dValorBaseCofins, objItem.dValorCofins, objItem.dValorPisRetido, objItem.dValorPisST, objItem.dValorCofinsRetido, _
                    objItem.dValorCofinsST, objItem.dValorISS, objItem.dPrecoUnitario, objItem.dPrecoTotal, objItem.sNatBCCred, objItem.dValorRecebido, objItem.dValorNoCR, objItem.sNCM, objItem.sCodAtividadeTab511, objItem.dAliquotaContrPrev, objItem.dValorRecContrPrev, objItem.sCodISSServ, objItem.dValorBaseISS, objItem.dAliquotaISS)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209401
                  
            'lErro = objEFD.objTela.Processa_Item(False)
            'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            objEFD.lNumTotalProcessados = objEFD.lNumTotalProcessados + 1
        
        Next
    End If

    SpedFiscal_Docs_Grava = SUCESSO

    Exit Function

Erro_SpedFiscal_Docs_Grava:

    SpedFiscal_Docs_Grava = gErr

    Select Case gErr
    
        Case 209400
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_SPEDDOCFISCAIS", gErr)
        
        Case 209401
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_SPEDDOCFISCAISITENS", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209402)

    End Select

    Exit Function
    
End Function

Private Function SpedFiscal_Docs_Prepara(ByVal lComando As Long, ByVal lNumIntSped As Long, ByVal sBloco As String, tDoc As typeSpedDocFiscais) As Long

Dim lErro As Long

On Error GoTo Erro_SpedFiscal_Docs_Prepara

    'Alocação de espaço no buffer
    tDoc.sBloco = String(STRING_MAXIMO, 0)
    tDoc.sRegistro = String(STRING_MAXIMO, 0)
    tDoc.sOperacao = String(STRING_MAXIMO, 0)
    tDoc.sEmitente = String(STRING_MAXIMO, 0)
    tDoc.sModelo = String(STRING_MAXIMO, 0)
    tDoc.sSerie = String(STRING_MAXIMO, 0)
    tDoc.sChaveEletronica = String(STRING_MAXIMO, 0)
    tDoc.sFrete = String(STRING_MAXIMO, 0)
    tDoc.sPagamento = String(STRING_MAXIMO, 0)
    tDoc.sCgc = String(STRING_MAXIMO, 0)
    tDoc.sIndNatFrtPis = String(STRING_MAXIMO, 0)
    tDoc.sIndNatFrtCofins = String(STRING_MAXIMO, 0)
    tDoc.sCodMunicServ = String(STRING_MAXIMO, 0)

    'Le a tabelaSpedDocFiscais
    lErro = Comando_Executar(lComando, "SELECT NumIntSped, TipoDoc, NumIntDoc, Bloco, Registro, FilialEmpresa, " & _
                "Cliente, FilialCli, Fornecedor, FilialForn, DataEmissao, DataES, ValorTotal, " & _
                "Operacao, Emitente, Modelo, Situacao, Serie, NumDocumento, ChaveEletronica, " & _
                "Frete, Pagamento, ValorDesconto, ValorFrete, ValorSeguro, ValorDespesas, ValorServico, " & _
                "ValorNaoTributado, ValorBaseICMS, ValorICMS, ValorBaseICMSST, ValorICMSST, ValorBaseIPI, ValorIPI, " & _
                "ValorBasePIS, ValorPIS, ValorBaseCofins, ValorCofins, ValorPisRetido, ValorPisST, ValorCofinsRetido, " & _
                "ValorCofinsST, ValorISS, CGC, CodMensagem, ValorMercadoria, IndNatFrtPis, IndNatFrtCofins, ValorRecebido, ValorNoCR, ValorRecContrPrev, CodMunicServ, ValorBaseISS, ValorBaseISSRet, ValorISSRet FROM SpedDocFiscais WHERE NumIntSped= ? AND Bloco= ? ORDER BY CGC, Registro ", _
                tDoc.lNumIntSped, tDoc.iTipoDoc, tDoc.lNumIntDoc, tDoc.sBloco, _
                tDoc.sRegistro, tDoc.iFilialEmpresa, tDoc.lCliente, tDoc.iFilialCli, tDoc.lFornecedor, _
                tDoc.iFilialForn, tDoc.dtDataEmissao, tDoc.dtDataES, tDoc.dValorTotal, tDoc.sOperacao, _
                tDoc.sEmitente, tDoc.sModelo, tDoc.iSituacao, tDoc.sSerie, tDoc.lNumDocumento, _
                tDoc.sChaveEletronica, tDoc.sFrete, tDoc.sPagamento, tDoc.dValorDesconto, tDoc.dValorFrete, _
                tDoc.dValorSeguro, tDoc.dValorDespesas, tDoc.dValorServico, tDoc.dValorNaoTributado, tDoc.dValorBaseICMS, _
                tDoc.dValorICMS, tDoc.dValorBaseICMSST, tDoc.dValorICMSST, tDoc.dValorBaseIPI, tDoc.dValorIPI, _
                tDoc.dValorBasePIS, tDoc.dValorPIS, tDoc.dValorBaseCofins, tDoc.dValorCofins, tDoc.dValorPisRetido, _
                tDoc.dValorPisST, tDoc.dValorCofinsRetido, tDoc.dValorCofinsST, tDoc.dValorISS, tDoc.sCgc, tDoc.lCodMensagem, tDoc.dValorMercadoria, tDoc.sIndNatFrtPis, tDoc.sIndNatFrtCofins, tDoc.dValorRecebido, tDoc.dValorNoCR, tDoc.dValorRecContrPrev, tDoc.sCodMunicServ, tDoc.dValorBaseISS, tDoc.dValorBaseISSRet, tDoc.dValorISSRet, _
                lNumIntSped, sBloco)
    If lErro <> AD_SQL_SUCESSO Then gError 209403
    
    SpedFiscal_Docs_Prepara = SUCESSO

    Exit Function

Erro_SpedFiscal_Docs_Prepara:

    SpedFiscal_Docs_Prepara = gErr

    Select Case gErr

        Case 209403
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209404)

    End Select

    Exit Function

End Function
    
Private Function SpedFiscal_DocItens_Le(ByVal lComando As Long, ByVal objDoc As ClassSpedDocFiscais) As Long

Dim lErro As Long
Dim tItem As typeSpedDocFiscaisItens
Dim objItem As ClassSpedDocFiscaisItens

On Error GoTo Erro_SpedFiscal_DocItens_Le

    'Alocação de espaço no buffer
    tItem.sRegistro = String(STRING_MAXIMO, 0)
    tItem.sProduto = String(STRING_MAXIMO, 0)
    tItem.sDescricao = String(STRING_MAXIMO, 0)
    tItem.sUM = String(STRING_MAXIMO, 0)
    tItem.sMovFisica = String(STRING_MAXIMO, 0)
    tItem.sICMSCST = String(STRING_MAXIMO, 0)
    tItem.sCFOP = String(STRING_MAXIMO, 0)
    tItem.sIPICST = String(STRING_MAXIMO, 0)
    tItem.sIPICodEnq = String(STRING_MAXIMO, 0)
    tItem.sPISCST = String(STRING_MAXIMO, 0)
    tItem.sCofinsCST = String(STRING_MAXIMO, 0)
    tItem.sNatBCCred = String(STRING_MAXIMO, 0)
    tItem.sNCM = String(STRING_MAXIMO, 0)
    tItem.sCodAtividadeTab511 = String(STRING_MAXIMO, 0)
    tItem.sCodISSServ = String(STRING_MAXIMO, 0)

    'Le a tabelaSpedDocFiscaisItens
    lErro = Comando_Executar(lComando, "SELECT NumIntSped, TipoDoc, NumIntDoc, Item, Registro, Produto, " & _
                "Descricao, Quantidade, PrecoUnitario, PrecoTotal, UM, MovFisica, ICMSCST, CFOP, ICMSAliquota, " & _
                "ICMSSTAliquota, IPICST, IPICodEnq, IPIALiquota, PISCST, PISAliquota, PISBCQtd, " & _
                "PISAliquotaQtd, CofinsCST, CofinsAliquota, CofinsBCQtd, CofinsAliquotaQtd, ValorDesconto, ValorFrete, " & _
                "ValorSeguro, ValorDespesas, ValorServico, ValorNaoTributado, ValorBaseICMS, ValorICMS, ValorBaseICMSST, " & _
                "ValorICMSST, ValorBaseIPI, ValorIPI, ValorBasePIS, ValorPIS, ValorBaseCofins, ValorCofins, " & _
                "ValorPisRetido, ValorPisST, ValorCofinsRetido, ValorCofinsST, ValorISS, NatBCCred, ValorRecebido, ValorNoCR, NCM, CodAtividadeTab511, AliquotaContrPrev, ValorRecContrPrev, CodISSServ, ValorBaseISS, AliquotaISS FROM SpedDocFiscaisItens WHERE NumIntSped= ?  AND TipoDoc= ?  AND NumIntDoc= ? ", _
                tItem.lNumIntSped, tItem.iTipoDoc, tItem.lNumIntDoc, tItem.iItem, _
                tItem.sRegistro, tItem.sProduto, tItem.sDescricao, tItem.dQuantidade, tItem.dPrecoUnitario, tItem.dPrecoTotal, tItem.sUM, _
                tItem.sMovFisica, tItem.sICMSCST, tItem.sCFOP, tItem.dICMSAliquota, tItem.dICMSSTAliquota, _
                tItem.sIPICST, tItem.sIPICodEnq, tItem.dIPIAliquota, tItem.sPISCST, tItem.dPISAliquota, _
                tItem.dPISBCQtd, tItem.dPISAliquotaQtd, tItem.sCofinsCST, tItem.dCOFINSAliquota, tItem.dCofinsBCQtd, _
                tItem.dCofinsAliquotaQtd, tItem.dValorDesconto, tItem.dValorFrete, tItem.dValorSeguro, tItem.dValorDespesas, _
                tItem.dValorServico, tItem.dValorNaoTributado, tItem.dValorBaseICMS, tItem.dValorICMS, tItem.dValorBaseICMSST, _
                tItem.dValorICMSST, tItem.dValorBaseIPI, tItem.dValorIPI, tItem.dValorBasePIS, tItem.dValorPIS, _
                tItem.dValorBaseCofins, tItem.dValorCofins, tItem.dValorPisRetido, tItem.dValorPisST, tItem.dValorCofinsRetido, _
                tItem.dValorCofinsST, tItem.dValorISS, tItem.sNatBCCred, tItem.dValorRecebido, tItem.dValorNoCR, tItem.sNCM, tItem.sCodAtividadeTab511, tItem.dAliquotaContrPrev, tItem.dValorRecContrPrev, tItem.sCodISSServ, tItem.dValorBaseISS, tItem.dAliquotaISS, _
                objDoc.lNumIntSped, objDoc.iTipoDoc, objDoc.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209405

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209406

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItem = New ClassSpedDocFiscaisItens
        
        Call SpedFiscal_DocItem_Copia(objItem, tItem)
        
        objDoc.colItens.Add objItem
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209407
        
    Loop

    SpedFiscal_DocItens_Le = SUCESSO

    Exit Function

Erro_SpedFiscal_DocItens_Le:

    SpedFiscal_DocItens_Le = gErr

    Select Case gErr

        Case 209405 To 209407
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAISITENS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209408)

    End Select

    Exit Function

End Function

Private Function SpedFiscal_Doc_Copia(objDoc As ClassSpedDocFiscais, tDoc As typeSpedDocFiscais) As Long
On Error GoTo Erro_SpedFiscal_Doc_Copia
    objDoc.lNumIntSped = tDoc.lNumIntSped
    objDoc.iTipoDoc = tDoc.iTipoDoc
    objDoc.lNumIntDoc = tDoc.lNumIntDoc
    objDoc.sBloco = tDoc.sBloco
    objDoc.sRegistro = tDoc.sRegistro
    objDoc.iFilialEmpresa = tDoc.iFilialEmpresa
    objDoc.lCliente = tDoc.lCliente
    objDoc.iFilialCli = tDoc.iFilialCli
    objDoc.lFornecedor = tDoc.lFornecedor
    objDoc.iFilialForn = tDoc.iFilialForn
    objDoc.dtDataEmissao = tDoc.dtDataEmissao
    objDoc.dtDataES = tDoc.dtDataES
    objDoc.dValorTotal = tDoc.dValorTotal
    objDoc.sOperacao = tDoc.sOperacao
    objDoc.sEmitente = tDoc.sEmitente
    objDoc.sModelo = tDoc.sModelo
    objDoc.iSituacao = tDoc.iSituacao
    objDoc.sSerie = tDoc.sSerie
    objDoc.lNumDocumento = tDoc.lNumDocumento
    objDoc.sChaveEletronica = tDoc.sChaveEletronica
    objDoc.sFrete = tDoc.sFrete
    objDoc.sPagamento = tDoc.sPagamento
    objDoc.dValorDesconto = tDoc.dValorDesconto
    objDoc.dValorFrete = tDoc.dValorFrete
    objDoc.dValorSeguro = tDoc.dValorSeguro
    objDoc.dValorDespesas = tDoc.dValorDespesas
    objDoc.dValorServico = tDoc.dValorServico
    objDoc.dValorNaoTributado = tDoc.dValorNaoTributado
    objDoc.dValorBaseICMS = tDoc.dValorBaseICMS
    objDoc.dValorICMS = tDoc.dValorICMS
    objDoc.dValorBaseICMSST = tDoc.dValorBaseICMSST
    objDoc.dValorICMSST = tDoc.dValorICMSST
    objDoc.dValorBaseIPI = tDoc.dValorBaseIPI
    objDoc.dValorIPI = tDoc.dValorIPI
    objDoc.dValorBasePIS = tDoc.dValorBasePIS
    objDoc.dValorPIS = tDoc.dValorPIS
    objDoc.dValorBaseCofins = tDoc.dValorBaseCofins
    objDoc.dValorCofins = tDoc.dValorCofins
    objDoc.dValorPisRetido = tDoc.dValorPisRetido
    objDoc.dValorPisST = tDoc.dValorPisST
    objDoc.dValorCofinsRetido = tDoc.dValorCofinsRetido
    objDoc.dValorCofinsST = tDoc.dValorCofinsST
    objDoc.dValorISS = tDoc.dValorISS
    objDoc.sCgc = tDoc.sCgc
    objDoc.lCodMensagem = tDoc.lCodMensagem
    objDoc.dValorMercadoria = tDoc.dValorMercadoria
    objDoc.sIndNatFrtPis = tDoc.sIndNatFrtPis
    objDoc.sIndNatFrtCofins = tDoc.sIndNatFrtCofins
    objDoc.dValorRecebido = tDoc.dValorRecebido
    objDoc.dValorNoCR = tDoc.dValorNoCR
    objDoc.dValorRecContrPrev = tDoc.dValorRecContrPrev
    
    objDoc.sCodMunicServ = tDoc.sCodMunicServ
    objDoc.dValorBaseISS = tDoc.dValorBaseISS
    objDoc.dValorBaseISSRet = tDoc.dValorBaseISSRet
    objDoc.dValorISSRet = tDoc.dValorISSRet
    
    SpedFiscal_Doc_Copia = SUCESSO
    Exit Function
Erro_SpedFiscal_Doc_Copia:
    SpedFiscal_Doc_Copia = gErr
    Select Case gErr
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209409)
    End Select
    Exit Function
End Function

Private Function SpedFiscal_DocItem_Copia(objItem As ClassSpedDocFiscaisItens, tItem As typeSpedDocFiscaisItens) As Long

On Error GoTo Erro_SpedFiscal_DocItem_Copia
    objItem.lNumIntSped = tItem.lNumIntSped
    objItem.iTipoDoc = tItem.iTipoDoc
    objItem.lNumIntDoc = tItem.lNumIntDoc
    objItem.iItem = tItem.iItem
    objItem.sRegistro = tItem.sRegistro
    objItem.sProduto = tItem.sProduto
    objItem.sDescricao = tItem.sDescricao
    objItem.dQuantidade = tItem.dQuantidade
    objItem.sUM = tItem.sUM
    objItem.sMovFisica = tItem.sMovFisica
    objItem.sICMSCST = tItem.sICMSCST
    objItem.sCFOP = tItem.sCFOP
    objItem.dICMSAliquota = tItem.dICMSAliquota
    objItem.dICMSSTAliquota = tItem.dICMSSTAliquota
    objItem.sIPICST = tItem.sIPICST
    objItem.sIPICodEnq = tItem.sIPICodEnq
    objItem.dIPIAliquota = tItem.dIPIAliquota
    objItem.sPISCST = tItem.sPISCST
    objItem.dPISAliquota = tItem.dPISAliquota
    objItem.dPISBCQtd = tItem.dPISBCQtd
    objItem.dPISAliquotaQtd = tItem.dPISAliquotaQtd
    objItem.sCofinsCST = tItem.sCofinsCST
    objItem.dCOFINSAliquota = tItem.dCOFINSAliquota
    objItem.dCofinsBCQtd = tItem.dCofinsBCQtd
    objItem.dCofinsAliquotaQtd = tItem.dCofinsAliquotaQtd
    objItem.dValorDesconto = tItem.dValorDesconto
    objItem.dValorFrete = tItem.dValorFrete
    objItem.dValorSeguro = tItem.dValorSeguro
    objItem.dValorDespesas = tItem.dValorDespesas
    objItem.dValorServico = tItem.dValorServico
    objItem.dValorNaoTributado = tItem.dValorNaoTributado
    objItem.dValorBaseICMS = tItem.dValorBaseICMS
    objItem.dValorICMS = tItem.dValorICMS
    objItem.dValorBaseICMSST = tItem.dValorBaseICMSST
    objItem.dValorICMSST = tItem.dValorICMSST
    objItem.dValorBaseIPI = tItem.dValorBaseIPI
    objItem.dValorIPI = tItem.dValorIPI
    objItem.dValorBasePIS = tItem.dValorBasePIS
    objItem.dValorPIS = tItem.dValorPIS
    objItem.dValorBaseCofins = tItem.dValorBaseCofins
    objItem.dValorCofins = tItem.dValorCofins
    objItem.dValorPisRetido = tItem.dValorPisRetido
    objItem.dValorPisST = tItem.dValorPisST
    objItem.dValorCofinsRetido = tItem.dValorCofinsRetido
    objItem.dValorCofinsST = tItem.dValorCofinsST
    objItem.dValorISS = tItem.dValorISS
    objItem.dPrecoUnitario = tItem.dPrecoUnitario
    objItem.dPrecoTotal = tItem.dPrecoTotal
    objItem.sNatBCCred = tItem.sNatBCCred
    objItem.dValorRecebido = tItem.dValorRecebido
    objItem.dValorNoCR = tItem.dValorNoCR
    objItem.sNCM = tItem.sNCM
    objItem.sCodAtividadeTab511 = tItem.sCodAtividadeTab511
    objItem.dAliquotaContrPrev = tItem.dAliquotaContrPrev
    objItem.dValorRecContrPrev = tItem.dValorRecContrPrev
    
    objItem.sCodISSServ = tItem.sCodISSServ
    objItem.dValorBaseISS = tItem.dValorBaseISS
    objItem.dAliquotaISS = tItem.dAliquotaISS
    
    SpedFiscal_DocItem_Copia = SUCESSO
    Exit Function
Erro_SpedFiscal_DocItem_Copia:
    SpedFiscal_DocItem_Copia = gErr
    Select Case gErr
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209410)
    End Select
    Exit Function
End Function

Private Function SpedFiscal_Gravar_BlocoB(ByVal objEFD As ClassEFDPisCofinsSel, lFile As Long) As Long

Dim lErro As Long, lErro1 As Long
Dim sRegistro As String, lSeq As Long
Dim sRegistroAux As String, lSeqAnt As Long
Dim lDifQtdRegs As Long, lIndice As Long

On Error GoTo Erro_SpedFiscal_Gravar_BlocoB

'    'A previsão de falta é o que estava inicialmente previsto menos o que já foi processado
'    'O real de falta é q quantidade total de registro que vai ser colocada no arquivo texto
'    lDifQtdRegs = (objEFD.lNumTotalPrevisto - objEFD.lNumTotalProcessados) - objEFD.lNumTotalReal
'    'Se estava previsto mais registros do que realmente falta avança a barra de progressão
'    If lDifQtdRegs > 0 Then
'        For lIndice = 1 To lDifQtdRegs
'            lErro = objEFD.objTela.Processa_Item(False) 'False significa que não está processando o item, só processa quando gera o registro no BD
'            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'        Next
'    End If

    'Open objEFD.sDiretorio For Output As #1

    sRegistroAux = String(STRING_MAXIMO, 0)
    sRegistro = ""

    lErro = Comando_Executar(objEFD.ProxComando, "SELECT R.Seq, R.Registro FROM SpedFiscalPis AS S, SpedFiscalPisReg AS R WHERE S.NumIntSped = R.NumIntSped AND S.Seq = R.Seq AND S.NumIntSped = ? ORDER BY R.Seq, R.SeqReg ", _
    lSeq, sRegistroAux, objEFD.lNumIntSped)
    If lErro <> AD_SQL_SUCESSO Then gError 209411

    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209412

    lSeqAnt = lSeq

    Do While lErro = AD_SQL_SUCESSO

        sRegistro = sRegistro + sRegistroAux

        lErro = Comando_BuscarProximo(objEFD.UltimoCmdUsado)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209413

        If lSeqAnt <> lSeq Or lErro = AD_SQL_SEM_DADOS Then
            lSeqAnt = lSeq
            Print #lFile, sRegistro
            'If lDifQtdRegs >= 0 Then
            '    lErro1 = objEFD.objTela.Processa_Item(False)
            '    If lErro1 <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            'Else
                'Se falta mais do que estava previsto não mexe na barra de progressão até falta igual
            '    lDifQtdRegs = lDifQtdRegs + 1
            'End If
            sRegistro = ""
        End If

    Loop

    'Close #1

    SpedFiscal_Gravar_BlocoB = SUCESSO

    Exit Function

Erro_SpedFiscal_Gravar_BlocoB:

    SpedFiscal_Gravar_BlocoB = gErr

   Select Case gErr

        Case 209411, 209412, 209413
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDFISCAL", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209414)

    End Select

    'Close #1

End Function

Private Function SpedFiscal_Trata_ValorRecebido(ByVal objDoc As ClassSpedDocFiscais, ByVal objEFD As ClassEFDPisCofinsSel) As Long

Dim lErro As Long
Dim objItem As ClassSpedDocFiscaisItens
Dim dValorNoCRItens1 As Double
Dim dValorNoCRItens2 As Double
Dim dValorRecebidoItens As Double
Dim dDif As Double, dFator As Double, dSoma As Double

On Error GoTo Erro_SpedFiscal_Trata_ValorRecebido

    'Rateia o valor a receber por item para poder quebrar por CST caso tenham CSTs diferentes
    dValorNoCRItens1 = 0
    For Each objItem In objDoc.colItens
        objItem.dValorNoCR = objItem.dPrecoTotal + objItem.dValorFrete + objItem.dValorDespesas + objItem.dValorSeguro + objItem.dValorICMSST + objItem.dValorIPI  'Teria que ratear IR, ver se o item é cobrado, etc, mas para efeito de cálculo inicial será considerado apenas valores básicos e depois será ajustado de acordo com o a receber gerado
        dValorNoCRItens1 = dValorNoCRItens1 + objItem.dValorNoCR
    Next
    dValorNoCRItens2 = dValorNoCRItens1
    
    If Abs(objDoc.dValorNoCR - dValorNoCRItens1) > DELTA_VALORMONETARIO Then
        dDif = objDoc.dValorNoCR - dValorNoCRItens1
        dValorNoCRItens2 = 0
        For Each objItem In objDoc.colItens
            objItem.dValorNoCR = objItem.dValorNoCR + Arredonda_Moeda(objItem.dValorNoCR * (dDif / dValorNoCRItens1))
            dValorNoCRItens2 = dValorNoCRItens2 + objItem.dValorNoCR
        Next
    End If

    If Abs(objDoc.dValorNoCR - dValorNoCRItens2) > DELTA_VALORMONETARIO Then
        dDif = objDoc.dValorNoCR - dValorNoCRItens2
        For Each objItem In objDoc.colItens
            objItem.dValorNoCR = objItem.dValorNoCR + dDif
            Exit For
        Next
    End If
    
    'Distribui o valor recebido por item
    dValorRecebidoItens = 0
    dFator = objDoc.dValorRecebido / objDoc.dValorNoCR
    For Each objItem In objDoc.colItens
        objItem.dValorRecebido = Arredonda_Moeda(objItem.dValorNoCR * dFator)
        dValorRecebidoItens = dValorRecebidoItens + objItem.dValorRecebido
    Next
    
    If Abs(objDoc.dValorRecebido - dValorRecebidoItens) > DELTA_VALORMONETARIO Then
        dDif = objDoc.dValorRecebido - dValorRecebidoItens
        For Each objItem In objDoc.colItens
            objItem.dValorRecebido = objItem.dValorRecebido + dDif
            Exit For
        Next
    End If

    'Calcula o valor recebido sujeito a contribuição previdenciária
    dSoma = 0
    For Each objItem In objDoc.colItens
        If Len(Trim(objItem.sCodAtividadeTab511)) > 0 Then
            objItem.dValorRecContrPrev = objItem.dValorRecebido
            dSoma = dSoma + objItem.dValorRecContrPrev
        End If
    Next
    objDoc.dValorRecContrPrev = dSoma
    
    SpedFiscal_Trata_ValorRecebido = SUCESSO

    Exit Function

Erro_SpedFiscal_Trata_ValorRecebido:

    SpedFiscal_Trata_ValorRecebido = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209402)

    End Select

    Exit Function
    
End Function

Function SFF_Texto(ByVal sTexto As String, Optional ByVal lTamMax As Long = 0, Optional bFazTrim As Boolean = True, Optional bRemoveEspacoDuplo As Boolean = True, Optional bRemoveAcentos As Boolean = False, Optional bUpperCase As Boolean = False) As String
Dim sTextoAux As String
Dim iIndice As Integer
Dim iPos As Integer
'1 Formato dos campos
'ALFANUMÉRICO(C): representados por “C” - todos os caracteres, excetuados os caracteres
'“|”(Pipe ou Barra Vertical: caractere 124 da Tabela ASCII) e os não-imprimíveis (caracteres 00 a 31 da Tabela ASCII);
'NUMÉRICO(N): representados por “N” - algarismos das posições 'de 48 a 57 da Tabela ASCII.

    sTextoAux = sTexto
    
    sTextoAux = Replace(sTextoAux, Chr(124), "") 'Remove pipe "|"
    
    For iIndice = 0 To 31 'Remove caracteres não-imprimíveis
        sTextoAux = Replace(sTextoAux, Chr(iIndice), "")
    Next
    
    If bFazTrim Then 'Remove espaços no início e no final (OPCIONAL)
        sTextoAux = Trim(sTextoAux)
    End If
    
    If bRemoveEspacoDuplo Then 'Remove espaços duplos "  " (OPCIONAL)
        iPos = InStr(1, sTextoAux, "  ")
        Do While iPos <> 0
            sTextoAux = Replace(sTextoAux, "  ", " ")
            iPos = InStr(1, sTextoAux, "  ")
        Loop
    End If
    
    If bRemoveAcentos Then 'Tira acentuação (OPCIONAL)
        sTextoAux = DesacentuaTexto(sTextoAux)
    End If
    
    If bUpperCase Then 'Muda para maiúsculo (OPCIONAL)
        sTextoAux = UCase(sTextoAux)
    End If
    
    If lTamMax > 0 Then 'Reduz o tamanho do texto (OPCIONAL)
        sTextoAux = left(sTextoAux, lTamMax)
    End If
    
    SFF_Texto = sTextoAux

    'SFF_Texto = left(Replace(Replace(Replace(sTexto, vbTab, " "), Chr(13), ""), Chr(10), ""), lTamMax)
End Function

Function SFF_Data(ByVal dtData As Date) As String
    If dtData <> DATA_NULA Then
        SFF_Data = Format(dtData, "ddmmyyyy")
    Else
        SFF_Data = ""
    End If
End Function

Function SFF_Num(ByVal dNum As Double, Optional ByVal iDec As Integer = 2) As String
Dim sFormato As String
    sFormato = "###0"
    If iDec > 0 Then sFormato = sFormato & "." & String(iDec, "0")
    SFF_Num = Format(dNum, sFormato)
End Function

Function SFF_Perc(ByVal dPerc As Double, Optional ByVal iDec As Integer = 2) As String
Dim sFormato As String
    sFormato = "###0"
    If iDec > 0 Then sFormato = sFormato & "." & String(iDec, "0")
    SFF_Perc = Format(dPerc * 100, sFormato)
End Function
'
'Private Function SpedFiscal_K001(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroK001 As ClassSpedRegistro) As Long
'
'Dim lErro As Long
'Dim objSpedRegistro As ClassSpedRegistro
'Dim alComando(1 To 1) As Long
'Dim iIndice As Integer
'Dim iIndMov As Integer
'
'
'On Error GoTo Erro_SpedFiscal_K001
'
'    'Abertura de Comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 204075
'    Next
'
'    iIndMov = 1 '??? sem movimento
'
'    objSpedFiscal.sTipo = "K001"
'    objSpedRegistroK001.sTipo = objSpedFiscal.sTipo
'
'    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
'    objSpedRegistroK001.iCount = objSpedRegistroK001.iCount + 1
'
'    objSpedFiscal.sRegistro = "|K001|" & iIndMov & "|"
'
'    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
'        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
'    If lErro <> AD_SQL_SUCESSO Then gError 204077
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'         Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    SpedFiscal_K001 = SUCESSO
'
'    Exit Function
'
'Erro_SpedFiscal_K001:
'
'   SpedFiscal_K001 = gErr
'
'   Select Case gErr
'
'        Case 204075
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 204077
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204078)
'
'    End Select
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'         Call Comando_Fechar(alComando(iIndice))
'    Next
'
'End Function
'
'Function SpedFiscal_K990(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroK990 As ClassSpedRegistro, ByVal objSpedRegistroK001 As ClassSpedRegistro) As Long
'
'Dim lErro As Long
'Dim objSpedRegistro As ClassSpedRegistro
'Dim alComando(1 To 1) As Long
'Dim iIndice As Integer
'Dim lTotal As Long
'
'
'On Error GoTo Erro_SpedFiscal_K990
'
'    'Abertura de Comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 204083
'    Next
'
'
'    objSpedFiscal.sTipo = "K990"
'    objSpedRegistroK990.sTipo = objSpedFiscal.sTipo
'
'    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
'    objSpedRegistroK990.iCount = objSpedRegistroK990.iCount + 1
'
'    lTotal = objSpedRegistroK001.iCount + objSpedRegistroK990.iCount
'
'    objSpedFiscal.sRegistro = "|K990|" & lTotal & "|"
'
'    If Len(objSpedFiscal.sRegistro) > 500 Then gError 204084
'
'    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
'    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
'    If lErro <> AD_SQL_SUCESSO Then gError 204085
'
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'         Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    SpedFiscal_K990 = SUCESSO
'
'    Exit Function
'
'Erro_SpedFiscal_K990:
'
'   SpedFiscal_K990 = gErr
'
'   Select Case gErr
'
'        Case 204083
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 204084
'            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
'
'        Case 204085
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204086)
'
'    End Select
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'         Call Comando_Fechar(alComando(iIndice))
'    Next
'
'End Function

Function SpedFiscal_K001(ByVal objFisSpedAux As ClassFISSpedAux) As Long
'REGISTRO K001: ABERTURA DO BLOCO K
'Este registro deve ser gerado para abertura do bloco K, indicando se há registros de informações no bloco.

Dim lErro As Long
Dim objSpedFiscal As ClassSpedFiscal
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lTotal As Long

On Error GoTo Erro_SpedFiscal_K001
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216248
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal

    objSpedFiscal.sTipo = "K001"
    objFisSpedAux.objSpedRegistroK001.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objFisSpedAux.objSpedRegistroK001.iCount = objFisSpedAux.objSpedRegistroK001.iCount + 1
    
    '0- Bloco com dados informados;
    '1- Bloco sem dados informados
    If objFisSpedAux.iIncluiRCPE = MARCADO Then
        objSpedFiscal.sRegistro = "|K001|0|"
    Else
        objSpedFiscal.sRegistro = "|K001|1|"
    End If
    If Len(objSpedFiscal.sRegistro) > 500 Then gError 216249
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 216250

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_K001 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_K001:

   SpedFiscal_K001 = gErr

   Select Case gErr
    
        Case 216248
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 216249
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 216250
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216251)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_K100(ByVal objFisSpedAux As ClassFISSpedAux) As Long
'REGISTRO K100: PERÍODO DE APURAÇÃO DO ICMS/IPI
'Este registro tem o objetivo de informar o período de apuração do ICMS ou do IPI, prevalecendo os períodos
'mais curtos. Contribuintes com mais de um período de apuração no mês declaram um registro K100 para cada período no
'mesmo arquivo. Não podem ser informados dois ou mais registros com os mesmos campos DT_INI e DT_FIN.
'Os períodos informados neste registro deverão abranger todo o período da escrituração, conforme informado no
'Registro 0000.

Dim lErro As Long
Dim objSpedFiscal  As ClassSpedFiscal
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lTotal As Long

On Error GoTo Erro_SpedFiscal_K100
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216251
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal

    objSpedFiscal.sTipo = "K100"
    objFisSpedAux.objSpedRegistroK100.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objFisSpedAux.objSpedRegistroK100.iCount = objFisSpedAux.objSpedRegistroK100.iCount + 1

    objSpedFiscal.sRegistro = "|K100|" & Format(objFisSpedAux.dtDataIni, "ddmmyyyy") & "|" & Format(objFisSpedAux.dtDataFim, "ddmmyyyy") & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 216252
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 216253

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_K100 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_K100:

   SpedFiscal_K100 = gErr

   Select Case gErr
    
        Case 216251
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 216252
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 216253
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204090)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_K200(ByVal objFisSpedAux As ClassFISSpedAux) As Long
'REGISTRO K200: ESTOQUE ESCRITURADO
'Este registro tem o objetivo de informar o estoque final escriturado do período de apuração informado no Registro
'K100, por tipo de estoque e por participante, nos casos em que couber, das mercadorias de tipos 00 – Mercadoria para
'revenda, 01 – Matéria-Prima, 02 - Embalagem, 03 – Produtos em Processo, 04 – Produto Acabado, 05 – Subproduto, 06 –
'Produto Intermediário e 10 – Outros Insumos – campo TIPO_ITEM do Registro 0200.
'A informação de estoque zero (quantidade = 0) não deixa de ser uma informação e o PVA não a impede.
'Entretanto, caso não seja prestada essa informação, será considerado que o estoque é igual a zero. Portanto, é desnecessária
'a informação de estoque zero caso não exista quantidade em estoque, independentemente de ter havido movimentação.
'A quantidade em estoque deve ser expressa, obrigatoriamente, na unidade de medida de controle de estoque
'constante no campo 06 do registro 0200 –UNID_INV.
'A chave deste registro são os campos: DT_EST, COD_ITEM, IND_EST e COD_PART (este, quando houver).
'O estoque escriturado informado no Registro K200 deve refletir a quantidade existente na data final do período de
'apuração informado no Registro K100, estoque este derivado dos apontamentos de estoque inicial / entrada / produção
'/consumo / saída / movimentação interna. Considerando isso, o estoque escriturado informado no K200 é resultante da
'seguinte fórmula:
'Estoque final = estoque inicial + entradas/produção/movimentação interna – Saída / consumo /movimentação
'interna.
'Os estabelecimentos equiparados a industriais e atacadistas devem informar o estoque escriturado – K200 - e,
'caso ocorram movimentações internas, o K220.

Dim lErro As Long
Dim iIndice As Integer
Dim sUMResult As String
Dim alComando(0 To 25) As Long
Dim lCliente As Long
Dim iFilialCli As Integer
Dim lFornecedor As Long
Dim iFilialForn As Integer
Dim dQuantInicial3 As Double
Dim dQuantInicial As Double
Dim sCodParticipante As String
Dim dtDataInv As Date
Dim iTipoReg As Integer, sProd As String, sUM As String, dQuantidade As Double
Dim dQtdeTotalTerc As Double, iClasseUM As Integer
Dim objEstTerc As ClassEstoqueTerc
Dim objSpedFiscal As ClassSpedFiscal

On Error GoTo Erro_SpedFiscal_K200
    
    dtDataInv = objFisSpedAux.dtDataFim

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216254
    Next
    
    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
       
    sUM = String(STRING_UM_SIGLA, 0)
    sProd = String(STRING_PRODUTO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT I.TipoRegInventario, I.Produto, I.SiglaUmEstoque, I.Quantidade, P.ClasseUM FROM RegInventarioRel AS I, Produtos AS P WHERE P.Codigo = I.Produto AND I.Data = ? AND I.FilialEmpresa = ? AND I.Quantidade > ? AND P.Natureza IN (?,?,?,?,?,?,?)  ", iTipoReg, sProd, sUM, dQuantidade, iClasseUM, dtDataInv, objFisSpedAux.iFilialEmpresa, QTDE_ESTOQUE_DELTA2, NATUREZA_PROD_PRODUTO_REVENDA, NATUREZA_PROD_MATERIA_PRIMA, NATUREZA_PROD_EMBALAGENS, NATUREZA_PROD_PRODUTO_EM_PROCESSO, NATUREZA_PROD_PRODUTO_ACABADO, NATUREZA_PROD_SUBPRODUTO, NATUREZA_PROD_OUTROS_INSUMOS)
    If lErro <> AD_SQL_SUCESSO Then gError 216255

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216256
    
    If lErro = AD_SQL_SEM_DADOS Then gError 216257

    Do While lErro = AD_SQL_SUCESSO
    
        lErro = SpedFiscal_0200(sProd, objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, 0, objFisSpedAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = SpedFiscal_0190(sUM, iClasseUM, objSpedFiscal, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, sUMResult, sProd, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        objSpedFiscal.sTipo = "K200"
        objFisSpedAux.objSpedRegistroK200.sTipo = objSpedFiscal.sTipo
    
'        Indicador de propriedade/posse do item:
'        0- Item de propriedade do informante e em seu poder;
'        1- Item de propriedade do informante em posse de terceiros;
'        2- Item de propriedade de terceiros em posse do informante
    
        Select Case iTipoReg
        
            Case 1 'Saldo Nosso na Própria Empresa
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objFisSpedAux.objSpedRegistroK200.iCount = objFisSpedAux.objSpedRegistroK200.iCount + 1
                
                If Abs(dQuantidade) < QTDE_ESTOQUE_DELTA Then
                    dQuantidade = 0
                End If
                If dQuantidade < 0 Then
                    dQuantidade = 0
                End If
                
                objSpedFiscal.sRegistro = "|K200|" & Format(dtDataInv, "ddmmyyyy") & "|" & SFF_Texto(sProd) & "|" & Format(dQuantidade, "###########0.000") & "|0||"
                
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 216258
        
                lErro = Comando_Executar(alComando(2), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
                If lErro <> AD_SQL_SUCESSO Then gError 216259
            
            Case 2 'Saldo Nosso em Poder de Terceiros
            
                If Abs(dQuantidade) < QTDE_ESTOQUE_DELTA Then
                    dQuantidade = 0
                End If
                If dQuantidade < 0 Then
                    dQuantidade = 0
                End If
                
                dQtdeTotalTerc = 0

                lErro = Comando_Executar(alComando(1), "SELECT Cliente, FilialCli, Fornecedor, FilialForn FROM SldMesEst2Terc WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? AND ((Cliente <> 0 OR Fornecedor <> 0) AND (FilialCli <> 0 OR FilialForn <> 0))", _
                    lCliente, iFilialCli, lFornecedor, iFilialForn, Year(dtDataInv), objFisSpedAux.iFilialEmpresa, sProd)
                If lErro <> AD_SQL_SUCESSO Then gError 216260
        
                lErro = Comando_BuscarPrimeiro(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216261
        
                Do While lErro = AD_SQL_SUCESSO
                
                    dQuantInicial = 0
                
                    For iIndice = 2 To 6
                    
                        Set objEstTerc = New ClassEstoqueTerc
                        
                        objEstTerc.dtData = dtDataInv
                        objEstTerc.iEscaninho = iIndice '2 = Conserto; 3 = Consig; 4 = Demo; 5 = Outros; 6 = Beneficiamento
                        objEstTerc.iFilialEmpresa = objFisSpedAux.iFilialEmpresa
                        objEstTerc.sProduto = sProd
                        
                        If lCliente > 0 Then
                            objEstTerc.iTipoCliForn = TIPO_TERC_CLIENTE
                            objEstTerc.lCliForn = lCliente
                            objEstTerc.iFilial = iFilialCli
                        Else
                            objEstTerc.iTipoCliForn = TIPO_TERC_FORNECEDOR
                            objEstTerc.lCliForn = lFornecedor
                            objEstTerc.iFilial = iFilialForn
                        End If
                        
                        lErro = CF("EstoqueTerc_Le_Saldo", objEstTerc, alComando(5 + (iIndice - 2)), alComando(10 + (iIndice - 2)))
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                        dQuantInicial = dQuantInicial + objEstTerc.dQuantidade
                        
                    Next
        
                    'se é um totalizador de cliente
                    If lCliente > 0 Then
                        lErro = SpedFiscal_0150_Cliente(lCliente, iFilialCli, objSpedFiscal, objFisSpedAux.objSpedRegistro0150, objFisSpedAux.objSpedRegistro0175, sCodParticipante, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    Else
                        lErro = SpedFiscal_0150_Fornecedor(lFornecedor, iFilialForn, objSpedFiscal, objFisSpedAux.objSpedRegistro0150, objFisSpedAux.objSpedRegistro0175, sCodParticipante, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    End If
                    objSpedFiscal.sTipo = "K200"
                    
                    If dQuantInicial < -QTDE_ESTOQUE_DELTA Then gError 216262 'Não pode ter quantidade negativa
                    If Abs(dQuantInicial) < QTDE_ESTOQUE_DELTA Then
                        dQuantInicial = 0
                    End If
                    If dQuantInicial < 0 Then
                        dQuantInicial = 0
                    End If
                    
                    dQtdeTotalTerc = dQtdeTotalTerc + dQuantInicial
        
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objFisSpedAux.objSpedRegistroK200.iCount = objFisSpedAux.objSpedRegistroK200.iCount + 1
                            
                    objSpedFiscal.sRegistro = "|K200|" & Format(dtDataInv, "ddmmyyyy") & "|" & SFF_Texto(sProd) & "|" & Format(dQuantInicial, "###########0.000") & "|1|" & sCodParticipante & "|"

                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 216263
        
                    lErro = Comando_Executar(alComando(2), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
                    If lErro <> AD_SQL_SUCESSO Then gError 216264
        
                    lErro = Comando_BuscarProximo(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216265
                
                Loop
                
                If Abs(dQtdeTotalTerc - dQuantidade) > QTDE_ESTOQUE_DELTA Then gError 216266 ' Nem toda quantidade de terceiro está com o terceiro especificado
        
            Case 3 'Saldo de Terceiros em Nosso Poder
            
                If Abs(dQuantidade) < QTDE_ESTOQUE_DELTA Then
                    dQuantidade = 0
                End If
                If dQuantidade < 0 Then
                    dQuantidade = 0
                End If
            
                dQtdeTotalTerc = 0
        
                 lErro = Comando_Executar(alComando(3), "SELECT Cliente, FilialCli, Fornecedor, FilialForn FROM SldMesEst1Terc WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? AND ((Cliente <> 0 OR Fornecedor <> 0) AND (FilialCli <> 0 OR FilialForn <> 0))", _
                    lCliente, iFilialCli, lFornecedor, iFilialForn, Year(dtDataInv), objFisSpedAux.iFilialEmpresa, sProd)
                If lErro <> AD_SQL_SUCESSO Then gError 216267
        
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216268
        
                Do While lErro = AD_SQL_SUCESSO
                
                    dQuantInicial3 = 0
                
                    For iIndice = 7 To 11
                    
                        Set objEstTerc = New ClassEstoqueTerc
                        
                        objEstTerc.dtData = dtDataInv
                        objEstTerc.iEscaninho = iIndice '7 = Conserto; 8 = Consig; 9 = Demo; 10 = Outros;  11 = Beneficiamento
                        objEstTerc.iFilialEmpresa = objFisSpedAux.iFilialEmpresa
                        objEstTerc.sProduto = sProd
                        
                        If lCliente > 0 Then
                            objEstTerc.iTipoCliForn = TIPO_TERC_CLIENTE
                            objEstTerc.lCliForn = lCliente
                            objEstTerc.iFilial = iFilialCli
                        Else
                            objEstTerc.iTipoCliForn = TIPO_TERC_FORNECEDOR
                            objEstTerc.lCliForn = lFornecedor
                            objEstTerc.iFilial = iFilialForn
                        End If
                        
                        lErro = CF("EstoqueTerc_Le_Saldo", objEstTerc, alComando(15 + (iIndice - 7)), alComando(19 + (iIndice - 7)))
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                        dQuantInicial3 = dQuantInicial3 + objEstTerc.dQuantidade
                        
                    Next
        
                    'se é um totalizador de cliente
                    If lCliente > 0 Then
                        lErro = SpedFiscal_0150_Cliente(lCliente, iFilialCli, objSpedFiscal, objFisSpedAux.objSpedRegistro0150, objFisSpedAux.objSpedRegistro0175, sCodParticipante, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    Else
                        lErro = SpedFiscal_0150_Fornecedor(lFornecedor, iFilialForn, objSpedFiscal, objFisSpedAux.objSpedRegistro0150, objFisSpedAux.objSpedRegistro0175, sCodParticipante, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    End If
                    objSpedFiscal.sTipo = "K200"
                    
                    If dQuantInicial3 < -QTDE_ESTOQUE_DELTA Then gError 216269 'Não pode ter quantidade negativa
                    If Abs(dQuantInicial3) < QTDE_ESTOQUE_DELTA Then
                        dQuantInicial3 = 0
                    End If
                    If dQuantInicial3 < 0 Then
                        dQuantInicial3 = 0
                    End If
                    
                    dQtdeTotalTerc = dQtdeTotalTerc + dQuantInicial3
        
                    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                    objFisSpedAux.objSpedRegistroK200.iCount = objFisSpedAux.objSpedRegistroK200.iCount + 1
        
                    objSpedFiscal.sRegistro = "|K200|" & Format(dtDataInv, "ddmmyyyy") & "|" & SFF_Texto(sProd) & "|" & Format(dQuantInicial, "###########0.000") & "|2|" & sCodParticipante & "|"
        
                    If Len(objSpedFiscal.sRegistro) > 500 Then gError 216270
        
                    lErro = Comando_Executar(alComando(2), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro) VALUES (?,?,?,?,?)", _
                    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
                    If lErro <> AD_SQL_SUCESSO Then gError 216271
        
                    lErro = Comando_BuscarProximo(alComando(3))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216272
                
                Loop
                
                If Abs(dQtdeTotalTerc - dQuantidade) > QTDE_ESTOQUE_DELTA Then gError 216273 ' Nem toda quantidade de terceiro está com o terceiro especificado
            
        End Select
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216274
    
    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
        
    SpedFiscal_K200 = SUCESSO

    Exit Function

Erro_SpedFiscal_K200:

    SpedFiscal_K200 = gErr

    Select Case gErr
        
        Case 216254
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 216255, 216256, 216274
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST1", gErr, Year(dtDataInv), objFisSpedAux.iFilialEmpresa)
        
        Case 216257
            Call Rotina_Erro(vbOKOnly, "ERRO_REGINVENTARIO_NAO_CADASTRADO1", gErr, Format(dtDataInv, "dd/mm/yyyy"))
        
        Case 216258, 216263, 216270
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 216259, 216264, 216271
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case 216267, 216268, 216272
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTTERC1", gErr)
        
        Case 216260, 216261, 216265
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTTERC2", gErr)
        
        Case 216262
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST2TERC_NEGATIVO", gErr, Format(dtDataInv, "dd/mm/yyyy"), objFisSpedAux.iFilialEmpresa, sProd, IIf(lCliente > 0, "Cliente", "Fornecedor"), IIf(lCliente > 0, lCliente, lFornecedor), IIf(lCliente > 0, iFilialCli, iFilialForn))
            Resume Next
        
        Case 216266
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDTERC2_DIF_QTDE_REGINV", gErr, Format(dtDataInv, "dd/mm/yyyy"), objFisSpedAux.iFilialEmpresa, sProd, Formata_Estoque(dQtdeTotalTerc), Formata_Estoque(dQuantidade))
            Resume Next
        
        Case 216269
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST1TERC_NEGATIVO", gErr, Format(dtDataInv, "dd/mm/yyyy"), objFisSpedAux.iFilialEmpresa, sProd, IIf(lCliente > 0, "Cliente", "Fornecedor"), IIf(lCliente > 0, lCliente, lFornecedor), IIf(lCliente > 0, iFilialCli, iFilialForn))
            Resume Next
        
        Case 216273
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDTERC1_DIF_QTDE_REGINV", gErr, Format(dtDataInv, "dd/mm/yyyy"), objFisSpedAux.iFilialEmpresa, sProd, Formata_Estoque(dQtdeTotalTerc), Formata_Estoque(dQuantidade))
            Resume Next
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216275)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function SpedFiscal_K230(ByVal objFisSpedAux As ClassFISSpedAux) As Long
'REGISTRO K230: ITENS PRODUZIDOS
'Este registro tem o objetivo de informar a produção acabada de produto em processo (tipo 03 – campo
'TIPO_ITEM do registro 0200) e produto acabado (tipo 04 – campo TIPO_ITEM do registro 0200), inclusive daquele
'industrializado para terceiro por encomenda. O produto resultante é classificado como tipo 03 – produto em processo,
'quando não estiver pronto para ser comercializado, mas estiver pronto para ser consumido em outra fase de produção. O
'produto resultante é classificado como tipo 04 – produto acabado, quando estiver pronto para ser comercializado.
'Deverá existir mesmo que a quantidade de produção acabada seja igual a zero, nas situações em que exista o
'consumo de item componente/insumo no registro filho K235. Nessa situação a produção ficou em elaboração. Essa
'produção em elaboração não é quantificada, uma vez que a matéria não é mais um insumo e nem é ainda um produto
'resultante.
'Devem ser informadas:
'a) as OP iniciadas e concluídas no período de apuração (K100);
'b) as OP iniciadas e não concluídas no período de apuração (OP em que a produção ficou em elaboração), em que
'haja informação de produção e/ou consumo de insumos (K235);
'c) as OP iniciadas em período anterior e concluídas no período de apuração;
'd) as OP iniciadas em período anterior e não concluídas no período de apuração, em que haja informação de
'produção e/ou consumo de insumos (K235).
' Quando a informação for por período de apuração (K100), o K230 somente deve ser informado caso ocorra
'produção no período, com o respectivo consumo de insumos no K235 para se ter essa produção, uma vez que não se teria
'como vincular a quantidade consumida de insumos com a quantidade produzida do produto resultante envolvendo mais de
'um período de apuração. Somente podemos ter produção igual a zero no K230 quando a informação for por ordem de
'produção e quando essa OP não for concluída até a data final do período de apuração do K100 e quando houver o
'apontamento de consumo de insumos no K235.
'A ordem de produção que não for finalizada no período de apuração deve informar a data de conclusão da ordem
'de produção em branco, campo 03 – DT_FIN_OP. No período seguinte, e assim sucessivamente, a ordem de produção
'deve ser informada até que seja concluída e caso exista apontamento de quantidade produzida e/ou quantidade consumida
'de insumo (K235).
'A quantidade de produção acabada deve ser expressa, obrigatoriamente, na unidade de medida de controle de
'estoque constante no campo 06 do registro 0200: UNID_INV.
'Quando houver identificação da ordem de produção, a chave deste registro são os campos: COD_DOC_OP e
'COD_ITEM.
'Nos casos em que a ordem de produção não for identificada, o campo chave passa a ser COD_ITEM.

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim objSpedFiscal As ClassSpedFiscal
Dim sSQL As String, sCodigoOP As String, sProduto As String, sSiglaUMEstoque As String
Dim lNumIntDocItemOP As Long, dtDataProd As Date, iClasseUM As Integer
Dim dtDataInicioProd As Date, dtDataFimProd As Date, dtDataEmissaoOP As Date, dFator As Double
Dim iSituacao As Integer, sSiglaUM As String, dQtdeProd As Double, dQtdeProdTotal As Double
Dim lSeq0200 As Long

On Error GoTo Erro_SpedFiscal_K230
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216276
    Next
    
    sCodigoOP = String(STRING_OPCODIGO, 0)
    sProduto = String(STRING_PRODUTO, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    
    'Obtém todas as OPs que tiverem ou produção ou requisição no período
    sSQL = "SELECT I.NumIntDoc, I.Codigo, I.Produto, I.DataInicioProd, O.DataEmissao, I.Situacao, P.ClasseUM, P.SiglaUMEstoque FROM OrdensProducaoTodas AS O, ItensOPTodos As I, Produtos AS P "
    sSQL = sSQL & "WHERE P.Codigo = I.Produto AND O.Codigo = I.Codigo AND O.FilialEmpresa = I.FilialEmpresa AND O.FilialEmpresa = ? AND "
    sSQL = sSQL & "EXISTS (SELECT * FROM MovimentoEstoque As M WHERE M.Data BETWEEN ? AND ? AND M.TipoMov IN (?,?,?,?,?,?) AND M.NumIntDocOrigem = I.NumIntDoc AND M.NumIntDocEst = 0) AND P.Natureza IN (?,?) "

    lErro = Comando_Executar(alComando(0), sSQL, lNumIntDocItemOP, sCodigoOP, sProduto, dtDataInicioProd, dtDataEmissaoOP, iSituacao, iClasseUM, sSiglaUMEstoque, objFisSpedAux.iFilialEmpresa, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, _
    MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3, MOV_EST_PRODUCAO_OUTROS, MOV_EST_REQ_PRODUCAO, MOV_EST_REQ_PRODUCAO_BENEF3, MOV_EST_REQ_PRODUCAO_OUTROS, NATUREZA_PROD_PRODUTO_ACABADO, NATUREZA_PROD_PRODUTO_EM_PROCESSO)
    If lErro <> AD_SQL_SUCESSO Then gError 216277
   
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216278
    
    Do While lErro = AD_SQL_SUCESSO
  
        lErro = SpedFiscal_0200(sProduto, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, lSeq0200, objFisSpedAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
  
        Set objSpedFiscal = objFisSpedAux.objSpedFiscal
        
        objSpedFiscal.sTipo = "K230"
        objFisSpedAux.objSpedRegistroK230.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objFisSpedAux.objSpedRegistroK230.iCount = objFisSpedAux.objSpedRegistroK230.iCount + 1
        
        dQtdeProdTotal = 0
        dtDataFimProd = DATA_NULA
        
        sSiglaUM = String(STRING_UM_SIGLA, 0)
        
        'Obtém o total produzido e data de finalização da produção (última entrada se estiver baixada)
        lErro = Comando_Executar(alComando(1), "SELECT Quantidade, SiglaUM, Data FROM MovimentoEstoque WHERE NumIntDocOrigem = ? AND FilialEmpresa = ? AND Data BETWEEN ? AND ? AND TipoMov IN (?,?,?) AND NumIntDocEst = 0", _
                                                dQtdeProd, sSiglaUM, dtDataProd, lNumIntDocItemOP, objFisSpedAux.iFilialEmpresa, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3, MOV_EST_PRODUCAO_OUTROS)
        If lErro <> AD_SQL_SUCESSO Then gError 216279
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216280
        
        Do While lErro = AD_SQL_SUCESSO
        
            If dtDataFimProd < dtDataProd Then dtDataFimProd = dtDataProd
        
            'Descobre o fator de conversao da UM de entrada p/UM de estoque do produto
            lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUM, sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            dQtdeProdTotal = dQtdeProdTotal + dQtdeProd * dFator
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216281
        
        Loop
        
        If iSituacao <> ITEMOP_SITUACAO_BAIXADA Then
            dtDataFimProd = DATA_NULA 'Se não baixou ainda não finalizou a produção
        End If
    
        '01 REG Texto fixo contendo "K230"
        '02 DT_INI_OP Data de início da ordem de produção
        '03 DT_FIN_OP Data de conclusão da ordem de produção
        '04 COD_DOC_OP Código de identificação da ordem de produção
        '05 COD_ITEM Código do item produzido (campo 02 do Registro 0200)
        '06 QTD_ENC Quantidade de produção acabada
        objSpedFiscal.sRegistro = "|K230|"
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(IIf(dtDataInicioProd = DATA_NULA, dtDataEmissaoOP, dtDataInicioProd), "ddmmyyyy") & "|" '02
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(dtDataFimProd = DATA_NULA, "", Format(dtDataFimProd, "ddmmyyyy")) & "|" '03
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(sCodigoOP) & "|" '04
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(sProduto) & "|" '05
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dQtdeProdTotal, "##0.000") & "|" '06
                
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 216282
        
        lErro = Comando_Executar(alComando(2), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 216283
               
        'trata o REGISTRO K235: INSUMOS CONSUMIDOS
        lErro = SpedFiscal_K235(objFisSpedAux, lNumIntDocItemOP, sProduto, objSpedFiscal.lSeq)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216284

    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_K230 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_K230:

   SpedFiscal_K230 = gErr

   Select Case gErr
    
        Case 216276
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 216277, 216278, 216284
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", gErr)
    
        Case 216279 To 216281
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
    
        Case 216282
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 216283
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216285)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_K235(ByVal objFisSpedAux As ClassFISSpedAux, ByVal lNumIntDocItemOP As Long, ByVal sProdPai As String, ByVal lSeqPai As Long) As Long
'REGISTRO K235: INSUMOS CONSUMIDOS
'Este registro tem o objetivo de informar o consumo de mercadoria no processo produtivo, vinculado ao produto
'resultante informado no campo COD_ITEM do Registro K230 – Itens Produzidos.
'Na industrialização efetuada para terceiro por encomenda devem ser considerados os insumos recebidos do
'encomendante e os insumos próprios do industrializador.
'Este registro é obrigatório quando existir o registro pai K230 e:
'a)a informação da quantidade produzida (K230) for por período de apuração(K100); ou
'b) a ordem de produção (K230) se iniciar e concluir no período de apuração (K100); ou
'c) a ordem de produção (K230) se iniciar no período de apuração (K100) e não for concluída no mesmo período.
'A quantidade consumida deve ser expressa, obrigatoriamente, na unidade de medida de controle de estoque
'constante no campo 06 do registro 0200: UNID_INV.
'A chave deste registro são os campos DT_SAÍDA e COD_ITEM.
'A partir de 01/01/2017, a chave deste registro são os campos DT_SAÍDA, COD_ITEM e, quando existir, o
'COD_INS_SUBST.

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(0 To 3) As Long, lSeq0200 As Long
Dim iIndice As Integer, sProdAux As String, iAux As Integer
Dim objSpedFiscal As ClassSpedFiscal, bAchou As Boolean
Dim iClasseUM As Integer, sSiglaUM As String, sSiglaUMEstoque As String, dFator As Double
Dim dQtdeReq As Double, dtData As Date, sProd1 As String, sProd2 As String
Dim objME As ClassItemMovEstoque, colME As New Collection

On Error GoTo Erro_SpedFiscal_K235
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216286
    Next
   
    sSiglaUM = String(STRING_UM_SIGLA, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    sProd1 = String(STRING_PRODUTO, 0)
    sProd2 = String(STRING_PRODUTO, 0)
    
    'Obtém o total requisitado para produção
    lErro = Comando_Executar(alComando(0), "SELECT M.Quantidade, M.SiglaUM, M.Data, P.Codigo, P.ClasseUM, P.SiglaUMEstoque FROM MovimentoEstoque AS M, Produtos As P WHERE P.Codigo = M.Produto AND M.NumIntDocOrigem = ? AND M.FilialEmpresa = ? AND M.Data BETWEEN ? AND ? AND M.TipoMov IN (?,?,?) AND NumIntDocEst = 0 AND P.Natureza IN (?,?,?,?,?,?,?,?)", _
                                            dQtdeReq, sSiglaUM, dtData, sProd1, iClasseUM, sSiglaUMEstoque, lNumIntDocItemOP, objFisSpedAux.iFilialEmpresa, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, MOV_EST_REQ_PRODUCAO, MOV_EST_REQ_PRODUCAO_BENEF3, MOV_EST_REQ_PRODUCAO_OUTROS, NATUREZA_PROD_PRODUTO_REVENDA, NATUREZA_PROD_MATERIA_PRIMA, NATUREZA_PROD_EMBALAGENS, NATUREZA_PROD_PRODUTO_EM_PROCESSO, NATUREZA_PROD_PRODUTO_INTERMEDIARIO, NATUREZA_PROD_OUTROS_INSUMOS, NATUREZA_PROD_PRODUTO_ACABADO, NATUREZA_PROD_SUBPRODUTO)
    If lErro <> AD_SQL_SUCESSO Then gError 216287
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216288
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Descobre o fator de conversao da UM de entrada p/UM de estoque do produto
        lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUM, sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        bAchou = False
        For Each objME In colME
            If objME.dtData = dtData And objME.sProduto = sProd1 Then
                bAchou = True
                objME.dQuantidade = objME.dQuantidade + dQtdeReq * dFator
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objME = New ClassItemMovEstoque
            objME.sProduto = sProd1
            objME.dtData = dtData
            objME.dQuantidade = dQtdeReq * dFator
            colME.Add objME
        End If
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216295
        
    Loop
    
    For Each objME In colME
    
        Set objSpedFiscal = objFisSpedAux.objSpedFiscal
        
        objSpedFiscal.sTipo = "K235"
        objFisSpedAux.objSpedRegistroK235.sTipo = objSpedFiscal.sTipo
       
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objFisSpedAux.objSpedRegistroK235.iCount = objFisSpedAux.objSpedRegistroK235.iCount + 1
        
        'Identificar produtos que não estão no kit padrão e ver quais produtos estão e tem o sProd1 como substituto no
        'cadastro de produtos
        lErro = Comando_Executar(alComando(2), "SELECT Nivel FROM Kit As K, ProdutoKit AS P WHERE K.ProdutoRaiz = P.ProdutoRaiz AND K.versao = P.Versao AND K.Situacao = 1 AND K.ProdutoRaiz = ? AND P.Produto = ? ", iAux, sProdPai, objME.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 216289
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216290
        
        If lErro = AD_SQL_SUCESSO Then
            sProd2 = "" 'o produto faz parte do kit logo não é substituto
        Else
        
            sProdAux = String(STRING_PRODUTO, 0)
        
            'Lê os produtos que esse pode susbtituir e identifica aquele que faz parte da composição do kit
            lErro = Comando_Executar(alComando(3), "SELECT Codigo FROM Produtos WHERE (Substituto1 = ? OR Substituto2 = ?) AND Codigo IN (SELECT P.Produto FROM Kit As K, ProdutoKit AS P WHERE K.ProdutoRaiz = P.ProdutoRaiz AND K.versao = P.Versao AND K.Situacao = 1 AND K.ProdutoRaiz = ?) ", sProdAux, objME.sProduto, objME.sProduto, sProdPai)
            If lErro <> AD_SQL_SUCESSO Then gError 216291
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216292
            
            If lErro = AD_SQL_SUCESSO Then
                sProd2 = sProdAux
            Else
                sProd2 = "" 'dar erro ou aviso ou deixar acertar no PVA ?
            End If
        
        End If
    
        '01 REG Texto fixo contendo "K235" C 4 - O
        '02 DT_SAÍDA Data de saída do estoque para alocação ao produto
        '03 COD_ITEM Código do item componente/insumo (campo 02 do Registro 0200)
        '04 QTD Quantidade consumida do item N - 3 O
        '05 COD_INS_SUBST Código do insumo que foi substituído, caso ocorra a substituição (campo 02 do Registro 0210)
        objSpedFiscal.sRegistro = "|K235|" '01
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(objME.dtData, "ddmmyyyy") & "|" '02
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(objME.sProduto) & "|" '03
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(objME.dQuantidade, "##0.000") & "|" '04
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(sProd2) & "|" '05
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 216293
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
        If lErro <> AD_SQL_SUCESSO Then gError 216294
    
        lErro = SpedFiscal_0200(objME.sProduto, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, lSeq0200, objFisSpedAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        If Len(Trim(sProd2)) > 0 Then
            lErro = SpedFiscal_0200(sProd2, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, lSeq0200, objFisSpedAux)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
    
    Next

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_K235 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_K235:

   SpedFiscal_K235 = gErr

   Select Case gErr
    
        Case 216286
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 216287, 216288, 216295
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
            
        Case 216289 To 216292
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)
    
        Case 216293
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 216294
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216296)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_K250(ByVal objFisSpedAux As ClassFISSpedAux) As Long
'REGISTRO K250: INDUSTRIALIZAÇÃO EFETUADA POR TERCEIROS – ITENS PRODUZIDOS
'Este registro tem o objetivo de informar os produtos que foram industrializados por terceiros por encomenda e sua
'quantidade.
'A quantidade produzida deve ser expressa, obrigatoriamente, na unidade de medida de controle de estoque
'constante no campo 06 do registro 0200: UNID_INV.
'A chave deste registro são os campos DT_PROD e COD_ITEM

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(0 To 1) As Long
Dim iIndice As Integer, dFator As Double, lCodMovEst As Long
Dim objSpedFiscal As ClassSpedFiscal, sSQL As String
Dim lNumIntDocOrigem As Long, iTipoNumIntDocOrigem As Integer, dtData As Date, sProduto As String, sSiglaUM As String, dQtde As Double, iClasseUM As Integer, sSiglaUMEstoque As String
Dim lSeq0200 As Long

On Error GoTo Erro_SpedFiscal_K250
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216297
    Next

    sProduto = String(STRING_PRODUTO, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    sSiglaUM = String(STRING_UM_SIGLA, 0)
    
    'Lê todos os movimentos de entrada de material beneficiado em 3os
    sSQL = "SELECT M.NumIntDocOrigem, M.TipoNumIntDocOrigem, M.Codigo, M.Data, M.Produto, M.SiglaUM, M.Quantidade, P.ClasseUM, P.SiglaUMEstoque "
    sSQL = sSQL & "FROM MovimentoEstoque As M, Produtos AS P "
    sSQL = sSQL & "WHERE M.Produto = P.Codigo AND M.FilialEmpresa = ? AND M.TipoMov = 139 AND M.Data BETWEEN ? AND ? AND M.NumIntDocEst = 0 AND P.Natureza IN (?,?)"

    lErro = Comando_Executar(alComando(0), sSQL, lNumIntDocOrigem, iTipoNumIntDocOrigem, lCodMovEst, dtData, sProduto, sSiglaUM, dQtde, iClasseUM, sSiglaUMEstoque, objFisSpedAux.iFilialEmpresa, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, NATUREZA_PROD_PRODUTO_ACABADO, NATUREZA_PROD_PRODUTO_EM_PROCESSO)
    If lErro <> AD_SQL_SUCESSO Then gError 216298
   
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216299
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = SpedFiscal_0200(sProduto, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, lSeq0200, objFisSpedAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        Set objSpedFiscal = objFisSpedAux.objSpedFiscal
        
        objSpedFiscal.sTipo = "K250"
        objFisSpedAux.objSpedRegistroK250.sTipo = objSpedFiscal.sTipo
        
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objFisSpedAux.objSpedRegistroK250.iCount = objFisSpedAux.objSpedRegistroK250.iCount + 1
    
        'Descobre o fator de conversao da UM de entrada p/UM de estoque do produto
        lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUM, sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
     
        '01 REG Texto fixo contendo "K250" C 4 - O
        '02 DT_PROD Data do reconhecimento da produção ocorrida no terceiro
        '03 COD_ITEM Código do item produzido (campo 02 do Registro 0200)
        '04 QTD Quantidade produzida
        objSpedFiscal.sRegistro = "|K250|" '01
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(dtData = DATA_NULA, "", Format(dtData, "ddmmyyyy")) & "|" '02
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(sProduto) & "|" '03
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(dQtde * dFator, "##0.000") & "|" '04
                
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 216300
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
        If lErro <> AD_SQL_SUCESSO Then gError 216301

        'trata o REGISTRO K255: INDUSTRIALIZAÇÃO EM TERCEIROS – INSUMOS CONSUMIDOS
        lErro = SpedFiscal_K255(objFisSpedAux, iTipoNumIntDocOrigem, lNumIntDocOrigem, lCodMovEst, sProduto, objSpedFiscal.lSeq)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216302

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_K250 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_K250:

   SpedFiscal_K250 = gErr

   Select Case gErr
    
        Case 216297
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 216298, 216299, 216302
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
    
        Case 216300
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 216301
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216303)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_K255(ByVal objFisSpedAux As ClassFISSpedAux, ByVal iTipoNumIntDocOrigem As Integer, ByVal lNumIntDocOrigem As Long, ByVal lCodMovEst As Long, ByVal sProdPai As String, ByVal lSeqPai As Long) As Long
'REGISTRO K255: INDUSTRIALIZAÇÃO EM TERCEIROS – INSUMOS CONSUMIDOS
'Este registro tem o objetivo de informar a quantidade de consumo do insumo que foi remetido para ser
'industrializado em terceiro, vinculado ao produto resultante informado no campo COD_ITEM do Registro K250. É
'obrigatório caso exista o registro pai K250.
'A quantidade consumida deve ser expressa, obrigatoriamente, na unidade de medida de controle de estoque
'constante no campo 06 do registro 0200: UNID_INV.
'A chave deste registro são os campos DT_CONS e COD_ITEM deste Registro.
'A partir de 01/01/2017, a chave deste registro são os campos DT_SAÍDA, COD_ITEM e, quando existir, o
'COD_INS_SUBST.

'????????????
'O manual diz que a quantidade desse registro é referente a produção da quantidade do registro k250 e que a data tem
'que compreender o período do k100, dito isso, se consumiu em período anterior para algo produzido nesse vai dar problema,
'mas como, normalmente, a devolução simbólica ocorre casada com a mercadoria produzida estou seguindo a risca o manual

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(0 To 4) As Long, iAux As Integer
Dim iIndice As Integer, bAchou As Boolean, lSeq0200 As Long
Dim objSpedFiscal As ClassSpedFiscal, sProdAux As String
Dim iClasseUM As Integer, sSiglaUM As String, sSiglaUMEstoque As String, dFator As Double
Dim dQtdeReq As Double, dtData As Date, sProd1 As String, sProd2 As String
Dim objME As ClassItemMovEstoque, colME As New Collection

On Error GoTo Erro_SpedFiscal_K255
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216304
    Next

   
    sSiglaUM = String(STRING_UM_SIGLA, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    sProd1 = String(STRING_PRODUTO, 0)
    sProd2 = String(STRING_PRODUTO, 0)
    
    'Obtém o material consumido na produção em terceiro nas NFs tipo "225-NF externa de entrada fat mat de 3os benef c\ rem simbólica"
    'Tipo Movto: "161 Nota Fiscal de Entrada - Dev.Simb.Beneficiamento"
    lErro = Comando_Executar(alComando(0), "SELECT M.Quantidade, M.SiglaUM, M.Data, P.Codigo, P.ClasseUM, P.SiglaUMEstoque FROM MovimentoEstoque AS M, Produtos As P WHERE P.Codigo = M.Produto AND M.TipoNumIntDocOrigem = 1 AND M.Codigo = ? AND M.FilialEmpresa = ? AND M.Data BETWEEN ? AND ? AND M.TipoMov = 161 AND NumIntDocEst = 0 AND P.Natureza IN (?,?,?,?,?,?,?,?)", _
                                            dQtdeReq, sSiglaUM, dtData, sProd1, iClasseUM, sSiglaUMEstoque, lCodMovEst, objFisSpedAux.iFilialEmpresa, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, NATUREZA_PROD_PRODUTO_REVENDA, NATUREZA_PROD_MATERIA_PRIMA, NATUREZA_PROD_EMBALAGENS, NATUREZA_PROD_PRODUTO_EM_PROCESSO, NATUREZA_PROD_PRODUTO_INTERMEDIARIO, NATUREZA_PROD_OUTROS_INSUMOS, NATUREZA_PROD_PRODUTO_ACABADO, NATUREZA_PROD_SUBPRODUTO)
    
    If lErro <> AD_SQL_SUCESSO Then gError 216305
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216306
    
    Do While lErro = AD_SQL_SUCESSO
       
        'Descobre o fator de conversao da UM de entrada p/UM de estoque do produto
        lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUM, sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
        bAchou = False
        For Each objME In colME
            If objME.dtData = dtData And objME.sProduto = sProd1 Then
                bAchou = True
                objME.dQuantidade = objME.dQuantidade + dQtdeReq * dFator
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objME = New ClassItemMovEstoque
            objME.sProduto = sProd1
            objME.dtData = dtData
            objME.dQuantidade = dQtdeReq * dFator
            colME.Add objME
        End If
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216307
    
    Loop
    
    sSiglaUM = String(STRING_UM_SIGLA, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    sProd1 = String(STRING_PRODUTO, 0)
    sProd2 = String(STRING_PRODUTO, 0)
    
    'Obtém o material consumido na produção em terceiro nas notas com aba de beneficiamento (tipos 132 e "133-NF externa de entrada dev. Simbólica de mat. Env.p/benef")
    'Tipo Movto: "91 -Industrialização em Terceiros de Mat. Prima Nossa - Saida"
    lErro = Comando_Executar(alComando(1), "SELECT M.Quantidade, M.SiglaUM, M.Data, P.Codigo, P.ClasseUM, P.SiglaUMEstoque FROM MovimentoEstoque AS M, Produtos As P, ItensNFiscal AS I, MovimentoEstoque AS X WHERE P.Codigo = M.Produto AND X.TipoNumIntDocOrigem = ? AND X.NumIntDocOrigem = ? AND I.NumInTDoc = X.NumIntDocOrigem AND M.TipoNumIntDocOrigem = 2 AND M.NumIntDocOrigem = I.NumIntNF AND M.FilialEmpresa = ? AND M.Data BETWEEN ? AND ? AND M.TipoMov = 91 AND M.NumIntDocEst = 0", _
                                            dQtdeReq, sSiglaUM, dtData, sProd1, iClasseUM, sSiglaUMEstoque, iTipoNumIntDocOrigem, lNumIntDocOrigem, objFisSpedAux.iFilialEmpresa, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 216308
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216309
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Descobre o fator de conversao da UM de entrada p/UM de estoque do produto
        lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUM, sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
        bAchou = False
        For Each objME In colME
            If objME.dtData = dtData And objME.sProduto = sProd1 Then
                bAchou = True
                objME.dQuantidade = objME.dQuantidade + dQtdeReq * dFator
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objME = New ClassItemMovEstoque
            objME.sProduto = sProd1
            objME.dtData = dtData
            objME.dQuantidade = dQtdeReq * dFator
            colME.Add objME
        End If
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216310
    
    Loop
    
    
    For Each objME In colME
        
        Set objSpedFiscal = objFisSpedAux.objSpedFiscal
        
        objSpedFiscal.sTipo = "K255"
        objFisSpedAux.objSpedRegistroK255.sTipo = objSpedFiscal.sTipo
    
        objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
        objFisSpedAux.objSpedRegistroK255.iCount = objFisSpedAux.objSpedRegistroK255.iCount + 1
    
        'Identificar produtos que não estão no kit padrão e ver quais produtos estão e tem o sProd1 como substituto no
        'cadastro de produtos
        lErro = Comando_Executar(alComando(2), "SELECT Nivel FROM Kit As K, ProdutoKit AS P WHERE K.ProdutoRaiz = P.ProdutoRaiz AND K.versao = P.Versao AND K.Situacao = 1 AND K.ProdutoRaiz = ? AND P.Produto = ? ", iAux, sProdPai, objME.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 216311
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216312
        
        If lErro = AD_SQL_SUCESSO Then
            sProd2 = "" 'o produto faz parte do kit logo não é substituto
        Else
        
            sProdAux = String(STRING_PRODUTO, 0)
        
            'Lê os produtos que esse pode susbtituir e identifica aquele que faz parte da composição do kit
            lErro = Comando_Executar(alComando(3), "SELECT Codigo FROM Produtos WHERE (Substituto1 = ? OR Substituto2 = ?) AND Codigo IN (SELECT P.Produto FROM Kit As K, ProdutoKit As P WHERE K.ProdutoRaiz = P.ProdutoRaiz AND K.versao = P.Versao AND K.Situacao = 1 AND K.ProdutoRaiz = ?) ", sProdAux, objME.sProduto, objME.sProduto, sProdPai)
            If lErro <> AD_SQL_SUCESSO Then gError 216313
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216314
            
            If lErro = AD_SQL_SUCESSO Then
                sProd2 = sProdAux
            Else
                sProd2 = "" 'dar erro ou aviso ou deixar acertar no PVA ?
            End If
        
        End If
        
        '01 REG Texto fixo contendo "K255"
        '02 DT_CONS Data do reconhecimento do consumo do insumo referente ao produto informado no campo 04 do Registro K250
        '03 COD_ITEM Código do insumo (campo 02 do Registro 0200)
        '04 QTD Quantidade de consumo do insumo.
        '05 COD_INS_SUBST Código do insumo que foi substituído, caso ocorra a substituição (campo 02 do Registro 0210)
        objSpedFiscal.sRegistro = "|K255|" '01
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(objME.dtData, "ddmmyyyy") & "|" '02
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(objME.sProduto) & "|" '03
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & Format(objME.dQuantidade, "##0.000") & "|" '04
        objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & SFF_Texto(sProd2) & "|" '05
    
        If Len(objSpedFiscal.sRegistro) > 500 Then gError 216315
        
        lErro = Comando_Executar(alComando(4), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
        objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPai)
        If lErro <> AD_SQL_SUCESSO Then gError 216316
    
        lErro = SpedFiscal_0200(objME.sProduto, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, lSeq0200, objFisSpedAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        If Len(Trim(sProd2)) > 0 Then
            lErro = SpedFiscal_0200(sProd2, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0200, objFisSpedAux.objSpedRegistro0205, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim, objFisSpedAux.objSpedRegistro0190, objFisSpedAux.objSpedRegistro0220, objFisSpedAux.colUM, lSeq0200, objFisSpedAux)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
    
    Next

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_K255 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_K255:

   SpedFiscal_K255 = gErr

   Select Case gErr
    
        Case 216304
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 216305 To 216310
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
            
        Case 216311 To 216314
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)
    
        Case 216315
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 216316
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216317)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function SpedFiscal_K990(ByVal objFisSpedAux As ClassFISSpedAux) As Long
'REGISTRO K990: ENCERRAMENTO DO BLOCO K
'Este registro destina-se a identificar o encerramento do bloco K e a informar a quantidade de linhas (registros)
'existentes no bloco.

Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lTotal As Long, objSpedFiscal As ClassSpedFiscal

On Error GoTo Erro_SpedFiscal_K990
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216318
    Next

    Set objSpedFiscal = objFisSpedAux.objSpedFiscal
    
    objSpedFiscal.sTipo = "K990"
    objFisSpedAux.objSpedRegistroK990.sTipo = objSpedFiscal.sTipo
    
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objFisSpedAux.objSpedRegistroK990.iCount = objFisSpedAux.objSpedRegistroK990.iCount + 1

    With objFisSpedAux
        lTotal = .objSpedRegistroK001.iCount + .objSpedRegistroK100.iCount + .objSpedRegistroK200.iCount + .objSpedRegistroK220.iCount + .objSpedRegistroK230.iCount + .objSpedRegistroK235.iCount + .objSpedRegistroK250.iCount + .objSpedRegistroK255.iCount + .objSpedRegistroK990.iCount
    End With

    objSpedFiscal.sRegistro = "|K990|" & lTotal & "|"

    If Len(objSpedFiscal.sRegistro) > 500 Then gError 216319
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro ) VALUES (?,?,?,?,?)", _
    objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro)
    If lErro <> AD_SQL_SUCESSO Then gError 216320


    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_K990 = SUCESSO

    Exit Function
    
Erro_SpedFiscal_K990:

   SpedFiscal_K990 = gErr

   Select Case gErr
    
        Case 216318
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 216319
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 216320
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216321)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function CTe_Obtem_CodsIBGE(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lNumIntNF As Long, sIBGEOrig As String, sIBGEDest As String) As Long

Dim lErro As Long
Dim iTipoNFiscal As Integer, lCliente As Long, iFilialCli As Integer, lFornecedor As Long, iFilialForn As Integer, lNumIntNotaOriginal As Long
Dim lClienteAux As Long, iFilialCliAux As Integer, lFornecedorAux As Long, iFilialFornAux As Integer
Dim sIBGECliForn As String

On Error GoTo Erro_CTe_Obtem_CodsIBGE
        
    lErro = Comando_Executar(lComando1, "SELECT TipoNFiscal, Cliente, FilialCli, Fornecedor, FilialForn, NumIntNotaOriginal FROM NFiscal AS N WHERE N.NumIntDoc = ? ", iTipoNFiscal, lClienteAux, iFilialCliAux, lFornecedorAux, iFilialFornAux, lNumIntNotaOriginal, lNumIntNF)
    If lErro <> AD_SQL_SUCESSO Then gError 204129

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204130
    
    If lNumIntNotaOriginal = 0 Then
        'Usa o cliente\fornecedor do próprio cte
        lCliente = lClienteAux
        iFilialCli = iFilialCliAux
        lFornecedor = lFornecedorAux
        iFilialForn = iFilialFornAux
    Else
        'Pega o cliente\fornecedor da nota original
        lErro = Comando_Executar(lComando2, "SELECT Cliente, FilialCli, Fornecedor, FilialForn FROM NFiscal AS N WHERE N.NumIntDoc = ? ", lCliente, iFilialCli, lFornecedor, iFilialForn, lNumIntNotaOriginal)
        If lErro <> AD_SQL_SUCESSO Then gError 204129
    
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204130
    End If
    
    sIBGECliForn = String(STRING_MAXIMO, 0)
    If lCliente <> 0 Then
        lErro = Comando_Executar(lComando3, "SELECT C.CodIBGE FROM FiliaisClientes AS F, Enderecos AS E, Cidades As C WHERE F.Endereco = E.Codigo AND E.Cidade = C.Descricao AND F.CodCliente = ? AND F.CodFilial = ? ", sIBGECliForn, lCliente, iFilialCli)
        If lErro <> AD_SQL_SUCESSO Then gError 204129
    
        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204130
    Else
        lErro = Comando_Executar(lComando4, "SELECT C.CodIBGE FROM FiliaisFornecedores AS F, Enderecos AS E, Cidades As C WHERE F.Endereco = E.Codigo AND E.Cidade = C.Descricao AND F.CodFornecedor = ? AND F.CodFilial = ? ", sIBGECliForn, lFornecedor, iFilialForn)
        If lErro <> AD_SQL_SUCESSO Then gError 204129
    
        lErro = Comando_BuscarPrimeiro(lComando4)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204130
    End If
    
    If iTipoNFiscal = 104 Then '104 Conhecimento de Frete - Compra ou Transferencia
        sIBGEOrig = sIBGECliForn
        sIBGEDest = gsCodIBGE
    Else '105 Conhecimento de Frete - Venda ou Transferencia
        sIBGEOrig = gsCodIBGE
        sIBGEDest = sIBGECliForn
    End If

    CTe_Obtem_CodsIBGE = SUCESSO

    Exit Function
    
Erro_CTe_Obtem_CodsIBGE:

   CTe_Obtem_CodsIBGE = gErr

   Select Case gErr
    

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216321)
     
    End Select


End Function

Private Function Processa_Bloco_B(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objFisSpedAux As ClassFISSpedAux) As Long

Dim lErro As Long
Dim bAbreBlocoB As Boolean
Dim colValores As New Collection

On Error GoTo Erro_Processa_Bloco_B

    '========================================================================================
    '=============================== INÍCIO DO BLOCO B ======================================
    '===================== BLOCO B: ESCRITURAÇÃO E APURAÇÃO DO ISS ==========================
    '========================================================================================
    
    'REGISTRO B001 N1: ABERTURA DO BLOCO B
    lErro = Processa_Registro_B001(objEFD, bAbreBlocoB)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If bAbreBlocoB Then
    
        'REGISTRO B020 N2: NOTA FISCAL (CÓDIGO 01), NOTA FISCAL DE SERVIÇOS (CÓDIGO 03), NOTA FISCAL DE SERVIÇOS AVULSA (CÓDIGO 3B), NOTA FISCAL DE PRODUTOR (CÓDIGO 04), CONHECIMENTO DE TRANSPORTE RODOVIÁRIO DE CARGAS(CÓDIGO 08), NF-e (CÓDIGO 55) e NFC-e (CÓDIGO 65).
        'REGISTRO B025 N3: DETALHAMENTO POR COMBINAÇÃO DE ALÍQUOTA E ITEM DA LISTA DE SERVIÇOS DA LC 116/2003)
        'REGISTRO B030 N2: NOTA FISCAL DE SERVIÇOS SIMPLIFICADA (CÓDIGO 3A)
        'REGISTRO B035 N3: DETALHAMENTO POR COMBINAÇÃO DE ALÍQUOTA E ITEM DA LISTA DE SERVIÇOS DA LC 116/2003)
        'REGISTRO B350 N2: SERVIÇOS PRESTADOS POR INSTITUIÇÕES FINANCEIRAS
        'REGISTRO B420 N2: TOTALIZAÇÃO DOS VALORES DE SERVIÇOS PRESTADOS POR COMBINAÇÃO DE ALÍQUOTA E ITEM DA LISTA DE SERVIÇOS DA LC 116/2003
        'REGISTRO B440 N2: TOTALIZAÇÃO DOS VALORES RETIDOS
        'REGISTRO B460 N2: DEDUÇÕES DO ISS
        'REGISTRO B470 N2: APURAÇÃO DO ISS
        'REGISTRO B500 N2: APURAÇÃO DO ISS SOCIEDADE UNIPROFISSIONAL
        'REGISTRO B510 N3: UNIPROFISSIONAL - EMPREGADOS E SÓCIOS
        
        lErro = Processa_Registro_B020(objEFD, objFisSpedAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Processa_Registro_B420(objEFD, colValores)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Processa_Registro_B440(objEFD, colValores, objFisSpedAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Processa_Registro_B470(objEFD, colValores)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    'REGISTRO B990 N1: ENCERRAMENTO DO BLOCO B
    lErro = Processa_Registro_B990(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Bloco_B = SUCESSO
    
    Exit Function
    
Erro_Processa_Bloco_B:

    Processa_Bloco_B = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209359)
     
    End Select
    
End Function

Private Function Processa_Registro_B001(ByVal objEFD As ClassEFDPisCofinsSel, bAbreBlocoB As Boolean) As Long
'REGISTRO B001: ABERTURA DO BLOCO B
Dim lErro As Long
Dim lContador As Long
Dim iTemReg As Integer

On Error GoTo Erro_Processa_Registro_B001

    '******** registro B001 *****************
           
    lErro = Comando_Executar(objEFD.ProxComando, "SELECT COUNT(*) FROM SpedDocFiscais WHERE NumIntSped = ? AND Bloco = ? AND Modelo <> '00'", lContador, objEFD.lNumIntSped, "B")
    If lErro <> AD_SQL_SUCESSO Then gError 209467

    lErro = Comando_BuscarPrimeiro(objEFD.UltimoCmdUsado)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209468
               
    objEFD.objSpedFiscal.sTipo = "B001"
    objEFD.objSpedFiscal.lSeqPai = 0

    If lContador > 0 Then
        iTemReg = 0
        bAbreBlocoB = True
    Else
        iTemReg = 1
        bAbreBlocoB = False
    End If

    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "B001")
    Call objEFD.Reg_Campo_add(2, "IND_DAD", CStr(iTemReg))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_B001 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_B001:

   Processa_Registro_B001 = gErr

   Select Case gErr
   
        Case 209467, 209468
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209469)
    
    End Select
    
    Exit Function

End Function

Private Function Processa_Registro_B020(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objFisSpedAux As ClassFISSpedAux) As Long
'REGISTRO B020: NOTA FISCAL (CÓDIGO 01), NOTA FISCAL DE SERVIÇOS (CÓDIGO
'03), NOTA FISCAL DE SERVIÇOS AVULSA (CÓDIGO 3B), NOTA FISCAL DE PRODUTOR
'(CÓDIGO 04), CONHECIMENTO DE TRANSPORTE RODOVIÁRIO DE CARGAS
'(CÓDIGO 08), NF-e (CÓDIGO 55) e NFC-e (CÓDIGO 65).
'Este registro deve ser gerado para cada documento fiscal código 01, 03, 3B, 04, 08, 55 e 65, conforme item 4.1.3
'do Ato COTEPE/ICMS nº 09, de 18 de abril de 2008, para os registros das aquisições e prestações de serviços sujeitas ao
'ISS.
'Exceção 1: Para documentos com código de situação (campo COD_SIT) cancelado (código “02”), cancelado
'extemporâneo (código “03”), Nota Fiscal Eletrônica (NF-e) denegada (código “04”), preencher somente os campos REG,
'IND_OPER, IND_EMIT, COD_MOD, COD_SIT, SER, NUM_DOC e CHV_NFE. Para COD_SIT = 05 (numeração
'inutilizada), todos os campos referidos anteriormente devem ser preenchidos, exceto o campo CHV_NFE. Demais campos
'deverão ser apresentados com conteúdo VAZIO “||”. Não informar registros filhos.
'Validação do Registro: Não podem ser informados, em um mesmo arquivo, dois ou mais registros B020 com
'a mesma combinação de valores dos campos formadores da chave do registro. A chave do registro B020 é:
'IND_OPER, COD_MOD, SER, NUM_DOC e DT_DOC

Dim lErro As Long
Dim sCodParticipante As String
Dim lSeq As Long, sChave As String
Dim objDoc As ClassSpedDocFiscais
Dim objDocItens As ClassSpedDocFiscaisItens
Dim tDoc As typeSpedDocFiscais
Dim alComando(0 To 2) As Long, iIndice As Integer

On Error GoTo Erro_Processa_Registro_B020


    '******************* registro B020 **********************************
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next
    
    lErro = SpedFiscal_Docs_Prepara(alComando(0), objEFD.lNumIntSped, "B", tDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209470
   
    Do While lErro = AD_SQL_SUCESSO
        
        Set objDoc = New ClassSpedDocFiscais
        
        lErro = SpedFiscal_Doc_Copia(objDoc, tDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = SpedFiscal_DocItens_Le(alComando(1), objDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        '******** registro B020 *****************
        objEFD.objSpedFiscal.sTipo = "B020"
        objEFD.objSpedFiscal.lSeqPai = 0
        
        If objDoc.lCliente > 0 Then
            lErro = SpedFiscal_0150_Cliente(objDoc.lCliente, objDoc.iFilialCli, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0150, objFisSpedAux.objSpedRegistro0175, sCodParticipante, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Else
            lErro = SpedFiscal_0150_Fornecedor(objDoc.lFornecedor, objDoc.iFilialForn, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0150, objFisSpedAux.objSpedRegistro0175, sCodParticipante, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
        
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "B020")
        Call objEFD.Reg_Campo_add(2, "IND_OPER", objDoc.sOperacao)
        Call objEFD.Reg_Campo_add(3, "IND_EMIT", objDoc.sEmitente)
        If objDoc.iSituacao <> 2 Then 'Se não é cancelado
            Call objEFD.Reg_Campo_add(4, "COD_PART", sCodParticipante)
        Else
            Call objEFD.Reg_Campo_add(4, "COD_PART", "")
        End If
        Call objEFD.Reg_Campo_add(5, "COD_MOD", Format(objDoc.sModelo))
        Call objEFD.Reg_Campo_add(6, "COD_SIT", Format(objDoc.iSituacao, "00"))
        Call objEFD.Reg_Campo_add(7, "SER", objDoc.sSerie)
        Call objEFD.Reg_Campo_add(8, "NUM_DOC", CStr(objDoc.lNumDocumento))
        Call objEFD.Reg_Campo_add(9, "CHV_NFE", objDoc.sChaveEletronica)
        
        If objDoc.iSituacao <> 2 Then 'Se não é cancelado
            Call objEFD.Reg_Campo_add(10, "DT_DOC", SFF_Data(objDoc.dtDataEmissao))
            Call objEFD.Reg_Campo_add(11, "COD_MUN_SERV", objDoc.sCodMunicServ) 'Tem que criar essa propriedade e carregá-la na tabela
            Call objEFD.Reg_Campo_add(12, "VL_CONT", SFF_Num(objDoc.dValorTotal))
            Call objEFD.Reg_Campo_add(13, "VL_MAT_TERC", "")
            Call objEFD.Reg_Campo_add(14, "VL_SUB", "")
            Call objEFD.Reg_Campo_add(15, "VL_ISNT_ISS", SFF_Num(objDoc.dValorNaoTributado))
            Call objEFD.Reg_Campo_add(16, "VL_DED_BC", "")
            Call objEFD.Reg_Campo_add(17, "VL_BC_ISS", SFF_Num(objDoc.dValorBaseISS)) 'Tem que criar essa propriedade e carregá-la na tabela
            Call objEFD.Reg_Campo_add(18, "VL_BC_ISS_RT", SFF_Num(objDoc.dValorBaseISSRet)) 'Tem que criar essa propriedade e carregá-la na tabela
            Call objEFD.Reg_Campo_add(19, "VL_ISS_RT", SFF_Num(objDoc.dValorISSRet)) 'Tem que criar essa propriedade e carregá-la na tabela
            Call objEFD.Reg_Campo_add(20, "VL_ISS", SFF_Num(objDoc.dValorISS))
            Call objEFD.Reg_Campo_add(21, "COD_INF_OBS", "")
        Else
            Call objEFD.Reg_Campo_add(10, "DT_DOC", "")
            Call objEFD.Reg_Campo_add(11, "COD_MUN_SERV", "")
            Call objEFD.Reg_Campo_add(12, "VL_CONT", "")
            Call objEFD.Reg_Campo_add(13, "VL_MAT_TERC", "")
            Call objEFD.Reg_Campo_add(14, "VL_SUB", "")
            Call objEFD.Reg_Campo_add(15, "VL_ISNT_ISS", "")
            Call objEFD.Reg_Campo_add(16, "VL_DED_BC", "")
            Call objEFD.Reg_Campo_add(17, "VL_BC_ISS", "")
            Call objEFD.Reg_Campo_add(18, "VL_BC_ISS_RT", "")
            Call objEFD.Reg_Campo_add(19, "VL_ISS_RT", "")
            Call objEFD.Reg_Campo_add(20, "VL_ISS", "")
            Call objEFD.Reg_Campo_add(21, "COD_INF_OBS", "")
        End If
        Call objEFD.Finaliza_Reg
        
        lSeq = objEFD.objSpedFiscal.lSeq
        
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objDoc.iSituacao <> 2 Then 'Se não é cancelado
            lErro = Processa_Registro_B025(objEFD, objDoc, lSeq)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209471
    
    Loop
        
    Processa_Registro_B020 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_B020:

   Processa_Registro_B020 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209473)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_A100", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_B025(ByVal objEFD As ClassEFDPisCofinsSel, ByVal objDoc As ClassSpedDocFiscais, ByVal lSeqPai As Long) As Long
'REGISTRO B025: DETALHAMENTO POR COMBINAÇÃO DE ALÍQUOTA E ITEM DA
'LISTA DE SERVIÇOS DA LC 116/2003)
'Este registro deve ser gerado para registrar de forma detalhada, por combinação de alíquota de incidência do ISS e
'Item da Lista de Serviços da Lei Complementar 116/2003, os valores informados no registro B020 “pai” (ou seja, o
'registro B020 que imediatamente o antecede no arquivo).
'Validação do Registro: Não podem ser informados, para um mesmo B020 “pai”, dois ou mais registros B025
'com a mesma combinação de valores dos campos: ALIQ_ISS e COD_SERV.
Dim lErro As Long
Dim objItem As ClassSpedDocFiscaisItens
Dim sChave As String

On Error GoTo Erro_Processa_Registro_B025

    '******** registro B025 *****************
    objEFD.objSpedFiscal.sTipo = "B025"
    objEFD.objSpedFiscal.lSeqPai = lSeqPai
    
    For Each objItem In objDoc.colItens
        
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "B025")
        Call objEFD.Reg_Campo_add(2, "VL_CONT_P", SFF_Num(objItem.dPrecoTotal))
        Call objEFD.Reg_Campo_add(3, "VL_BC_ISS_P", SFF_Num(objItem.dValorBaseISS))
        Call objEFD.Reg_Campo_add(4, "ALIQ_ISS", SFF_Perc(objItem.dAliquotaISS))
        Call objEFD.Reg_Campo_add(5, "VL_ISS_P", SFF_Num(objItem.dValorISS))
        Call objEFD.Reg_Campo_add(6, "VL_ISNT_ISS_P", SFF_Num(objItem.dValorNaoTributado))
        Call objEFD.Reg_Campo_add(7, "COD_SERV", SFF_Texto(objItem.sCodISSServ))
        Call objEFD.Finaliza_Reg
           
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next
    
    Processa_Registro_B025 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_B025:

   Processa_Registro_B025 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209477)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_A170", sChave)
    
    Exit Function
    
End Function

Private Function Processa_Registro_B420(ByVal objEFD As ClassEFDPisCofinsSel, ByVal colValores As Collection) As Long
'REGISTRO B420: TOTALIZAÇÃO DOS VALORES DE SERVIÇOS PRESTADOS POR
'COMBINAÇÃO DE ALÍQUOTA E ITEM DA LISTA DE SERVIÇOS DA LC 116/2003
'Este registro deve ser gerado para registrar de forma detalhada, por combinação de alíquota de incidência do ISS e
'Item da Lista de Serviços da Lei Complementar 116/2003, os valores totais das prestações de serviços realizadas no
'período.
'Validação do Registro: Não podem ser informados, em um mesmo arquivo, dois ou mais registros B420 com a
'mesma combinação de valores dos campos: ALIQ_ISS e COD_SERV.
Dim lErro As Long, iIndice As Integer
Dim tItem As typeSpedDocFiscaisItens
Dim alComando(0 To 0) As Long
Dim sChave As String
Dim sSQL As String
Dim objVL_CONT As New AdmCampoValor
Dim objVL_ISNT As New AdmCampoValor
Dim objVL_BC_ISS As New AdmCampoValor
Dim objVL_ISS As New AdmCampoValor

On Error GoTo Erro_Processa_Registro_B420

    '******************* registro B420 **********************************
    objEFD.objSpedFiscal.sTipo = "B420"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    objVL_CONT.sCampo = "VL_CONT"
    objVL_CONT.vValor = CDbl(0)
    objVL_ISNT.sCampo = "VL_ISNT"
    objVL_ISNT.vValor = CDbl(0)
    objVL_BC_ISS.sCampo = "VL_BC_ISS"
    objVL_BC_ISS.vValor = CDbl(0)
    objVL_ISS.sCampo = "VL_ISS"
    objVL_ISS.vValor = CDbl(0)

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next

    tItem.sCodISSServ = String(STRING_MAXIMO, 0)
    
    sSQL = "SELECT I.CodISSServ, I.AliquotaISS, "
    sSQL = sSQL + " " + "SUM (I.PrecoTotal), SUM(I.ValorBaseISS), SUM(I.ValorNaoTributado), SUM(I.ValorISS) "
    sSQL = sSQL + " " + "FROM SpedDocFiscais AS D, SpedDocFiscaisItens AS I"
    sSQL = sSQL + " " + "WHERE D.NumIntSped = I.NumIntSped AND D.TipoDoc = I.TipoDoc AND D.NumIntDoc = I.NumIntDoc AND D.Bloco = 'B' AND D.Registro = 'B020' AND D.Situacao<> 2 AND I.NumIntSped = ? "
    sSQL = sSQL + " " + "GROUP BY I.CodISSServ, I.AliquotaISS"
    sSQL = sSQL + " " + "ORDER BY I.CodISSServ, I.AliquotaISS"

    lErro = Comando_Executar(alComando(0), sSQL, tItem.sCodISSServ, tItem.dAliquotaISS, tItem.dPrecoTotal, tItem.dValorBaseISS, tItem.dValorNaoTributado, tItem.dValorISS, objEFD.lNumIntSped)
    If lErro <> SUCESSO Then gError 211975

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211976

    Do While lErro = AD_SQL_SUCESSO

        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "B420")
        Call objEFD.Reg_Campo_add(2, "VL_CONT", SFF_Num(tItem.dPrecoTotal))
        Call objEFD.Reg_Campo_add(3, "VL_BC_ISS", SFF_Num(tItem.dValorBaseISS))
        Call objEFD.Reg_Campo_add(4, "ALIQ_ISS", SFF_Perc(tItem.dAliquotaISS, 4))
        Call objEFD.Reg_Campo_add(5, "VL_ISNT_ISS", SFF_Num(tItem.dValorNaoTributado))
        Call objEFD.Reg_Campo_add(6, "VL_ISS", SFF_Num(tItem.dValorISS))
        Call objEFD.Reg_Campo_add(7, "COD_SERV", tItem.sCodISSServ)

        Call objEFD.Finaliza_Reg
        
        objVL_CONT.vValor = objVL_CONT.vValor + tItem.dPrecoTotal
        objVL_BC_ISS.vValor = objVL_BC_ISS.vValor + tItem.dValorBaseISS
        objVL_ISNT.vValor = objVL_ISNT.vValor + tItem.dValorNaoTributado
        objVL_ISS.vValor = objVL_ISS.vValor + tItem.dValorISS

        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211977

    Loop
    
    colValores.Add objVL_CONT
    colValores.Add objVL_BC_ISS
    colValores.Add objVL_ISNT
    colValores.Add objVL_ISS

    Processa_Registro_B420 = SUCESSO

    Exit Function

Erro_Processa_Registro_B420:

   Processa_Registro_B420 = gErr

   Select Case gErr
   
        Case 211975 To 211977
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211978)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_B420", sChave)

    Exit Function

End Function

Private Function Processa_Registro_B440(ByVal objEFD As ClassEFDPisCofinsSel, ByVal colValores As Collection, ByVal objFisSpedAux As ClassFISSpedAux) As Long
'REGISTRO B440: TOTALIZAÇÃO DOS VALORES RETIDOS
'Este registro deve ser gerado para registrar os valores retidos tanto referentes às prestações (declarante na condição
'de prestador) quanto às aquisições (declarante na condição de tomador). Os valores são informados por tomador e/ou
'prestador conforme o tipo de operação.
'O registro deve ser informado ainda que não tenha havido retenção de ISS nas aquisições e prestações do declarante,
'neste caso, informar a base de cálculo de retenção e ISS retido zerados.
'Validação do Registro: Não podem ser informados, em um mesmo arquivo, dois ou mais registros B440 com a
'mesma combinação de valores dos campos: IND_OPER e COD_PART.
Dim lErro As Long, iIndice As Integer
Dim tItem As typeSpedDocFiscaisItens
Dim tDoc As typeSpedDocFiscais
Dim alComando(0 To 0) As Long
Dim sChave As String
Dim sSQL As String
Dim sCodParticipante As String
Dim objVL_BC_ISS_RT As New AdmCampoValor
Dim objVL_ISS_RT As New AdmCampoValor

On Error GoTo Erro_Processa_Registro_B440

    '******************* registro B440 **********************************
    objEFD.objSpedFiscal.sTipo = "B440"
    objEFD.objSpedFiscal.lSeqPai = 0

    objVL_BC_ISS_RT.sCampo = "VL_BC_ISS_RT"
    objVL_BC_ISS_RT.vValor = CDbl(0)
    objVL_ISS_RT.sCampo = "VL_ISS_RT"
    objVL_ISS_RT.vValor = CDbl(0)

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = objEFD.ProxComando
    Next

    tDoc.sOperacao = String(STRING_MAXIMO, 0)
    
    sSQL = "SELECT D.Operacao, D.Cliente, D.FilialCli, D.Fornecedor, D.FilialForn,  "
    sSQL = sSQL + " " + "SUM (D.ValorTotal), SUM(D.ValorBaseISSRet), SUM(D.ValorISSRet) "
    sSQL = sSQL + " " + "FROM SpedDocFiscais AS D "
    sSQL = sSQL + " " + "WHERE D.Bloco = 'B' AND D.Situacao<> 2 AND D.NumIntSped = ? "
    sSQL = sSQL + " " + "GROUP BY D.Operacao, D.Cliente, D.FilialCli, D.Fornecedor, D.FilialForn"
    sSQL = sSQL + " " + "ORDER BY D.Operacao, D.Cliente, D.FilialCli, D.Fornecedor, D.FilialForn"

    lErro = Comando_Executar(alComando(0), sSQL, tDoc.sOperacao, tDoc.lCliente, tDoc.iFilialCli, tDoc.lFornecedor, tDoc.iFilialForn, tDoc.dValorTotal, tDoc.dValorBaseISSRet, tDoc.dValorISSRet, objEFD.lNumIntSped)
    If lErro <> SUCESSO Then gError 211975

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211976

    Do While lErro = AD_SQL_SUCESSO

        If tDoc.lCliente > 0 Then
            lErro = SpedFiscal_0150_Cliente(tDoc.lCliente, tDoc.iFilialCli, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0150, objFisSpedAux.objSpedRegistro0175, sCodParticipante, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Else
            lErro = SpedFiscal_0150_Fornecedor(tDoc.lFornecedor, tDoc.iFilialForn, objFisSpedAux.objSpedFiscal, objFisSpedAux.objSpedRegistro0150, objFisSpedAux.objSpedRegistro0175, sCodParticipante, objFisSpedAux.dtDataIni, objFisSpedAux.dtDataFim)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
        
        Call objEFD.Inicia_Reg
        Call objEFD.Reg_Campo_add(1, "REG", "B440")
        Call objEFD.Reg_Campo_add(2, "IND_OPER", tDoc.sOperacao)
        Call objEFD.Reg_Campo_add(3, "COD_PART", sCodParticipante)
        Call objEFD.Reg_Campo_add(4, "VL_CONT_RT", SFF_Num(tDoc.dValorTotal))
        Call objEFD.Reg_Campo_add(5, "VL_BC_ISS_RT", SFF_Num(tDoc.dValorBaseISSRet))
        Call objEFD.Reg_Campo_add(6, "VL_ISS_RT", SFF_Num(tDoc.dValorISSRet))

        Call objEFD.Finaliza_Reg
        
        If tDoc.sOperacao = "1" Then
            objVL_BC_ISS_RT.vValor = objVL_BC_ISS_RT.vValor + tDoc.dValorBaseISSRet
            objVL_ISS_RT.vValor = objVL_ISS_RT.vValor + tDoc.dValorISSRet
        End If
        
        lErro = Registro_Insere_BD(objEFD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211977

    Loop

    colValores.Add objVL_BC_ISS_RT
    colValores.Add objVL_ISS_RT

    Processa_Registro_B440 = SUCESSO

    Exit Function

Erro_Processa_Registro_B440:

   Processa_Registro_B440 = gErr

   Select Case gErr
   
        Case 211975 To 211977
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDDOCFISCAIS", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211978)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_PROCESSAMENTO", gErr, "Processa_Registro_B440", sChave)

    Exit Function

End Function

Private Function Processa_Registro_B470(ByVal objEFD As ClassEFDPisCofinsSel, ByVal colValores As Collection) As Long
'REGISTRO B470: APURAÇÃO DO ISS
'Este registro deve ser gerado para registrar os totais referentes às prestações de serviço do declarante e para apurar
'os valores a recolher do ISS próprio, do ISS retido pelo declarante na condição de tomador e do ISS Uniprofissional.
Dim lErro As Long

On Error GoTo Erro_Processa_Registro_B470

    '******************* registro B470 **********************************
    objEFD.objSpedFiscal.sTipo = "B470"
    objEFD.objSpedFiscal.lSeqPai = 0
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "B470")
    Call objEFD.Reg_Campo_add(2, "VL_CONT", SFF_Num(colValores_Obtem("VL_CONT", colValores)))
    Call objEFD.Reg_Campo_add(3, "VL_MAT_TERC", SFF_Num(colValores_Obtem("VL_MAT_TERC", colValores)))
    Call objEFD.Reg_Campo_add(4, "VL_MAT_PROP", SFF_Num(colValores_Obtem("VL_MAT_PROP", colValores)))
    Call objEFD.Reg_Campo_add(5, "VL_SUB", SFF_Num(colValores_Obtem("VL_SUB", colValores)))
    Call objEFD.Reg_Campo_add(6, "VL_ISNT", SFF_Num(colValores_Obtem("VL_ISNT", colValores)))
    Call objEFD.Reg_Campo_add(7, "VL_DED_BC", SFF_Num(colValores_Obtem("VL_MAT_TERC", colValores) + colValores_Obtem("VL_MAT_PROP", colValores) + colValores_Obtem("VL_SUB", colValores) + colValores_Obtem("VL_ISNT", colValores)))
    Call objEFD.Reg_Campo_add(8, "VL_BC_ISS", SFF_Num(colValores_Obtem("VL_BC_ISS", colValores)))
    Call objEFD.Reg_Campo_add(9, "VL_BC_ISS_RT", SFF_Num(colValores_Obtem("VL_BC_ISS_RT", colValores)))
    Call objEFD.Reg_Campo_add(10, "VL_ISS", SFF_Num(colValores_Obtem("VL_ISS", colValores)))
    Call objEFD.Reg_Campo_add(11, "VL_ISS_RT", SFF_Num(colValores_Obtem("VL_ISS_RT", colValores)))
    Call objEFD.Reg_Campo_add(12, "VL_DED", SFF_Num(colValores_Obtem("VL_DED", colValores)))
    Call objEFD.Reg_Campo_add(13, "VL_ISS_REC", SFF_Num(colValores_Obtem("VL_ISS", colValores) - colValores_Obtem("VL_ISS_RT", colValores) - colValores_Obtem("VL_DED", colValores)))
    Call objEFD.Reg_Campo_add(14, "VL_ISS_ST", SFF_Num(colValores_Obtem("VL_ISS_ST", colValores)))
    Call objEFD.Reg_Campo_add(15, "VL_ISS_REC_UNI", SFF_Num(colValores_Obtem("VL_ISS_REC_UNI", colValores)))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_B470 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_B470:

   Processa_Registro_B470 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209478)
    
    End Select
    
    Exit Function
    
End Function

Private Function colValores_Obtem(ByVal sCampo As String, ByVal colValores As Collection) As Double
Dim objVL As New AdmCampoValor
Dim dValor As Double
    For Each objVL In colValores
        If UCase(sCampo) = UCase(objVL.sCampo) Then
            dValor = objVL.vValor
            Exit For
        End If
    Next
    colValores_Obtem = dValor
End Function

Private Function Processa_Registro_B990(ByVal objEFD As ClassEFDPisCofinsSel) As Long
'REGISTRO B990: ENCERRAMENTO DO BLOCO B
'Este registro destina-se a identificar o encerramento do bloco B e informar a quantidade de linhas (registros)
'existentes no bloco.
Dim lErro As Long
Dim lContador As Long
Dim objSpedRegistro As ClassSpedRegistro

On Error GoTo Erro_Processa_Registro_B990

    '******************* registro B990 **********************************
    objEFD.objSpedFiscal.sTipo = "B990"
    objEFD.objSpedFiscal.lSeqPai = 0

    lContador = 1 ' o próprio registro conta por isso inicia em 1
        
    For Each objSpedRegistro In objEFD.colSpedRegistro
        If left(objSpedRegistro.sTipo, 1) = "B" Then
            lContador = lContador + objSpedRegistro.lCount
        End If
    Next
    
    Call objEFD.Inicia_Reg
    Call objEFD.Reg_Campo_add(1, "REG", "B990")
    Call objEFD.Reg_Campo_add(2, "QTD_LIN_C", CStr(lContador))
    Call objEFD.Finaliza_Reg
    
    lErro = Registro_Insere_BD(objEFD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Processa_Registro_B990 = SUCESSO

    Exit Function
    
Erro_Processa_Registro_B990:

   Processa_Registro_B990 = gErr

   Select Case gErr
   
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209478)
    
    End Select
    
    Exit Function
    
End Function

Private Function SpedFiscal_C190_NFCe(tTributacaoNF As typeTributacaoNF, objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC190 As ClassSpedRegistro, ByVal lSeqPaiNF As Long, ByRef tNFiscal As typeNFiscal, ByRef tTipoDocInfo As typeTipoDocInfo, ByVal colSpedC190 As Collection, ByVal iNFe As Integer, ByVal dTotalRateioNF As Double, ByVal objFisSpedAux As ClassFISSpedAux) As Long


Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objSpedC190 As ClassSpedC190

On Error GoTo Erro_SpedFiscal_C190_NFCe
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203990
    Next

    'se for nfce
    If tTipoDocInfo.iModeloArqICMS = 65 Then 'Or iNFe = 1 Then
    
        'Está errado porque falta o IPI e ICMS ST (Wagner)
'        If tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto <> dTotalRateioNF And colSpedC190.Count > 0 Then
'
'            Set objSpedC190 = colSpedC190.Item(1)
'
'            objSpedC190.dValorOper = objSpedC190.dValorOper + Round(tNFiscal.dValorFrete + tNFiscal.dValorOutrasDespesas + tNFiscal.dValorSeguro - tNFiscal.dValorDesconto - dTotalRateioNF, 2)
'
'        End If
        
        For Each objSpedC190 In colSpedC190
        
            If objSpedC190.dValorOper > 0 Or objSpedC190.dValorBCICMS > 0 Or objSpedC190.dValorICMS > 0 Or _
               objSpedC190.dValorBCICMSST > 0 Or objSpedC190.dValorICMSST > 0 Or objSpedC190.dValorRedBC > 0 Or _
               objSpedC190.dValorIPI > 0 Then
        
                objSpedFiscal.sTipo = "C190"
                objSpedRegistroC190.sTipo = objSpedFiscal.sTipo
                
                objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
                objSpedRegistroC190.iCount = objSpedRegistroC190.iCount + 1
    
                objSpedFiscal.sRegistro = "|C190|" & Format(objSpedC190.iCSTICMS, "000") & "|" & Format(objSpedC190.sCFOP, "0000") & "|" & Format(objSpedC190.dAliquotaICMS * 100, "fixed") & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorOper <= 0, 0, Format(objSpedC190.dValorOper, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorBCICMS <= 0, 0, Format(objSpedC190.dValorBCICMS, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorICMS <= 0, 0, Format(objSpedC190.dValorICMS, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorBCICMSST <= 0, 0, Format(objSpedC190.dValorBCICMSST, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorICMSST <= 0, 0, Format(objSpedC190.dValorICMSST, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorRedBC <= 0, 0, Format(objSpedC190.dValorRedBC, "fixed")) & "|"
                objSpedFiscal.sRegistro = objSpedFiscal.sRegistro & IIf(objSpedC190.dValorIPI <= 0, 0, Format(objSpedC190.dValorIPI, "fixed")) & "||"
               
                If Len(objSpedFiscal.sRegistro) > 500 Then gError 203991
                
                lErro = Comando_Executar(alComando(1), "INSERT INTO SpedFiscal( NumIntSped, Seq, FilialEmpresa, Tipo, Registro, SeqPai ) VALUES (?,?,?,?,?,?)", _
                objSpedFiscal.lNumIntSped, objSpedFiscal.lSeq, objSpedFiscal.iFilialEmpresa, objSpedFiscal.sTipo, objSpedFiscal.sRegistro, lSeqPaiNF)
                If lErro <> AD_SQL_SUCESSO Then gError 203992
                
            End If
            
        Next
    
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C190_NFCe = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C190_NFCe:

   SpedFiscal_C190_NFCe = gErr

   Select Case gErr
    
        Case 203990
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203991
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203992
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203993)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Private Function SpedFiscal_C170_NFCe(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC170 As ClassSpedRegistro, ByVal objSpedRegistroC172 As ClassSpedRegistro, ByVal objSpedRegistro0200 As ClassSpedRegistro, ByVal objSpedRegistro0205 As ClassSpedRegistro, ByVal objSpedRegistro0190 As ClassSpedRegistro, ByVal objSpedRegistro0220 As ClassSpedRegistro, ByVal lSeqPaiNF As Long, ByRef tNFiscal As typeNFiscal, ByRef tTipoDocInfo As typeTipoDocInfo, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal iNFe As Integer, dTotalRateioNF As Double, ByVal colUM As Collection, ByVal colSpedE510 As Collection, dValDebICMS_objSpedC190 As Double, dValCredICMS_objSpedC190 As Double, colSpedC190 As Collection, ByVal objFilialEmpresa As AdmFiliais, ByRef tTributacaoNF As typeTributacaoNF, ByVal sData As String, colSpedC190UF As Collection, ByVal objFisSpedAux As ClassFISSpedAux) As Long

'Private Function SpedFiscal_C170(objSpedFiscal As ClassSpedFiscal, ByVal objSpedRegistroC170 As ClassSpedRegistro, ByVal objSpedRegistroC172 As ClassSpedRegistro, ByVal objSpedRegistro0200 As ClassSpedRegistro, ByVal objSpedRegistro0205 As ClassSpedRegistro, ByVal objSpedRegistro0190 As ClassSpedRegistro, ByVal objSpedRegistro0220 As ClassSpedRegistro, ByVal lSeqPaiNF As Long, ByRef tNFiscal As typeNFiscal, ByRef tTipoDocInfo As typeTipoDocInfo, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal iNFe As Integer, dTotalRateioNF As Double, ByVal colUM As Collection, ByVal colSpedE510 As Collection, dValDebICMS_objSpedC190 As Double, dValCredICMS_objSpedC190 As Double, colSpedC190 As Collection, ByVal objFilialEmpresa As AdmFiliais, ByRef tTributacaoNF As typeTributacaoNF, ByVal sData As String, _
'dTotalRateioNFICMSBase As Double, dTotalRateioNFICMSValor As Double, dTotalRateioNFICMSSubstBase As Double, dTotalRateioNFICMSSubstValor As Double, dTotalRateioNFIPIBaseCalculo As Double, dTotalRateioNFIPIValor As Double) As Long


Dim lErro As Long
Dim objSpedRegistro As ClassSpedRegistro
Dim alComando(1 To 5) As Long
Dim iIndice As Integer
Dim tItemNF As typeItemNF
Dim tTribItemNF As typeTributacaoDocItem
Dim iTipoTribCST_ICMS As Integer
Dim iIPICSTEntrada As Integer
Dim iIPICSTSaida As Integer
Dim iPISCST As Integer
Dim dPISAliquota As Double
Dim iCOFINSCST As Integer
Dim dCOFINSAliquota As Double
Dim tProduto As typeProduto
Dim objSpedE510Aux As ClassSpedE510Aux
Dim colSpedE510Aux As Collection
Dim dValorItemNFTotal As Double
Dim sUMResult As String
Dim lSeqPai As Long
Dim iISSIncide As Integer
Dim objSpedC190 As ClassSpedC190
Dim objSpedE510 As ClassSpedE510
Dim lSeq0200 As Long
Dim iAchou As Integer
Dim iTipoRegistro As Integer
Dim dTotalRateioNFICMSBase As Double
Dim objSpedC190UF As ClassSpedC190UF
Dim dIPIValor As Double, objTribItem As ClassTributacaoDocItem
Dim dValorTotalItem As Double, dICMSSubstValorExterno As Double
Dim iTipoTribCST_ICMSAux As Integer
Dim dValorFCP As Double
Dim dValorFCPST As Double
Dim dValorFCPRet As Double

On Error GoTo Erro_SpedFiscal_C170_NFCe
        
    Set colSpedC190 = New Collection
        
    If Len(Trim(sData)) > 0 Then sData = left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 4)
        
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203971
    Next

    '****************************** registro C170 *****************************************
    'se for nfce
    If tTipoDocInfo.iModeloArqICMS = 65 Then
    
        dTotalRateioNF = 0
    
        tItemNF.sProduto = String(STRING_PRODUTO, 0)
        tItemNF.sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
        tItemNF.sUnidadeMed = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
    
        lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Item, Produto, DescricaoItem, Quantidade, UnidadeMed, ClasseUM, PrecoUnitario, ValorDesconto FROM ItensNFiscal, Produtos  WHERE Produto = Codigo AND NumIntNF  = ?", _
        tItemNF.lNumIntDoc, tItemNF.iItem, tItemNF.sProduto, tItemNF.sDescricaoItem, tItemNF.dQuantidade, tItemNF.sUnidadeMed, tProduto.iClasseUM, tItemNF.dPrecoUnitario, tItemNF.dValorDesconto, tNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 203972
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203973
    
        Set colSpedE510Aux = New Collection
    
        Do While lErro = AD_SQL_SUCESSO
    
            tTribItemNF.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
            tTribItemNF.sCST = String(STRING_CST, 0)
     
            lErro = Comando_Executar(alComando(3), "SELECT CST, OrigemMercadoria, COFINSCredito, TipoTribCST, NaturezaOP, ICMSBase, ICMSAliquota, ICMSValor, ICMSCredito, ICMSSubstBase, ICMSSubstAliquota, ICMSSubstValor, IPI.CSTEntrada, IPI.CSTSaida, " & _
                "IPIBaseCalculo, IPIAliquota, IPIValor, IPICredito, PISCredito, TipoTributacao, ICMSPercRedBase, DespImpICMSBase, DespImpICMSValor, PISTipo, PISTipoCalculo, PISBase, PISAliquota, PISAliquotaValor, PISQtde, PISValor, COFINSTipo, COFINSTipoCalculo, COFINSBase, COFINSAliquota, COFINSAliquotaValor, COFINSQtde, COFINSValor, ValorFreteItem, ValorSeguroItem, ValorOutrasDespesasItem, DescontoGrid, ICMSvFCP, ICMSvFCPST, ICMSvFCPSTRet FROM TributacaoDocItem, tiposTribICMS AS ICMS, TiposTribIPI AS IPI  WHERE NumIntDoc = ? AND Item = ? AND ICMS.Tipo = ICMSTipo AND IPI.Tipo = IPITipo", _
                tTribItemNF.sCST, tTribItemNF.iOrigemMercadoria, tTribItemNF.dCOFINSCredito, iTipoTribCST_ICMSAux, tTribItemNF.sNaturezaOp, tTribItemNF.dICMSBase, tTribItemNF.dICMSAliquota, tTribItemNF.dICMSValor, tTribItemNF.dICMSCredito, tTribItemNF.dICMSSubstBase, tTribItemNF.dICMSSubstAliquota, tTribItemNF.dICMSSubstValor, _
                iIPICSTEntrada, iIPICSTSaida, tTribItemNF.dIPIBaseCalculo, tTribItemNF.dIPIAliquota, tTribItemNF.dIPIValor, tTribItemNF.dIPICredito, tTribItemNF.dPISCredito, tTribItemNF.iTipoTributacao, tTribItemNF.dICMSPercRedBase, tTribItemNF.dDespImpICMSBase, tTribItemNF.dDespImpICMSValor, _
                tTribItemNF.iPISTipo, tTribItemNF.iPISTipoCalculo, tTribItemNF.dPISBase, tTribItemNF.dPISAliquota, tTribItemNF.dPISAliquotaValor, tTribItemNF.dPISQtde, tTribItemNF.dPISValor, tTribItemNF.iCOFINSTipo, tTribItemNF.iCOFINSTipoCalculo, tTribItemNF.dCOFINSBase, tTribItemNF.dCOFINSAliquota, tTribItemNF.dCOFINSAliquotaValor, tTribItemNF.dCOFINSQtde, tTribItemNF.dCOFINSValor, tTribItemNF.dValorFreteItem, tTribItemNF.dValorSeguroItem, tTribItemNF.dValorOutrasDespesasItem, tTribItemNF.dDescontoGrid, dValorFCP, dValorFCPST, dValorFCPRet, _
                tNFiscal.lNumIntDoc, tItemNF.iItem)
            If lErro <> AD_SQL_SUCESSO Then gError 203974
        
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203975
    
            iTipoTribCST_ICMS = iTipoTribCST_ICMSAux

            dICMSSubstValorExterno = tTribItemNF.dICMSSubstValor

            Set objTribItem = New ClassTributacaoDocItem
            objTribItem.lNumIntDocItem = tItemNF.lNumIntDoc
            objTribItem.iTipoDoc = TIPODOC_TRIB_NF
            lErro = ItemNF_Le_BasePropria(objTribItem, iTipoTribCST_ICMS, iIPICSTEntrada)
            If lErro <> SUCESSO Then gError 203976
            
            iTipoTribCST_ICMS = iTipoTribCST_ICMS + (tTribItemNF.iOrigemMercadoria * 100)
            
            tTribItemNF.dICMSBase = objTribItem.dICMSBase
            tTribItemNF.dICMSAliquota = objTribItem.dICMSAliquota
            tTribItemNF.dICMSValor = objTribItem.dICMSValor
            tTribItemNF.dICMSCredito = objTribItem.dICMSCredito
            
            tTribItemNF.dICMSSubstBase = objTribItem.dICMSSubstBase
            tTribItemNF.dICMSSubstAliquota = objTribItem.dICMSSubstAliquota
            tTribItemNF.dICMSSubstValor = objTribItem.dICMSSubstValor
            tTribItemNF.dICMSSubstPercMVA = objTribItem.dICMSSubstPercMVA
            tTribItemNF.dICMSSubstPercRedBase = objTribItem.dICMSSubstPercRedBase
        
            dIPIValor = tTribItemNF.dIPIValor

            tTribItemNF.dIPIBaseCalculo = objTribItem.dIPIBaseCalculo
            tTribItemNF.dIPIAliquota = objTribItem.dIPIAliquota
            tTribItemNF.dIPIValor = objTribItem.dIPIValor
            tTribItemNF.dIPICredito = objTribItem.dIPICredito
    
            iAchou = 0
    
            For Each objSpedC190 In colSpedC190
            
                If objSpedC190.iCSTICMS = iTipoTribCST_ICMS And objSpedC190.sCFOP = tTribItemNF.sNaturezaOp And objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota Then
                    
                    dValorTotalItem = Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + dICMSSubstValorExterno)
                    If tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC Then
                        If tTribItemNF.dICMSBase <> 0 And tTribItemNF.dICMSBase > dValorTotalItem And gobjCRFAT.iNFImportacaoTribFlag03 = DESMARCADO Then
                            dValorTotalItem = tTribItemNF.dICMSBase
                        Else
                            If gobjCRFAT.iNFImportacaoTribFlag04 = DESMARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dICMSValor
                            If gobjCRFAT.iNFImportacaoTribFlag08 = MARCADO Then dValorTotalItem = dValorTotalItem - tTribItemNF.dValorOutrasDespesasItem
                            If gobjCRFAT.iNFImportacaoTribFlag05 = MARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dIIValor
                            
                            If gobjCRFAT.iNFImportacaoTribFlag12 = MARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dPISValor + tTribItemNF.dCOFINSValor

                        End If
                    End If
                    
                    objSpedC190.dValorOper = objSpedC190.dValorOper + dValorTotalItem 'Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + tTribItemNF.dICMSSubstValor)
                    dTotalRateioNF = dTotalRateioNF + Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + dICMSSubstValorExterno)
                    

                    
                    objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS + tTribItemNF.dICMSBase

                    
'mario                    objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dICMSValor
                    objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dICMSCredito
'                    If tNFiscal.iTipoNFiscal = 119 Then
'                        objSpedC190.dValorICMS = objSpedC190.dValorICMS + tTribItemNF.dDespImpICMSValor
'                    End If

                    
                    objSpedC190.dValorRedBC = objSpedC190.dValorRedBC + Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase, 2) '??????? esta certo????
                    objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS - Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase, 2) 'objSpedC190.dValorRedBC
                    objSpedC190.dValorFCP = objSpedC190.dValorFCP + dValorFCP
                    objSpedC190.dValorFCPST = objSpedC190.dValorFCPST + dValorFCPST
                    objSpedC190.dValorFCPRet = objSpedC190.dValorFCPRet + dValorFCPRet
                    iAchou = 1
                    Exit For
                End If
                
            Next
    
            If iAchou = 0 Then
            
                Set objSpedC190 = New ClassSpedC190
                
                objSpedC190.iCSTICMS = iTipoTribCST_ICMS
                objSpedC190.sCFOP = tTribItemNF.sNaturezaOp
                objSpedC190.dAliquotaICMS = tTribItemNF.dICMSAliquota
                
                dValorTotalItem = Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + dICMSSubstValorExterno)
                If tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or tNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC Then
                    If tTribItemNF.dICMSBase <> 0 And tTribItemNF.dICMSBase > dValorTotalItem And gobjCRFAT.iNFImportacaoTribFlag03 = DESMARCADO Then
                        dValorTotalItem = tTribItemNF.dICMSBase
                    Else
                        If gobjCRFAT.iNFImportacaoTribFlag04 = DESMARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dICMSValor
                        If gobjCRFAT.iNFImportacaoTribFlag08 = MARCADO Then dValorTotalItem = dValorTotalItem - tTribItemNF.dValorOutrasDespesasItem
                        If gobjCRFAT.iNFImportacaoTribFlag05 = MARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dIIValor
                    
                        If gobjCRFAT.iNFImportacaoTribFlag12 = MARCADO Then dValorTotalItem = dValorTotalItem + tTribItemNF.dPISValor + tTribItemNF.dCOFINSValor
                    
                    End If
                End If
                
                objSpedC190.dValorOper = dValorTotalItem 'Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + tTribItemNF.dICMSSubstValor)
                dTotalRateioNF = dTotalRateioNF + Arredonda_Moeda((tItemNF.dQuantidade * tItemNF.dPrecoUnitario) + tTribItemNF.dValorFreteItem + tTribItemNF.dValorOutrasDespesasItem + tTribItemNF.dValorSeguroItem - tTribItemNF.dDescontoGrid + dIPIValor + dICMSSubstValorExterno)
                
                objSpedC190.dValorBCICMS = tTribItemNF.dICMSBase
                objSpedC190.dValorICMS = tTribItemNF.dICMSCredito
                objSpedC190.dValorRedBC = Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase)
                objSpedC190.dValorBCICMS = objSpedC190.dValorBCICMS - Arredonda_Moeda(tTribItemNF.dICMSBase * tTribItemNF.dICMSPercRedBase) 'objSpedC190.dValorRedBC
                objSpedC190.dValorFCP = dValorFCP
                objSpedC190.dValorFCPST = dValorFCPST
                objSpedC190.dValorFCPRet = dValorFCPRet
                
                colSpedC190.Add objSpedC190
                
            End If
    
            iAchou = 0

            For Each objSpedC190UF In colSpedC190UF
            
                If objSpedC190UF.sUF = tNFiscal.sUFSubstTrib And objSpedC190UF.sCFOP = tTribItemNF.sNaturezaOp Then
                
                    objSpedC190UF.dValorBCICMSST = objSpedC190UF.dValorBCICMSST + tTribItemNF.dICMSSubstBase
                    objSpedC190UF.dValorICMSST = objSpedC190UF.dValorICMSST + tTribItemNF.dICMSSubstValor
                    
                    iAchou = 1
                    Exit For
                End If
                
            Next
    
            If iAchou = 0 Then
            
                Set objSpedC190UF = New ClassSpedC190UF
                
                objSpedC190UF.sUF = tNFiscal.sUFSubstTrib
                objSpedC190UF.sCFOP = tTribItemNF.sNaturezaOp
                
                objSpedC190UF.dValorBCICMSST = tTribItemNF.dICMSSubstBase
                objSpedC190UF.dValorICMSST = tTribItemNF.dICMSSubstValor
                                
                colSpedC190UF.Add objSpedC190UF
                
            End If
    
            'soma os valores de ICMS dos registros C100 de saida que estiverem com data de saida dentro do periodo para poder colocar
            'posteriormente no registro E110
            If Len(Trim(sData)) > 0 Then
                If CDate(sData) >= dtDataIni And CDate(sData) <= dtDataFim And tTipoDocInfo.iTipo = 2 Then
    'mario                dValDebICMS_objSpedC190 = dValDebICMS_objSpedC190 + tTribItemNF.dICMSValor
                    dValDebICMS_objSpedC190 = dValDebICMS_objSpedC190 + tTribItemNF.dICMSCredito
                End If
            End If
    
            'soma os valores de ICMS dos registros C100 de entrada que estiverem com data de entrada dentro do periodo para poder colocar
            'posteriormente no registro E110
            If Len(Trim(sData)) > 0 Then
                If CDate(sData) >= dtDataIni And CDate(sData) <= dtDataFim And (tTipoDocInfo.iTipo = 1 Or tTipoDocInfo.iTipo = 3) Then
    '                dValCredICMS_objSpedC190 = dValCredICMS_objSpedC190 + tTribItemNF.dICMSValor
                    dValCredICMS_objSpedC190 = dValCredICMS_objSpedC190 + tTribItemNF.dICMSCredito
    '                If tNFiscal.iTipoNFiscal = 119 Then
    '                    dValCredICMS_objSpedC190 = dValCredICMS_objSpedC190 + tTribItemNF.dDespImpICMSValor
    '                End If
                    
                End If
            End If
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203988
    
        Loop
        
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    SpedFiscal_C170_NFCe = SUCESSO

    Exit Function
    
Erro_SpedFiscal_C170_NFCe:

   SpedFiscal_C170_NFCe = gErr

   Select Case gErr
    
        Case 203971
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 203972, 203973, 203988
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)
    
        Case 203974, 203975
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOITEMNF", gErr)
    
        Case 203976 To 203981, 204190 To 204192, 204206
    
        Case 203982, 203986
            Call Rotina_Erro(vbOKOnly, "ERRO_SPEDFISCAL_REGISTRO_MAIOR_MAXIMO", gErr)
    
        Case 203983, 203987
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDFISCAL", gErr)
        
        Case 203984, 203985
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBUTACAOMOVTO", gErr)
        
        Case 204190 To 204192
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLNF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203989)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

End Function



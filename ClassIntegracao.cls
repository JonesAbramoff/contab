VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIntegracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function Conexao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Private Declare Function Conexao_FecharExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_PrepararPosInt Lib "ADSQLMN.DLL" Alias "AD_Comando_PrepararPos" (ByVal lComando As Long, ByVal lpSQLStmt As String, ByVal lSelect As Long) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Dim glConexao1 As Long
Dim glConexao2 As Long

Dim gcolLog As New Collection

Dim giValidacaoManual As Integer

Dim gsDiretorioImport As String
Dim gsDiretorioImportTemp As String
Dim gsDiretorioExport As String
Dim gsDiretorioExportTemp As String

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function IntegracaoCliente_Le_Aux(tIntegracaoCliente As typeIntegracaoCliente, sTodosCampos As String) As Long

Dim lErro As Long

On Error GoTo Erro_IntegracaoCliente_Le_Aux

    'Alocação de espaço no buffer
    tIntegracaoCliente.sRazaoSocial = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sNomeReduzido = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sObservacao = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sFilialNome = String(STRING_FILIAL_NOME, 0)
    tIntegracaoCliente.sCgc = String(STRING_CGC, 0)
    tIntegracaoCliente.sRG = String(STRING_RG, 0)
    tIntegracaoCliente.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tIntegracaoCliente.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    
    tIntegracaoCliente.sEndereco = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sBairro = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sCidade = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sSiglaEstado = String(STRING_ESTADO, 0)
    tIntegracaoCliente.sCEP = String(STRING_CEP, 0)
    tIntegracaoCliente.sTelefone1 = String(STRING_TELEFONE, 0)
    tIntegracaoCliente.sTelefone2 = String(STRING_TELEFONE, 0)
    tIntegracaoCliente.sEmail = String(STRING_EMAIL, 0)
    tIntegracaoCliente.sFax = String(STRING_FAX, 0)
    tIntegracaoCliente.sContato = String(STRING_CONTATO, 0)
    
    tIntegracaoCliente.sEnderecoEnt = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sBairroEnt = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sCidadeEnt = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sSiglaEstadoEnt = String(STRING_ESTADO, 0)
    tIntegracaoCliente.sCEPEnt = String(STRING_CEP, 0)
    tIntegracaoCliente.sTelefone1Ent = String(STRING_TELEFONE, 0)
    tIntegracaoCliente.sTelefone2Ent = String(STRING_TELEFONE, 0)
    tIntegracaoCliente.sEmailEnt = String(STRING_EMAIL, 0)
    tIntegracaoCliente.sFaxEnt = String(STRING_FAX, 0)
    tIntegracaoCliente.sContatoEnt = String(STRING_CONTATO, 0)
    
    tIntegracaoCliente.sEnderecoCobr = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sBairroCobr = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sCidadeCobr = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sSiglaEstadoCobr = String(STRING_ESTADO, 0)
    tIntegracaoCliente.sCEPCobr = String(STRING_CEP, 0)
    tIntegracaoCliente.sTelefone1Cobr = String(STRING_TELEFONE, 0)
    tIntegracaoCliente.sTelefone2Cobr = String(STRING_TELEFONE, 0)
    tIntegracaoCliente.sEmailCobr = String(STRING_EMAIL, 0)
    tIntegracaoCliente.sFaxCobr = String(STRING_FAX, 0)
    tIntegracaoCliente.sContatoCobr = String(STRING_CONTATO, 0)
    tIntegracaoCliente.sTipoCliente = String(STRING_INT_TIPOCLIENTE_DESC, 0)

    tIntegracaoCliente.sObservacaoFilial = String(STRING_CLIENTE_OBSERVACAO, 0)
    
    
    tIntegracaoCliente.sLogradouro = String(STRING_ENDERECO_LOGRADOURO, 0)
    tIntegracaoCliente.sComplemento = String(STRING_ENDERECO_COMPLEMENTO, 0)
    tIntegracaoCliente.sTipoLogradouro = String(STRING_ENDERECO_TIPOLOGRADOURO, 0)
    tIntegracaoCliente.sEmail2 = String(STRING_EMAIL, 0)
    tIntegracaoCliente.sTelNumero1 = String(STRING_TELEFONE, 0)
    tIntegracaoCliente.sTelNumero2 = String(STRING_TELEFONE, 0)
    tIntegracaoCliente.sFaxNumero = String(STRING_FAX, 0)
    tIntegracaoCliente.sLogradouroEnt = String(STRING_ENDERECO_LOGRADOURO, 0)
    tIntegracaoCliente.sComplementoEnt = String(STRING_ENDERECO_COMPLEMENTO, 0)
    tIntegracaoCliente.sTipoLogradouroEnt = String(STRING_ENDERECO_TIPOLOGRADOURO, 0)
    tIntegracaoCliente.sEmail2Ent = String(STRING_EMAIL, 0)
    tIntegracaoCliente.sTelNumero1Ent = String(STRING_TELEFONE, 0)
    tIntegracaoCliente.sTelNumero2Ent = String(STRING_TELEFONE, 0)
    tIntegracaoCliente.sFaxNumeroEnt = String(STRING_FAX, 0)
    tIntegracaoCliente.sReferencia = String(STRING_MAXIMO, 0)
    tIntegracaoCliente.sReferenciaEnt = String(STRING_MAXIMO, 0)
    
    sTodosCampos = "NumIntDoc, NumIntGer, NumIntArq, SeqRegistro, TipoInt, CodCliente, " & _
                "CodFilial, RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, " & _
                "Ativo, FilialNome, CGC, RG, InscricaoEstadual, InscricaoMunicipal, Endereco, " & _
                "Bairro, Cidade, SiglaEstado, CodigoPais, CEP, Telefone1, Telefone2, " & _
                "Email, Fax, Contato, EnderecoEnt, BairroEnt, CidadeEnt, SiglaEstadoEnt, " & _
                "CodigoPaisEnt, CEPEnt, Telefone1Ent, Telefone2Ent, EmailEnt, FaxEnt, ContatoEnt, " & _
                "EnderecoCobr, BairroCobr, CidadeCobr, SiglaEstadoCobr, CodigoPaisCobr, CEPCobr, Telefone1Cobr, " & _
                "Telefone2Cobr, EmailCobr, FaxCobr, ContatoCobr, ComErro, DataAtualizacao, Vendedor, ObservacaoFilial, CodExterno, TipoCliente"
   
   
    sTodosCampos = sTodosCampos & ", Logradouro, Complemento, TipoLogradouro, Email2, Numero , TelDDD1, TelDDD2, FaxDDD, TelNumero1, TelNumero2, FaxNumero, LogradouroEnt, ComplementoEnt, TipoLogradouroEnt, Email2Ent, NumeroEnt, TelDDD1Ent, TelDDD2Ent, FaxDDDEnt, TelNumero1Ent, TelNumero2Ent, FaxNumeroEnt, TabelaPreco, Referencia, ReferenciaEnt"
   
    IntegracaoCliente_Le_Aux = SUCESSO

    Exit Function

Erro_IntegracaoCliente_Le_Aux:

    IntegracaoCliente_Le_Aux = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189905)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function IntegracaoPV_Le_Aux(tIntegracaoPV As typeIntegracaoPV, sTodosCampos As String) As Long

Dim lErro As Long

On Error GoTo Erro_IntegracaoPV_Le_Aux

    'Alocação de espaço no buffer
    tIntegracaoPV.sCGCCliente = String(STRING_CGC, 0)
    tIntegracaoPV.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tIntegracaoPV.sNomeCondPagtoExterno = String(STRING_NOME_EXTERNO_INT, 0)
    tIntegracaoPV.sNomeTransportadoraExterno = String(STRING_NOME_EXTERNO_INT, 0)
    tIntegracaoPV.sMensagemPedido = String(STRING_PEDIDOVENDA_MENSAGEM_NOTA, 0)
    tIntegracaoPV.sNomeVendedorExterno = String(STRING_NOME_EXTERNO_INT, 0)
    tIntegracaoPV.sNomeAlmoxarifadoExterno = String(STRING_NOME_EXTERNO_INT, 0)
    tIntegracaoPV.sCodProduto = String(STRING_PRODUTO, 0)
    tIntegracaoPV.sCodProdutoExterno = String(STRING_NOME_EXTERNO_INT, 0)
    tIntegracaoPV.sDescricaoItem = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tIntegracaoPV.sUM = String(STRING_UM_SIGLA, 0)
    tIntegracaoPV.sPedRepr = String(STRING_PEDIDOVENDA_PEDIDO_REPRESENTANTE, 0)
    
    sTodosCampos = "NumIntDoc, NumIntGer, NumIntArq, SeqRegistro, TipoInt, CodPedidoExterno, " & _
                "DataEmissao, CodClienteExterno, CodClienteCorporator, CodFilialCorporator, CGCCliente, NaturezaOP, CodTabelaPreco, " & _
                "CodCondPagto, NomeCondPagtoExterno, CodCondPagtoExterno, FilialEmpresaFat, ValorDescontoPedido, Frete, ValorFretePedido, " & _
                "ValorSeguroPedido, ValorOutrasDespesasPedido, FilialEmpresaEnt, TrazerTranspAuto, CodTransportadora, CodTransportadoraExterno, NomeTransportadoraExterno, " & _
                "TrazerMensagemAuto, MensagemPedido, TrazerPesoAuto, PesoBruto, PesoLiquido, TrazerCanalVendaAuto, CanalVenda, " & _
                "CodVendedor, CodVendedorExterno, NomeVendedorExterno, TrazerComissaoAuto, ValorComissao, TrazerReservaAuto, CodAlmoxarifado, " & _
                "CodAlmoxarifadoExterno, NomeAlmoxarifadoExterno, Item, CodProduto, CodProdutoExterno, TrazerDescricaoAuto, DescricaoItem, " & _
                "QuantidadePedida, QuantidadeCancelada, UM, PrecoUnitario, ValorDescontoItem, DataEntrega, ComErro, " & _
                "DataAtualizacao, CodTabelaPrecoItem, PedRepr "
   
    IntegracaoPV_Le_Aux = SUCESSO

    Exit Function

Erro_IntegracaoPV_Le_Aux:

    IntegracaoPV_Le_Aux = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190051)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function IntegracaoSldProd_Le_Aux(tIntegracaoSldProd As typeIntegracaoSldProd, sTodosCampos As String) As Long

Dim lErro As Long

On Error GoTo Erro_IntegracaoSldProd_Le_Aux

    'Alocação de espaço no buffer
    tIntegracaoSldProd.sCodProduto = String(STRING_PRODUTO, 0)

    sTodosCampos = "NumIntDoc, NumIntGer, NumIntArq, SeqRegistro, TipoInt, CodProduto, Almoxarifado, SaldoDisp, ComErro, DataAtualizacao "
   
    IntegracaoSldProd_Le_Aux = SUCESSO

    Exit Function

Erro_IntegracaoSldProd_Le_Aux:

    IntegracaoSldProd_Le_Aux = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190138)

    End Select

    Exit Function

End Function


Private Function IntegracaoCliente_Le_SemArq(ByVal lComando As Long, tIntegracaoCliente As typeIntegracaoCliente, ByVal objIntegracaoCliente As ClassIntegracaoCliente) As Long

Dim lErro As Long
Dim sTodosCampos As String

On Error GoTo Erro_IntegracaoCliente_Le_SemArq

    lErro = IntegracaoCliente_Le_Aux(tIntegracaoCliente, sTodosCampos)
    If lErro <> SUCESSO Then gError 189906

    'Le a tabelaIntegracaoCliente
    lErro = Comando_ExecutarPos(lComando, "SELECT " & sTodosCampos & " FROM IntegracaoCliente WHERE NumIntArq = ? ORDER BY NumIntGer, SeqRegistro", 0, _
                tIntegracaoCliente.lNumIntDoc, tIntegracaoCliente.lNumIntGer, tIntegracaoCliente.lNumIntArq, tIntegracaoCliente.lSeqRegistro, _
                tIntegracaoCliente.iTipoInt, tIntegracaoCliente.lCodCliente, tIntegracaoCliente.iCodFilial, tIntegracaoCliente.sRazaoSocial, tIntegracaoCliente.sNomeReduzido, _
                tIntegracaoCliente.iTipo, tIntegracaoCliente.sObservacao, tIntegracaoCliente.dLimiteCredito, tIntegracaoCliente.iCondicaoPagto, tIntegracaoCliente.iAtivo, _
                tIntegracaoCliente.sFilialNome, tIntegracaoCliente.sCgc, tIntegracaoCliente.sRG, tIntegracaoCliente.sInscricaoEstadual, tIntegracaoCliente.sInscricaoMunicipal, _
                tIntegracaoCliente.sEndereco, tIntegracaoCliente.sBairro, tIntegracaoCliente.sCidade, tIntegracaoCliente.sSiglaEstado, tIntegracaoCliente.iCodigoPais, _
                tIntegracaoCliente.sCEP, tIntegracaoCliente.sTelefone1, tIntegracaoCliente.sTelefone2, tIntegracaoCliente.sEmail, tIntegracaoCliente.sFax, _
                tIntegracaoCliente.sContato, tIntegracaoCliente.sEnderecoEnt, tIntegracaoCliente.sBairroEnt, tIntegracaoCliente.sCidadeEnt, tIntegracaoCliente.sSiglaEstadoEnt, _
                tIntegracaoCliente.iCodigoPaisEnt, tIntegracaoCliente.sCEPEnt, tIntegracaoCliente.sTelefone1Ent, tIntegracaoCliente.sTelefone2Ent, tIntegracaoCliente.sEmailEnt, _
                tIntegracaoCliente.sFaxEnt, tIntegracaoCliente.sContatoEnt, tIntegracaoCliente.sEnderecoCobr, tIntegracaoCliente.sBairroCobr, tIntegracaoCliente.sCidadeCobr, _
                tIntegracaoCliente.sSiglaEstadoCobr, tIntegracaoCliente.iCodigoPaisCobr, tIntegracaoCliente.sCEPCobr, tIntegracaoCliente.sTelefone1Cobr, tIntegracaoCliente.sTelefone2Cobr, _
                tIntegracaoCliente.sEmailCobr, tIntegracaoCliente.sFaxCobr, tIntegracaoCliente.sContatoCobr, tIntegracaoCliente.iComErro, tIntegracaoCliente.dtDataAtualizacao, tIntegracaoCliente.iVendedor, tIntegracaoCliente.sObservacaoFilial, tIntegracaoCliente.lCodExterno, tIntegracaoCliente.sTipoCliente, _
                tIntegracaoCliente.sLogradouro, tIntegracaoCliente.sComplemento, tIntegracaoCliente.sTipoLogradouro, _
                tIntegracaoCliente.sEmail2, tIntegracaoCliente.lNumero, tIntegracaoCliente.iTelDDD1, _
                tIntegracaoCliente.iTelDDD2, tIntegracaoCliente.iFaxDDD, tIntegracaoCliente.sTelNumero1, _
                tIntegracaoCliente.sTelNumero2, tIntegracaoCliente.sFaxNumero, tIntegracaoCliente.sLogradouroEnt, _
                tIntegracaoCliente.sComplementoEnt, tIntegracaoCliente.sTipoLogradouroEnt, tIntegracaoCliente.sEmail2Ent, _
                tIntegracaoCliente.lNumeroEnt, tIntegracaoCliente.iTelDDD1Ent, tIntegracaoCliente.iTelDDD2Ent, _
                tIntegracaoCliente.iFaxDDDEnt, tIntegracaoCliente.sTelNumero1Ent, tIntegracaoCliente.sTelNumero2Ent, _
                tIntegracaoCliente.sFaxNumeroEnt, tIntegracaoCliente.iTabelaPreco, 0)
    If lErro <> AD_SQL_SUCESSO Then gError 189736
    
    IntegracaoCliente_Le_SemArq = SUCESSO

    Exit Function

Erro_IntegracaoCliente_Le_SemArq:

    IntegracaoCliente_Le_SemArq = gErr

    Select Case gErr

        Case 189736
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INTEGRACAOCLIENTE", gErr)
            
        Case 189906

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189737)

    End Select

    Exit Function

End Function

Private Function IntegracaoSldProd_Le_SemArq(ByVal lComando As Long, tIntegracaoSldProd As typeIntegracaoSldProd, ByVal objIntegracaoSldProd As ClassIntegracaoSldProd) As Long

Dim lErro As Long
Dim sTodosCampos As String

On Error GoTo Erro_IntegracaoSldProd_Le_SemArq

    lErro = IntegracaoSldProd_Le_Aux(tIntegracaoSldProd, sTodosCampos)
    If lErro <> SUCESSO Then gError 190134

    'Le a tabelaIntegracaoSldProd
    lErro = Comando_ExecutarPos(lComando, "SELECT " & sTodosCampos & " FROM IntegracaoSldProd WHERE NumIntArq = ? ORDER BY NumIntGer, SeqRegistro", 0, _
                tIntegracaoSldProd.lNumIntDoc, tIntegracaoSldProd.lNumIntGer, tIntegracaoSldProd.lNumIntArq, tIntegracaoSldProd.lSeqRegistro, _
                tIntegracaoSldProd.iTipoInt, tIntegracaoSldProd.sCodProduto, tIntegracaoSldProd.iAlmoxarifado, tIntegracaoSldProd.dSaldoDisp, tIntegracaoSldProd.iComErro, tIntegracaoSldProd.dtDataAtualizacao, _
                0)
    If lErro <> AD_SQL_SUCESSO Then gError 190135
    
    IntegracaoSldProd_Le_SemArq = SUCESSO

    Exit Function

Erro_IntegracaoSldProd_Le_SemArq:

    IntegracaoSldProd_Le_SemArq = gErr

    Select Case gErr

        Case 190134
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INTEGRACAOSldProd", gErr)
            
        Case 190135

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190136)

    End Select

    Exit Function

End Function

Private Function IntegracaoCliente_Le_ImpSemAtu(ByVal lComando As Long, tIntegracaoCliente As typeIntegracaoCliente, ByVal objIntegracaoCliente As ClassIntegracaoCliente) As Long

Dim lErro As Long
Dim sTodosCampos As String

On Error GoTo Erro_IntegracaoCliente_Le_ImpSemAtu

    lErro = IntegracaoCliente_Le_Aux(tIntegracaoCliente, sTodosCampos)
    If lErro <> SUCESSO Then gError 189907

    'Le a tabelaIntegracaoCliente
    lErro = Comando_Executar(lComando, "SELECT " & sTodosCampos & " FROM IntegracaoCliente WHERE TipoInt = ? AND DataAtualizacao = ? AND NumIntArq = ? ORDER BY NumIntArq, SeqRegistro", _
                tIntegracaoCliente.lNumIntDoc, tIntegracaoCliente.lNumIntGer, tIntegracaoCliente.lNumIntArq, tIntegracaoCliente.lSeqRegistro, _
                tIntegracaoCliente.iTipoInt, tIntegracaoCliente.lCodCliente, tIntegracaoCliente.iCodFilial, tIntegracaoCliente.sRazaoSocial, tIntegracaoCliente.sNomeReduzido, _
                tIntegracaoCliente.iTipo, tIntegracaoCliente.sObservacao, tIntegracaoCliente.dLimiteCredito, tIntegracaoCliente.iCondicaoPagto, tIntegracaoCliente.iAtivo, _
                tIntegracaoCliente.sFilialNome, tIntegracaoCliente.sCgc, tIntegracaoCliente.sRG, tIntegracaoCliente.sInscricaoEstadual, tIntegracaoCliente.sInscricaoMunicipal, _
                tIntegracaoCliente.sEndereco, tIntegracaoCliente.sBairro, tIntegracaoCliente.sCidade, tIntegracaoCliente.sSiglaEstado, tIntegracaoCliente.iCodigoPais, _
                tIntegracaoCliente.sCEP, tIntegracaoCliente.sTelefone1, tIntegracaoCliente.sTelefone2, tIntegracaoCliente.sEmail, tIntegracaoCliente.sFax, _
                tIntegracaoCliente.sContato, tIntegracaoCliente.sEnderecoEnt, tIntegracaoCliente.sBairroEnt, tIntegracaoCliente.sCidadeEnt, tIntegracaoCliente.sSiglaEstadoEnt, _
                tIntegracaoCliente.iCodigoPaisEnt, tIntegracaoCliente.sCEPEnt, tIntegracaoCliente.sTelefone1Ent, tIntegracaoCliente.sTelefone2Ent, tIntegracaoCliente.sEmailEnt, _
                tIntegracaoCliente.sFaxEnt, tIntegracaoCliente.sContatoEnt, tIntegracaoCliente.sEnderecoCobr, tIntegracaoCliente.sBairroCobr, tIntegracaoCliente.sCidadeCobr, _
                tIntegracaoCliente.sSiglaEstadoCobr, tIntegracaoCliente.iCodigoPaisCobr, tIntegracaoCliente.sCEPCobr, tIntegracaoCliente.sTelefone1Cobr, tIntegracaoCliente.sTelefone2Cobr, _
                tIntegracaoCliente.sEmailCobr, tIntegracaoCliente.sFaxCobr, tIntegracaoCliente.sContatoCobr, tIntegracaoCliente.iComErro, tIntegracaoCliente.dtDataAtualizacao, tIntegracaoCliente.iVendedor, tIntegracaoCliente.sObservacaoFilial, tIntegracaoCliente.lCodExterno, tIntegracaoCliente.sTipoCliente, _
                tIntegracaoCliente.sLogradouro, tIntegracaoCliente.sComplemento, tIntegracaoCliente.sTipoLogradouro, _
                tIntegracaoCliente.sEmail2, tIntegracaoCliente.lNumero, tIntegracaoCliente.iTelDDD1, _
                tIntegracaoCliente.iTelDDD2, tIntegracaoCliente.iFaxDDD, tIntegracaoCliente.sTelNumero1, _
                tIntegracaoCliente.sTelNumero2, tIntegracaoCliente.sFaxNumero, tIntegracaoCliente.sLogradouroEnt, _
                tIntegracaoCliente.sComplementoEnt, tIntegracaoCliente.sTipoLogradouroEnt, tIntegracaoCliente.sEmail2Ent, _
                tIntegracaoCliente.lNumeroEnt, tIntegracaoCliente.iTelDDD1Ent, tIntegracaoCliente.iTelDDD2Ent, _
                tIntegracaoCliente.iFaxDDDEnt, tIntegracaoCliente.sTelNumero1Ent, tIntegracaoCliente.sTelNumero2Ent, _
                tIntegracaoCliente.sFaxNumeroEnt, tIntegracaoCliente.iTabelaPreco, _
                TIPO_INTEGRACAO_IMPORTACAO, DATA_NULA, objIntegracaoCliente.lNumIntArq)
    If lErro <> AD_SQL_SUCESSO Then gError 189908
    
    IntegracaoCliente_Le_ImpSemAtu = SUCESSO

    Exit Function

Erro_IntegracaoCliente_Le_ImpSemAtu:

    IntegracaoCliente_Le_ImpSemAtu = gErr

    Select Case gErr
    
        Case 189907

        Case 189908
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INTEGRACAOCLIENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189909)

    End Select

    Exit Function

End Function

Private Function IntegracaoPV_Le_ImpSemAtu(ByVal lComando As Long, tIntegracaoPV As typeIntegracaoPV, ByVal objIntegracaoPV As ClassIntegracaoPV) As Long

Dim lErro As Long
Dim sTodosCampos As String

On Error GoTo Erro_IntegracaoPV_Le_ImpSemAtu

    lErro = IntegracaoPV_Le_Aux(tIntegracaoPV, sTodosCampos)
    If lErro <> SUCESSO Then gError 190048

    'Le a tabelaIntegracaoPV
    lErro = Comando_Executar(lComando, "SELECT " & sTodosCampos & " FROM IntegracaoPV WHERE TipoInt = ? AND DataAtualizacao = ? AND NumIntArq = ? AND FilialEmpresa IN (?,?) ORDER BY NumIntArq, CodPedidoExterno, Item,  SeqRegistro", _
                tIntegracaoPV.lNumIntDoc, tIntegracaoPV.lNumIntGer, tIntegracaoPV.lNumIntArq, tIntegracaoPV.lSeqRegistro, _
                tIntegracaoPV.iTipoInt, tIntegracaoPV.lCodPedidoExterno, tIntegracaoPV.dtDataEmissao, tIntegracaoPV.lCodClienteExterno, tIntegracaoPV.lCodClienteCorporator, _
                tIntegracaoPV.iCodFilialCorporator, tIntegracaoPV.sCGCCliente, tIntegracaoPV.sNaturezaOp, tIntegracaoPV.iCodTabelaPreco, tIntegracaoPV.iCodCondPagto, _
                tIntegracaoPV.sNomeCondPagtoExterno, tIntegracaoPV.iCodCondPagtoExterno, tIntegracaoPV.iFilialEmpresaFat, tIntegracaoPV.dValorDescontoPedido, tIntegracaoPV.iFrete, _
                tIntegracaoPV.dValorFretePedido, tIntegracaoPV.dValorSeguroPedido, tIntegracaoPV.dValorOutrasDespesasPedido, tIntegracaoPV.iFilialEmpresaEnt, tIntegracaoPV.iTrazerTranspAuto, _
                tIntegracaoPV.iCodTransportadora, tIntegracaoPV.iCodTransportadoraExterno, tIntegracaoPV.sNomeTransportadoraExterno, tIntegracaoPV.iTrazerMensagemAuto, tIntegracaoPV.sMensagemPedido, _
                tIntegracaoPV.iTrazerPesoAuto, tIntegracaoPV.dPesoBruto, tIntegracaoPV.dPesoLiquido, tIntegracaoPV.iTrazerCanalVendaAuto, tIntegracaoPV.iCanalVenda, _
                tIntegracaoPV.iCodVendedor, tIntegracaoPV.iCodVendedorExterno, tIntegracaoPV.sNomeVendedorExterno, tIntegracaoPV.iTrazerComissaoAuto, tIntegracaoPV.dValorComissao, _
                tIntegracaoPV.iTrazerReservaAuto, tIntegracaoPV.iCodAlmoxarifado, tIntegracaoPV.iCodAlmoxarifadoExterno, tIntegracaoPV.sNomeAlmoxarifadoExterno, tIntegracaoPV.iItem, _
                tIntegracaoPV.sCodProduto, tIntegracaoPV.sCodProdutoExterno, tIntegracaoPV.iTrazerDescricaoAuto, tIntegracaoPV.sDescricaoItem, tIntegracaoPV.dQuantidadePedida, _
                tIntegracaoPV.dQuantidadeCancelada, tIntegracaoPV.sUM, tIntegracaoPV.dPrecoUnitario, tIntegracaoPV.dValorDescontoItem, tIntegracaoPV.dtDataEntrega, _
                tIntegracaoPV.iComErro, tIntegracaoPV.dtDataAtualizacao, tIntegracaoPV.iCodTabelaPrecoItem, tIntegracaoPV.sPedRepr, _
                TIPO_INTEGRACAO_IMPORTACAO, DATA_NULA, objIntegracaoPV.lNumIntArq, EMPRESA_TODA, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 190049
    
    IntegracaoPV_Le_ImpSemAtu = SUCESSO

    Exit Function

Erro_IntegracaoPV_Le_ImpSemAtu:

    IntegracaoPV_Le_ImpSemAtu = gErr

    Select Case gErr
    
        Case 190048

        Case 190049
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IntegracaoPV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190050)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function IntegracaoCliente_Le_Proximo(ByVal lComando As Long, tIntegracaoCliente As typeIntegracaoCliente, objIntegracaoCliente As ClassIntegracaoCliente) As Long

Dim lErro As Long

On Error GoTo Erro_IntegracaoCliente_Le_Proximo

    Set objIntegracaoCliente = New ClassIntegracaoCliente

    'Busca Primeiro
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189730

    If lErro <> AD_SQL_SUCESSO Then gError ERRO_LEITURA_SEM_DADOS

    objIntegracaoCliente.lNumIntDoc = tIntegracaoCliente.lNumIntDoc
    objIntegracaoCliente.lNumIntGer = tIntegracaoCliente.lNumIntGer
    objIntegracaoCliente.lNumIntArq = tIntegracaoCliente.lNumIntArq
    objIntegracaoCliente.lSeqRegistro = tIntegracaoCliente.lSeqRegistro
    objIntegracaoCliente.iTipoInt = tIntegracaoCliente.iTipoInt
    objIntegracaoCliente.lCodCliente = tIntegracaoCliente.lCodCliente
    objIntegracaoCliente.iCodFilial = tIntegracaoCliente.iCodFilial
    objIntegracaoCliente.sRazaoSocial = left(tIntegracaoCliente.sRazaoSocial, STRING_CLIENTE_RAZAO_SOCIAL)
    objIntegracaoCliente.sNomeReduzido = left(tIntegracaoCliente.sNomeReduzido, STRING_CLIENTE_NOME_REDUZIDO)
    objIntegracaoCliente.iTipo = tIntegracaoCliente.iTipo
    objIntegracaoCliente.sObservacao = left(tIntegracaoCliente.sObservacao, STRING_CLIENTE_OBSERVACAO)
    objIntegracaoCliente.dLimiteCredito = tIntegracaoCliente.dLimiteCredito
    objIntegracaoCliente.iCondicaoPagto = tIntegracaoCliente.iCondicaoPagto
    objIntegracaoCliente.iAtivo = tIntegracaoCliente.iAtivo
    objIntegracaoCliente.sFilialNome = tIntegracaoCliente.sFilialNome
    objIntegracaoCliente.sCgc = tIntegracaoCliente.sCgc
    objIntegracaoCliente.sRG = tIntegracaoCliente.sRG
    objIntegracaoCliente.sInscricaoEstadual = tIntegracaoCliente.sInscricaoEstadual
    objIntegracaoCliente.sInscricaoMunicipal = tIntegracaoCliente.sInscricaoMunicipal
    objIntegracaoCliente.sEndereco = left(tIntegracaoCliente.sEndereco, STRING_ENDERECO)
    objIntegracaoCliente.sBairro = left(tIntegracaoCliente.sBairro, STRING_BAIRRO)
    objIntegracaoCliente.sCidade = left(tIntegracaoCliente.sCidade, STRING_CIDADE)
    objIntegracaoCliente.sSiglaEstado = tIntegracaoCliente.sSiglaEstado
    objIntegracaoCliente.iCodigoPais = tIntegracaoCliente.iCodigoPais
    objIntegracaoCliente.sCEP = tIntegracaoCliente.sCEP
    objIntegracaoCliente.sTelefone1 = left(tIntegracaoCliente.sTelefone1, STRING_TELEFONE)
    objIntegracaoCliente.sTelefone2 = left(tIntegracaoCliente.sTelefone2, STRING_TELEFONE)
    objIntegracaoCliente.sEmail = left(tIntegracaoCliente.sEmail, STRING_EMAIL)
    objIntegracaoCliente.sFax = left(tIntegracaoCliente.sFax, STRING_FAX)
    objIntegracaoCliente.sContato = left(tIntegracaoCliente.sContato, STRING_CONTATO)
    objIntegracaoCliente.sEnderecoEnt = left(tIntegracaoCliente.sEnderecoEnt, STRING_ENDERECO)
    objIntegracaoCliente.sBairroEnt = left(tIntegracaoCliente.sBairroEnt, STRING_BAIRRO)
    objIntegracaoCliente.sCidadeEnt = left(tIntegracaoCliente.sCidadeEnt, STRING_CIDADE)
    objIntegracaoCliente.sSiglaEstadoEnt = tIntegracaoCliente.sSiglaEstadoEnt
    objIntegracaoCliente.iCodigoPaisEnt = tIntegracaoCliente.iCodigoPaisEnt
    objIntegracaoCliente.sCEPEnt = tIntegracaoCliente.sCEPEnt
    objIntegracaoCliente.sTelefone1Ent = left(tIntegracaoCliente.sTelefone1Ent, STRING_TELEFONE)
    objIntegracaoCliente.sTelefone2Ent = left(tIntegracaoCliente.sTelefone2Ent, STRING_TELEFONE)
    objIntegracaoCliente.sEmailEnt = left(tIntegracaoCliente.sEmailEnt, STRING_EMAIL)
    objIntegracaoCliente.sFaxEnt = left(tIntegracaoCliente.sFaxEnt, STRING_FAX)
    objIntegracaoCliente.sContatoEnt = left(tIntegracaoCliente.sContatoEnt, STRING_CONTATO)
    objIntegracaoCliente.sEnderecoCobr = left(tIntegracaoCliente.sEnderecoCobr, STRING_ENDERECO)
    objIntegracaoCliente.sBairroCobr = left(tIntegracaoCliente.sBairroCobr, STRING_BAIRRO)
    objIntegracaoCliente.sCidadeCobr = left(tIntegracaoCliente.sCidadeCobr, STRING_CIDADE)
    objIntegracaoCliente.sSiglaEstadoCobr = tIntegracaoCliente.sSiglaEstadoCobr
    objIntegracaoCliente.iCodigoPaisCobr = tIntegracaoCliente.iCodigoPaisCobr
    objIntegracaoCliente.sCEPCobr = tIntegracaoCliente.sCEPCobr
    objIntegracaoCliente.sTelefone1Cobr = left(tIntegracaoCliente.sTelefone1Cobr, STRING_TELEFONE)
    objIntegracaoCliente.sTelefone2Cobr = left(tIntegracaoCliente.sTelefone2Cobr, STRING_TELEFONE)
    objIntegracaoCliente.sEmailCobr = left(tIntegracaoCliente.sEmailCobr, STRING_EMAIL)
    objIntegracaoCliente.sFaxCobr = left(tIntegracaoCliente.sFaxCobr, STRING_FAX)
    objIntegracaoCliente.sContatoCobr = left(tIntegracaoCliente.sContatoCobr, STRING_CONTATO)
    objIntegracaoCliente.iComErro = tIntegracaoCliente.iComErro
    objIntegracaoCliente.dtDataAtualizacao = tIntegracaoCliente.dtDataAtualizacao
    objIntegracaoCliente.iVendedor = tIntegracaoCliente.iVendedor
    objIntegracaoCliente.sObservacaoFilial = left(tIntegracaoCliente.sObservacaoFilial, STRING_CLIENTE_OBSERVACAO)
    objIntegracaoCliente.lCodExterno = tIntegracaoCliente.lCodExterno
    objIntegracaoCliente.sTipoCliente = tIntegracaoCliente.sTipoCliente
    objIntegracaoCliente.sLogradouro = left(tIntegracaoCliente.sLogradouro, STRING_ENDERECO_LOGRADOURO)
    objIntegracaoCliente.sComplemento = left(tIntegracaoCliente.sComplemento, STRING_ENDERECO_COMPLEMENTO)
    objIntegracaoCliente.sTipoLogradouro = left(tIntegracaoCliente.sTipoLogradouro, STRING_ENDERECO_TIPOLOGRADOURO)
    objIntegracaoCliente.sEmail2 = tIntegracaoCliente.sEmail2
    objIntegracaoCliente.lNumero = tIntegracaoCliente.lNumero
    objIntegracaoCliente.iTelDDD1 = tIntegracaoCliente.iTelDDD1
    objIntegracaoCliente.iTelDDD2 = tIntegracaoCliente.iTelDDD2
    objIntegracaoCliente.iFaxDDD = tIntegracaoCliente.iFaxDDD
    objIntegracaoCliente.sTelNumero1 = left(tIntegracaoCliente.sTelNumero1, STRING_TELEFONE)
    objIntegracaoCliente.sTelNumero2 = left(tIntegracaoCliente.sTelNumero2, STRING_TELEFONE)
    objIntegracaoCliente.sFaxNumero = left(tIntegracaoCliente.sFaxNumero, STRING_FAX)
    objIntegracaoCliente.sLogradouroEnt = left(tIntegracaoCliente.sLogradouroEnt, STRING_ENDERECO_LOGRADOURO)
    objIntegracaoCliente.sComplementoEnt = left(tIntegracaoCliente.sComplementoEnt, STRING_ENDERECO_COMPLEMENTO)
    objIntegracaoCliente.sTipoLogradouroEnt = left(tIntegracaoCliente.sTipoLogradouroEnt, STRING_ENDERECO_TIPOLOGRADOURO)
    objIntegracaoCliente.sEmail2Ent = tIntegracaoCliente.sEmail2Ent
    objIntegracaoCliente.lNumeroEnt = tIntegracaoCliente.lNumeroEnt
    objIntegracaoCliente.iTelDDD1Ent = tIntegracaoCliente.iTelDDD1Ent
    objIntegracaoCliente.iTelDDD2Ent = tIntegracaoCliente.iTelDDD2Ent
    objIntegracaoCliente.iFaxDDDEnt = tIntegracaoCliente.iFaxDDDEnt
    objIntegracaoCliente.sTelNumero1Ent = left(tIntegracaoCliente.sTelNumero1Ent, STRING_TELEFONE)
    objIntegracaoCliente.sTelNumero2Ent = left(tIntegracaoCliente.sTelNumero2Ent, STRING_TELEFONE)
    objIntegracaoCliente.sFaxNumeroEnt = left(tIntegracaoCliente.sFaxNumeroEnt, STRING_FAX)
    objIntegracaoCliente.iTabelaPreco = tIntegracaoCliente.iTabelaPreco
    objIntegracaoCliente.sReferencia = tIntegracaoCliente.sReferencia
    objIntegracaoCliente.sReferenciaEnt = tIntegracaoCliente.sReferenciaEnt
    
    IntegracaoCliente_Le_Proximo = SUCESSO

    Exit Function

Erro_IntegracaoCliente_Le_Proximo:

    IntegracaoCliente_Le_Proximo = gErr

    Select Case gErr
        
        Case 189730
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INTEGRACAOCLIENTE", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189739)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function IntegracaoPV_Le_Proximo(ByVal lComando As Long, tIntegracaoPV As typeIntegracaoPV, objIntegracaoPV As ClassIntegracaoPV) As Long

Dim lErro As Long

On Error GoTo Erro_IntegracaoPV_Le_Proximo

    Set objIntegracaoPV = New ClassIntegracaoPV

    'Busca Primeiro
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189052

    If lErro <> AD_SQL_SUCESSO Then gError ERRO_LEITURA_SEM_DADOS

    objIntegracaoPV.lNumIntDoc = tIntegracaoPV.lNumIntDoc
    objIntegracaoPV.lNumIntGer = tIntegracaoPV.lNumIntGer
    objIntegracaoPV.lNumIntArq = tIntegracaoPV.lNumIntArq
    objIntegracaoPV.lSeqRegistro = tIntegracaoPV.lSeqRegistro
    objIntegracaoPV.iTipoInt = tIntegracaoPV.iTipoInt
    objIntegracaoPV.lCodPedidoExterno = tIntegracaoPV.lCodPedidoExterno
    objIntegracaoPV.dtDataEmissao = tIntegracaoPV.dtDataEmissao
    objIntegracaoPV.lCodClienteExterno = tIntegracaoPV.lCodClienteExterno
    objIntegracaoPV.lCodClienteCorporator = tIntegracaoPV.lCodClienteCorporator
    objIntegracaoPV.iCodFilialCorporator = tIntegracaoPV.iCodFilialCorporator
    objIntegracaoPV.sCGCCliente = tIntegracaoPV.sCGCCliente
    objIntegracaoPV.sNaturezaOp = tIntegracaoPV.sNaturezaOp
    objIntegracaoPV.iCodTabelaPreco = tIntegracaoPV.iCodTabelaPreco
    objIntegracaoPV.iCodCondPagto = tIntegracaoPV.iCodCondPagto
    objIntegracaoPV.sNomeCondPagtoExterno = tIntegracaoPV.sNomeCondPagtoExterno
    objIntegracaoPV.iCodCondPagtoExterno = tIntegracaoPV.iCodCondPagtoExterno
    objIntegracaoPV.iFilialEmpresaFat = tIntegracaoPV.iFilialEmpresaFat
    objIntegracaoPV.dValorDescontoPedido = tIntegracaoPV.dValorDescontoPedido
    objIntegracaoPV.iFrete = tIntegracaoPV.iFrete
    objIntegracaoPV.dValorFretePedido = tIntegracaoPV.dValorFretePedido
    objIntegracaoPV.dValorSeguroPedido = tIntegracaoPV.dValorSeguroPedido
    objIntegracaoPV.dValorOutrasDespesasPedido = tIntegracaoPV.dValorOutrasDespesasPedido
    objIntegracaoPV.iFilialEmpresaEnt = tIntegracaoPV.iFilialEmpresaEnt
    objIntegracaoPV.iTrazerTranspAuto = tIntegracaoPV.iTrazerTranspAuto
    objIntegracaoPV.iCodTransportadora = tIntegracaoPV.iCodTransportadora
    objIntegracaoPV.iCodTransportadoraExterno = tIntegracaoPV.iCodTransportadoraExterno
    objIntegracaoPV.sNomeTransportadoraExterno = tIntegracaoPV.sNomeTransportadoraExterno
    objIntegracaoPV.iTrazerMensagemAuto = tIntegracaoPV.iTrazerMensagemAuto
    objIntegracaoPV.sMensagemPedido = tIntegracaoPV.sMensagemPedido
    objIntegracaoPV.iTrazerPesoAuto = tIntegracaoPV.iTrazerPesoAuto
    objIntegracaoPV.dPesoBruto = tIntegracaoPV.dPesoBruto
    objIntegracaoPV.dPesoLiquido = tIntegracaoPV.dPesoLiquido
    objIntegracaoPV.iTrazerCanalVendaAuto = tIntegracaoPV.iTrazerCanalVendaAuto
    objIntegracaoPV.iCanalVenda = tIntegracaoPV.iCanalVenda
    objIntegracaoPV.iCodVendedor = tIntegracaoPV.iCodVendedor
    objIntegracaoPV.iCodVendedorExterno = tIntegracaoPV.iCodVendedorExterno
    objIntegracaoPV.sNomeVendedorExterno = tIntegracaoPV.sNomeVendedorExterno
    objIntegracaoPV.iTrazerComissaoAuto = tIntegracaoPV.iTrazerComissaoAuto
    objIntegracaoPV.dValorComissao = tIntegracaoPV.dValorComissao
    objIntegracaoPV.iTrazerReservaAuto = tIntegracaoPV.iTrazerReservaAuto
    objIntegracaoPV.iCodAlmoxarifado = tIntegracaoPV.iCodAlmoxarifado
    objIntegracaoPV.iCodAlmoxarifadoExterno = tIntegracaoPV.iCodAlmoxarifadoExterno
    objIntegracaoPV.sNomeAlmoxarifadoExterno = tIntegracaoPV.sNomeAlmoxarifadoExterno
    objIntegracaoPV.iItem = tIntegracaoPV.iItem
    objIntegracaoPV.sCodProduto = tIntegracaoPV.sCodProduto
    objIntegracaoPV.sCodProdutoExterno = tIntegracaoPV.sCodProdutoExterno
    objIntegracaoPV.iTrazerDescricaoAuto = tIntegracaoPV.iTrazerDescricaoAuto
    objIntegracaoPV.sDescricaoItem = tIntegracaoPV.sDescricaoItem
    objIntegracaoPV.dQuantidadePedida = tIntegracaoPV.dQuantidadePedida
    objIntegracaoPV.dQuantidadeCancelada = tIntegracaoPV.dQuantidadeCancelada
    objIntegracaoPV.sUM = tIntegracaoPV.sUM
    objIntegracaoPV.dPrecoUnitario = tIntegracaoPV.dPrecoUnitario
    objIntegracaoPV.dValorDescontoItem = tIntegracaoPV.dValorDescontoItem
    objIntegracaoPV.dtDataEntrega = tIntegracaoPV.dtDataEntrega
    objIntegracaoPV.iComErro = tIntegracaoPV.iComErro
    objIntegracaoPV.dtDataAtualizacao = tIntegracaoPV.dtDataAtualizacao
    objIntegracaoPV.iCodTabelaPrecoItem = tIntegracaoPV.iCodTabelaPrecoItem
    objIntegracaoPV.sPedRepr = tIntegracaoPV.sPedRepr

    IntegracaoPV_Le_Proximo = SUCESSO

    Exit Function

Erro_IntegracaoPV_Le_Proximo:

    IntegracaoPV_Le_Proximo = gErr

    Select Case gErr
        
        Case 189052
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INTEGRACAOCLIENTE", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189053)

    End Select

    Exit Function

End Function

Private Function IntegracaoSldProd_Le_Proximo(ByVal lComando As Long, tIntegracaoSldProd As typeIntegracaoSldProd, objIntegracaoSldProd As ClassIntegracaoSldProd) As Long

Dim lErro As Long

On Error GoTo Erro_IntegracaoSldProd_Le_Proximo

    Set objIntegracaoSldProd = New ClassIntegracaoSldProd

    'Busca Primeiro
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190137

    If lErro <> AD_SQL_SUCESSO Then gError ERRO_LEITURA_SEM_DADOS

    objIntegracaoSldProd.lNumIntDoc = tIntegracaoSldProd.lNumIntDoc
    objIntegracaoSldProd.lNumIntGer = tIntegracaoSldProd.lNumIntGer
    objIntegracaoSldProd.lNumIntArq = tIntegracaoSldProd.lNumIntArq
    objIntegracaoSldProd.lSeqRegistro = tIntegracaoSldProd.lSeqRegistro
    objIntegracaoSldProd.iTipoInt = tIntegracaoSldProd.iTipoInt
    objIntegracaoSldProd.sCodProduto = tIntegracaoSldProd.sCodProduto
    objIntegracaoSldProd.iAlmoxarifado = tIntegracaoSldProd.iAlmoxarifado
    objIntegracaoSldProd.dSaldoDisp = tIntegracaoSldProd.dSaldoDisp
    objIntegracaoSldProd.iComErro = tIntegracaoSldProd.iComErro
    objIntegracaoSldProd.dtDataAtualizacao = tIntegracaoSldProd.dtDataAtualizacao

    IntegracaoSldProd_Le_Proximo = SUCESSO

    Exit Function

Erro_IntegracaoSldProd_Le_Proximo:

    IntegracaoSldProd_Le_Proximo = gErr

    Select Case gErr
        
        Case 190137
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INTEGRACAOSLDPROD", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190139)

    End Select

    Exit Function

End Function


'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function ArqImportacao_Le(ByVal lComando As Long, ByVal objArqImportacao As ClassArqImportacao) As Long

Dim lErro As Long
Dim tArqImportacao As typeArqImportacao

On Error GoTo Erro_ArqImportacao_Le

    'Alocação de espaço no buffer
    tArqImportacao.sNomeArquivo = String(STRING_NOME_ARQUIVO, 0)
    tArqImportacao.sUsuario = String(STRING_USUARIO, 0)

    'Le a tabelaArqImportacao
    lErro = Comando_Executar(lComando, "SELECT NumIntArq, TipoArq, NomeArquivo, DataImportacao, HoraImportacao, DataAtualizacao, " & _
                "HoraAtualizacao, Usuario FROM ArqImportacao WHERE NomeArquivo= ? ", _
                tArqImportacao.lNumIntArq, tArqImportacao.iTipoArq, tArqImportacao.sNomeArquivo, tArqImportacao.dtDataImportacao, _
                tArqImportacao.dHoraImportacao, tArqImportacao.dtDataAtualizacao, tArqImportacao.dHoraAtualizacao, tArqImportacao.sUsuario, _
                objArqImportacao.sNomeArquivo)
    If lErro <> AD_SQL_SUCESSO Then gError 189724

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189725

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objArqImportacao.lNumIntArq = tArqImportacao.lNumIntArq
    objArqImportacao.iTipoArq = tArqImportacao.iTipoArq
    objArqImportacao.sNomeArquivo = tArqImportacao.sNomeArquivo
    objArqImportacao.dtDataImportacao = tArqImportacao.dtDataImportacao
    objArqImportacao.dHoraImportacao = tArqImportacao.dHoraImportacao
    objArqImportacao.dtDataAtualizacao = tArqImportacao.dtDataAtualizacao
    objArqImportacao.dHoraAtualizacao = tArqImportacao.dHoraAtualizacao
    objArqImportacao.sUsuario = tArqImportacao.sUsuario

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ArqImportacao_Le = SUCESSO

    Exit Function

Erro_ArqImportacao_Le:

    ArqImportacao_Le = gErr

    Select Case gErr

        Case 189724, 189725
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQIMPORTACAO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189726)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function ArqExportacao_Le(ByVal lComando As Long, ByVal objArqExportacao As ClassArqExportacao) As Long

Dim lErro As Long
Dim tArqExportacao As typeArqExportacao

On Error GoTo Erro_ArqExportacao_Le

    'Alocação de espaço no buffer
    tArqExportacao.sNomeArquivo = String(STRING_NOME_ARQUIVO, 0)
    tArqExportacao.sUsuario = String(STRING_USUARIO, 0)

    'Le a tabelaArqExportacao
    lErro = Comando_Executar(lComando, "SELECT NumIntArq, NomeArquivo, TipoArq, DataExportacao, HoraExportacao, Usuario FROM ArqExportacao WHERE NomeArquivo= ? ", _
                tArqExportacao.lNumIntArq, tArqExportacao.sNomeArquivo, tArqExportacao.iTipoArq, tArqExportacao.dtDataExportacao, _
                tArqExportacao.dHoraExportacao, tArqExportacao.sUsuario, _
                objArqExportacao.sNomeArquivo)
    If lErro <> AD_SQL_SUCESSO Then gError 189702

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189703

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objArqExportacao.lNumIntArq = tArqExportacao.lNumIntArq
    objArqExportacao.sNomeArquivo = tArqExportacao.sNomeArquivo
    objArqExportacao.iTipoArq = tArqExportacao.iTipoArq
    objArqExportacao.dtDataExportacao = tArqExportacao.dtDataExportacao
    objArqExportacao.dHoraExportacao = tArqExportacao.dHoraExportacao
    objArqExportacao.sUsuario = tArqExportacao.sUsuario

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ArqExportacao_Le = SUCESSO

    Exit Function

Erro_ArqExportacao_Le:

    ArqExportacao_Le = gErr

    Select Case gErr

        Case 189702, 189703
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQEXPORTACAO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189704)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function TipoArqIntegracao_Le(ByVal objTipoArqIntegracao As ClassTipoArqIntegracao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTipoArqIntegracao As typeTipoArqIntegracao

On Error GoTo Erro_TipoArqIntegracao_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 189697

    'Alocação de espaço no buffer
    tTipoArqIntegracao.sDescricao = String(STRING_DESCRICAO, 0)
    tTipoArqIntegracao.sSiglaArq = String(STRING_SIGLA_ARQ, 0)
    tTipoArqIntegracao.sTabela = String(STRING_TABELA, 0)

    'Le a tabelaTipoArqIntegracao
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, SiglaArq, Tabela FROM TipoArqIntegracao WHERE Codigo= ? ", _
                tTipoArqIntegracao.iCodigo, tTipoArqIntegracao.sDescricao, tTipoArqIntegracao.sSiglaArq, tTipoArqIntegracao.sTabela, _
                objTipoArqIntegracao.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 189698

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189699

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTipoArqIntegracao.iCodigo = tTipoArqIntegracao.iCodigo
    objTipoArqIntegracao.sDescricao = tTipoArqIntegracao.sDescricao
    objTipoArqIntegracao.sSiglaArq = tTipoArqIntegracao.sSiglaArq
    objTipoArqIntegracao.sTabela = tTipoArqIntegracao.sTabela

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TipoArqIntegracao_Le = SUCESSO

    Exit Function

Erro_TipoArqIntegracao_Le:

    TipoArqIntegracao_Le = gErr

    Select Case gErr

        Case 189697
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189698, 189699
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOARQINTEGRACAO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189701)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function IntegracaoCliente_Insere(ByVal lComando As Long, ByVal objIntegracaoCliente As ClassIntegracaoCliente) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_IntegracaoCliente_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_INTEGRACAOCLIENTE", lNumIntDoc)
    If lErro <> SUCESSO Then gError 189743

    objIntegracaoCliente.lNumIntDoc = lNumIntDoc

    With objIntegracaoCliente
        lErro = Comando_Executar(lComando, "INSERT INTO IntegracaoCliente( NumIntDoc, NumIntGer, NumIntArq, SeqRegistro, TipoInt, CodCliente, CodFilial, RazaoSocial, NomeReduzido, " & _
                "Tipo, Observacao, LimiteCredito, CondicaoPagto, Ativo, FilialNome, CGC, RG, InscricaoEstadual, InscricaoMunicipal, " & _
                "Endereco, Bairro, Cidade, SiglaEstado, CodigoPais, CEP, Telefone1, Telefone2, Email, Fax, " & _
                "Contato, EnderecoEnt, BairroEnt, CidadeEnt, SiglaEstadoEnt, CodigoPaisEnt, CEPEnt, Telefone1Ent, Telefone2Ent, EmailEnt, " & _
                "FaxEnt, ContatoEnt, EnderecoCobr, BairroCobr, CidadeCobr, SiglaEstadoCobr, CodigoPaisCobr, CEPCobr, Telefone1Cobr, Telefone2Cobr, " & _
                "EmailCobr, FaxCobr, ContatoCobr, ComErro, DataAtualizacao, Vendedor, ObservacaoFilial, CodExterno, TipoCliente, " & _
                "Logradouro,Complemento,TipoLogradouro,Email2,Numero,TelDDD1,TelDDD2,FaxDDD,TelNumero1,TelNumero2,FaxNumero,LogradouroEnt, " & _
                "ComplementoEnt,TipoLogradouroEnt,Email2Ent,NumeroEnt,TelDDD1Ent,TelDDD2Ent,FaxDDDEnt,TelNumero1Ent,TelNumero2Ent,FaxNumeroEnt,TabelaPreco) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                .lNumIntDoc, .lNumIntGer, .lNumIntArq, .lSeqRegistro, .iTipoInt, .lCodCliente, .iCodFilial, .sRazaoSocial, .sNomeReduzido, _
                .iTipo, .sObservacao, .dLimiteCredito, .iCondicaoPagto, .iAtivo, .sFilialNome, .sCgc, .sRG, .sInscricaoEstadual, .sInscricaoMunicipal, _
                .sEndereco, .sBairro, .sCidade, .sSiglaEstado, .iCodigoPais, .sCEP, .sTelefone1, .sTelefone2, .sEmail, .sFax, _
                .sContato, .sEnderecoEnt, .sBairroEnt, .sCidadeEnt, .sSiglaEstadoEnt, .iCodigoPaisEnt, .sCEPEnt, .sTelefone1Ent, .sTelefone2Ent, .sEmailEnt, _
                .sFaxEnt, .sContatoEnt, .sEnderecoCobr, .sBairroCobr, .sCidadeCobr, .sSiglaEstadoCobr, .iCodigoPaisCobr, .sCEPCobr, .sTelefone1Cobr, .sTelefone2Cobr, _
                .sEmailCobr, .sFaxCobr, .sContatoCobr, 0, DATA_NULA, .iVendedor, .sObservacaoFilial, .lCodExterno, .sTipoCliente, _
                .sLogradouro, .sComplemento, .sTipoLogradouro, .sEmail2, .lNumero, .iTelDDD1, .iTelDDD2, .iFaxDDD, .sTelNumero1, .sTelNumero2, .sFaxNumero, .sLogradouroEnt, _
                .sComplementoEnt, .sTipoLogradouroEnt, .sEmail2Ent, .lNumeroEnt, .iTelDDD1Ent, .iTelDDD2Ent, .iFaxDDDEnt, .sTelNumero1Ent, .sTelNumero2Ent, .sFaxNumeroEnt, .iTabelaPreco)
    End With
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189744

    IntegracaoCliente_Insere = SUCESSO

    Exit Function

Erro_IntegracaoCliente_Insere:

    IntegracaoCliente_Insere = gErr

    Select Case gErr

        Case 189743

        Case 189744
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INTEGRACAOCLIENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189745)

    End Select

    Exit Function

End Function

Private Function IntegracaoPV_Insere(ByVal lComando As Long, ByVal objIntegracaoPV As ClassIntegracaoPV) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_IntegracaoPV_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_INTEGRACAOPV", lNumIntDoc)
    If lErro <> SUCESSO Then gError 190045

    objIntegracaoPV.lNumIntDoc = lNumIntDoc

    lErro = Comando_Executar(lComando, "INSERT INTO IntegracaoPV( NumIntDoc, NumIntGer, NumIntArq, SeqRegistro, " & _
                "TipoInt, CodPedidoExterno, DataEmissao, CodClienteExterno, CodClienteCorporator, " & _
                "CodFilialCorporator, CGCCliente, NaturezaOP, CodTabelaPreco, CodCondPagto, " & _
                "NomeCondPagtoExterno, CodCondPagtoExterno, FilialEmpresaFat, ValorDescontoPedido, Frete, " & _
                "ValorFretePedido, ValorSeguroPedido, ValorOutrasDespesasPedido, FilialEmpresaEnt, TrazerTranspAuto, " & _
                "CodTransportadora, CodTransportadoraExterno, NomeTransportadoraExterno, TrazerMensagemAuto, MensagemPedido, " & _
                "TrazerPesoAuto, PesoBruto, PesoLiquido, TrazerCanalVendaAuto, CanalVenda, " & _
                "CodVendedor, CodVendedorExterno, NomeVendedorExterno, TrazerComissaoAuto, ValorComissao, " & _
                "TrazerReservaAuto, CodAlmoxarifado, CodAlmoxarifadoExterno, NomeAlmoxarifadoExterno, Item, " & _
                "CodProduto, CodProdutoExterno, TrazerDescricaoAuto, DescricaoItem, QuantidadePedida, " & _
                "QuantidadeCancelada, UM, PrecoUnitario, ValorDescontoItem, DataEntrega, " & _
                "ComErro, DataAtualizacao,PedRepr, FilialEmpresa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objIntegracaoPV.lNumIntDoc, objIntegracaoPV.lNumIntGer, objIntegracaoPV.lNumIntArq, objIntegracaoPV.lSeqRegistro, _
                objIntegracaoPV.iTipoInt, objIntegracaoPV.lCodPedidoExterno, objIntegracaoPV.dtDataEmissao, objIntegracaoPV.lCodClienteExterno, objIntegracaoPV.lCodClienteCorporator, _
                objIntegracaoPV.iCodFilialCorporator, objIntegracaoPV.sCGCCliente, objIntegracaoPV.sNaturezaOp, objIntegracaoPV.iCodTabelaPreco, objIntegracaoPV.iCodCondPagto, _
                objIntegracaoPV.sNomeCondPagtoExterno, objIntegracaoPV.iCodCondPagtoExterno, objIntegracaoPV.iFilialEmpresaFat, objIntegracaoPV.dValorDescontoPedido, objIntegracaoPV.iFrete, _
                objIntegracaoPV.dValorFretePedido, objIntegracaoPV.dValorSeguroPedido, objIntegracaoPV.dValorOutrasDespesasPedido, objIntegracaoPV.iFilialEmpresaEnt, objIntegracaoPV.iTrazerTranspAuto, _
                objIntegracaoPV.iCodTransportadora, objIntegracaoPV.iCodTransportadoraExterno, objIntegracaoPV.sNomeTransportadoraExterno, objIntegracaoPV.iTrazerMensagemAuto, objIntegracaoPV.sMensagemPedido, _
                objIntegracaoPV.iTrazerPesoAuto, objIntegracaoPV.dPesoBruto, objIntegracaoPV.dPesoLiquido, objIntegracaoPV.iTrazerCanalVendaAuto, objIntegracaoPV.iCanalVenda, _
                objIntegracaoPV.iCodVendedor, objIntegracaoPV.iCodVendedorExterno, objIntegracaoPV.sNomeVendedorExterno, objIntegracaoPV.iTrazerComissaoAuto, objIntegracaoPV.dValorComissao, _
                objIntegracaoPV.iTrazerReservaAuto, objIntegracaoPV.iCodAlmoxarifado, objIntegracaoPV.iCodAlmoxarifadoExterno, objIntegracaoPV.sNomeAlmoxarifadoExterno, objIntegracaoPV.iItem, _
                objIntegracaoPV.sCodProduto, objIntegracaoPV.sCodProdutoExterno, objIntegracaoPV.iTrazerDescricaoAuto, objIntegracaoPV.sDescricaoItem, objIntegracaoPV.dQuantidadePedida, _
                objIntegracaoPV.dQuantidadeCancelada, objIntegracaoPV.sUM, objIntegracaoPV.dPrecoUnitario, objIntegracaoPV.dValorDescontoItem, objIntegracaoPV.dtDataEntrega, _
                0, DATA_NULA, objIntegracaoPV.sPedRepr, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190046

     IntegracaoPV_Insere = SUCESSO

    Exit Function

Erro_IntegracaoPV_Insere:

    IntegracaoPV_Insere = gErr

    Select Case gErr

        Case 190045

        Case 190046
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INTEGRACAOPV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190047)

    End Select

    Exit Function

End Function

Private Function IntegracaoSldProd_Insere(ByVal lComando As Long, ByVal objIntegracaoSldProd As ClassIntegracaoSldProd) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_IntegracaoSldProd_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_INTEGRACAOSLDPROD", lNumIntDoc)
    If lErro <> SUCESSO Then gError 190121

    objIntegracaoSldProd.lNumIntDoc = lNumIntDoc

    lErro = Comando_Executar(lComando, "INSERT INTO IntegracaoSldProd( NumIntDoc, NumIntGer, NumIntArq, SeqRegistro, " & _
            "TipoInt, CodProduto,Almoxarifado, SaldoDisp, ComErro, DataAtualizacao) VALUES (?,?,?,?,?,?,?,?,?,?)", _
            objIntegracaoSldProd.lNumIntDoc, objIntegracaoSldProd.lNumIntGer, objIntegracaoSldProd.lNumIntArq, objIntegracaoSldProd.lSeqRegistro, _
            objIntegracaoSldProd.iTipoInt, objIntegracaoSldProd.sCodProduto, objIntegracaoSldProd.iAlmoxarifado, objIntegracaoSldProd.dSaldoDisp, 0, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190122

     IntegracaoSldProd_Insere = SUCESSO

    Exit Function

Erro_IntegracaoSldProd_Insere:

    IntegracaoSldProd_Insere = gErr

    Select Case gErr

        Case 190121

        Case 190122
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INTEGRACAOSLDPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190123)

    End Select

    Exit Function

End Function

Private Function IntegracaoCliente_Atualiza(ByVal lComando As Long, ByVal lComandoPos As Long, ByVal objIntegracaoCliente As ClassIntegracaoCliente) As Long

Dim lErro As Long

On Error GoTo Erro_IntegracaoCliente_Atualiza

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(lComandoPos)
    If lErro <> SUCESSO Then gError 189746

    lErro = Comando_ExecutarPos(lComando, "UPDATE IntegracaoCliente SET NumIntArq= ?, ComErro= ? , DataAtualizacao= ? ", lComandoPos, objIntegracaoCliente.lNumIntArq, objIntegracaoCliente.iComErro, objIntegracaoCliente.dtDataAtualizacao)
    If lErro <> AD_SQL_SUCESSO Then
        gError 189747
    End If

    IntegracaoCliente_Atualiza = SUCESSO

    Exit Function

Erro_IntegracaoCliente_Atualiza:

    IntegracaoCliente_Atualiza = gErr

    Select Case gErr

        Case 189746
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_INTEGRACAOCLIENTE", gErr)

        Case 189747
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_INTEGRACAOCLIENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189748)

    End Select

    Exit Function

End Function

Private Function IntegracaoSldProd_Atualiza(ByVal lComando As Long, ByVal lComandoPos As Long, ByVal objIntegracaoSldProd As ClassIntegracaoSldProd) As Long

Dim lErro As Long

On Error GoTo Erro_IntegracaoSldProd_Atualiza

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(lComandoPos)
    If lErro <> SUCESSO Then gError 190140

    lErro = Comando_ExecutarPos(lComando, "UPDATE IntegracaoSldProd SET NumIntArq= ?, ComErro= ? , DataAtualizacao= ? ", lComandoPos, objIntegracaoSldProd.lNumIntArq, objIntegracaoSldProd.iComErro, objIntegracaoSldProd.dtDataAtualizacao)
    If lErro <> AD_SQL_SUCESSO Then
        gError 190141
    End If

    IntegracaoSldProd_Atualiza = SUCESSO

    Exit Function

Erro_IntegracaoSldProd_Atualiza:

    IntegracaoSldProd_Atualiza = gErr

    Select Case gErr

        Case 190140
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_INTEGRACAOSLDPROD", gErr)

        Case 190141
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_INTEGRACAOSLDPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190142)

    End Select

    Exit Function

End Function

Private Function ArqImportacao_Insere(ByVal lComando As Long, ByVal objArqImportacao As ClassArqImportacao) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ArqImportacao_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_ARQINTEGRACAO", lNumIntDoc)
    If lErro <> SUCESSO Then gError 189733

    objArqImportacao.lNumIntArq = lNumIntDoc

    lErro = Comando_Executar(lComando, "INSERT INTO ArqImportacao( NumIntArq, TipoArq, NomeArquivo, DataImportacao, " & _
            "HoraImportacao, DataAtualizacao, HoraAtualizacao, Usuario) VALUES (?,?,?,?,?,?,?,?)", _
            objArqImportacao.lNumIntArq, objArqImportacao.iTipoArq, objArqImportacao.sNomeArquivo, Date, _
            CDbl(Time), DATA_NULA, 0, gsUsuario)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189734

    ArqImportacao_Insere = SUCESSO

    Exit Function

Erro_ArqImportacao_Insere:

    ArqImportacao_Insere = gErr

    Select Case gErr
    
        Case 189733

        Case 189734
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ARQIMPORTACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189735)

    End Select

    Exit Function

End Function

Private Function ArqImportacao_Atualiza(ByVal lComando As Long, ByVal lComandoPos As Long) As Long

Dim lErro As Long

On Error GoTo Erro_ArqImportacao_Atualiza

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(lComandoPos)
    If lErro <> SUCESSO Then gError 189727

    lErro = Comando_ExecutarPos(lComando, "UPDATE ArqImportacao SET  DataAtualizacao= ? , HoraAtualizacao= ? ", lComandoPos, Date, CDbl(Time))
    If lErro <> AD_SQL_SUCESSO Then gError 189728

    ArqImportacao_Atualiza = SUCESSO

    Exit Function

Erro_ArqImportacao_Atualiza:

    ArqImportacao_Atualiza = gErr

    Select Case gErr

        Case 189727
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ARQIMPORTACAO", gErr)

        Case 189728
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ARQIMPORTACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189729)

    End Select

    Exit Function

End Function

Private Function ArqExportacaoAux_Insere(ByVal lComando As Long, ByVal objArqExportacaoAux As ClassArqExportacaoAux) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lTransacao As Long

On Error GoTo Erro_ArqExportacaoAux_Insere

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 189806
    
    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_ARQEXPORTACAOAUX", lNumIntDoc)
    If lErro <> SUCESSO Then gError 189733

    objArqExportacaoAux.lNumIntGeracao = lNumIntDoc
    
    lErro = Comando_Executar(lComando, "INSERT INTO ArqExportacaoAux( NumIntGeracao, DataGeracao, HoraGeracao, Usuario, " & _
            "Exportar, ExpDataDe, ExpDataAte) VALUES (?,?,?,?,?,?,?)", _
            objArqExportacaoAux.lNumIntGeracao, Date, CDbl(Time), gsUsuario, _
            objArqExportacaoAux.iExportar, objArqExportacaoAux.dtExpDataDe, objArqExportacaoAux.dtExpDataAte)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189721

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 189807
    
    ArqExportacaoAux_Insere = SUCESSO

    Exit Function

Erro_ArqExportacaoAux_Insere:

    ArqExportacaoAux_Insere = gErr

    Select Case gErr

        Case 189720
        
        Case 189721
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ARQEXPORTACAOAUX", gErr)

        Case 189806
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 189807
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189722)

    End Select
    
    Call Transacao_Rollback

    Exit Function

End Function

Private Function ArqExportacao_Insere(ByVal lComando As Long, ByVal objArqExportacao As ClassArqExportacao) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ArqExportacao_Insere

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_ARQINTEGRACAO", lNumIntDoc)
    If lErro <> SUCESSO Then gError 189711

    objArqExportacao.lNumIntArq = lNumIntDoc

    lErro = Comando_Executar(lComando, "INSERT INTO ArqExportacao( NumIntArq, NomeArquivo, TipoArq, DataExportacao, " & _
            "HoraExportacao, Usuario) VALUES (?,?,?,?,?,?)", _
            objArqExportacao.lNumIntArq, objArqExportacao.sNomeArquivo, objArqExportacao.iTipoArq, Date, _
            CDbl(Time), gsUsuario)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189712

    ArqExportacao_Insere = SUCESSO

    Exit Function

Erro_ArqExportacao_Insere:

    ArqExportacao_Insere = gErr

    Select Case gErr

        Case 189711
        
        Case 189712
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ARQEXPORTACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189713)

    End Select

    Exit Function

End Function

Public Function Importa_Dados(ByVal objImportacaoAux As ClassArqImportacaoAux) As Long
        
Dim lErro As Long
Dim iIndice As Integer
Dim alComandoExt(1 To 6) As Long
Dim alComando(1 To 28) As Long
Dim alComandoAd(1 To 31) As Long
Dim objArqImpArq As ClassArqImportacaoArq
Dim objTipoArq As ClassTipoArqIntegracao
Dim colTipoArq As New Collection
Dim bJaImp As Boolean
Dim sDiretorio As String
Dim lRetorno As Long
Dim objFSO As New FileSystemObject
Dim objFile As File, colAux As New Collection

On Error GoTo Erro_Importa_Dados

    'MsgBox ("1")

    lErro = Abre_Conexoes
    If lErro <> SUCESSO Then gError 189912
    
    'abertura de comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
        If alComandoExt(iIndice) = 0 Then gError 189913
    Next
    
    'abertura de comando
    For iIndice = LBound(alComandoAd) To UBound(alComandoAd)
        alComandoAd(iIndice) = Comando_Abrir()
        If alComandoAd(iIndice) = 0 Then gError 189914
    Next
    
    'abertura de comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 189914
    Next
    
    'Preenche as tabelas de IntegracaoPV e IntegracaoCliente com outros dados que não dos arquivos padrão
    lErro = Importa_Dados_Customizado(objImportacaoAux)
    If lErro <> SUCESSO Then gError 189915
    
    giValidacaoManual = objImportacaoAux.iValidacaoManual
    
    'MsgBox ("3")
    
    'Obtém o diretório onde estão os arquivos
    sDiretorio = String(512, 0)
    lRetorno = GetPrivateProfileString("Geral", "dirArqImportTemp", "c:\", sDiretorio, 512, "ADM100.INI")
    sDiretorio = left(sDiretorio, lRetorno)
    
    gsDiretorioImportTemp = sDiretorio
    
    'MsgBox ("4")
    
    lErro = CF("TipoArqIntegracao_Le_Todos", colTipoArq, TIPO_INTEGRACAO_IMPORTACAO)
    If lErro <> SUCESSO Then gError 189915
    
    'MsgBox ("4.1")
    
    gsDiretorioImport = objImportacaoAux.sDiretorio
    
    'MsgBox ("4.2")
    'MsgBox ("4.2a" & objImportacaoAux.sDiretorio)
    
    lErro = ImportaArq_Inicia(objImportacaoAux.colArquivos.Count)
    If lErro <> SUCESSO Then gError 189916
    
    'MsgBox ("5")
    
    'Importa os clientes antes dos pedidos
    For Each objArqImpArq In objImportacaoAux.colArquivos
        If objArqImpArq.objTipoArq.iCodigo = TIPO_ARQ_CLIENTE Then
            colAux.Add objArqImpArq
        End If
    Next
    For Each objArqImpArq In objImportacaoAux.colArquivos
        If objArqImpArq.objTipoArq.iCodigo = TIPO_ARQ_PV Then
            colAux.Add objArqImpArq
        End If
    Next
    
    'Gera as informações da tabela temporária
    For Each objArqImpArq In colAux
    
        lErro = CF("ArqImport_Verifica_JaImportado", objArqImpArq.sNomeArquivo, bJaImp)
        If lErro <> SUCESSO Then gError 189962
        
        TelaAcompanhaBatch.NomeArqImp = objArqImpArq.sNomeArquivo
        
        DoEvents
        
        If Not bJaImp Then
        
            'MsgBox ("6")
            
            'Copia o arquivo
            FileCopy objArqImpArq.objFile.Path, gsDiretorioImportTemp & objArqImpArq.sNomeArquivo
        
            Set objFile = objFSO.GetFile(gsDiretorioImportTemp & objArqImpArq.sNomeArquivo)
                        
            Set objArqImpArq.objFile = objFile
        
            Select Case objArqImpArq.objTipoArq.iCodigo
            
                Case TIPO_ARQ_CLIENTE
                
                    lErro = Importa_Dados_Cliente_Arq(alComando(1), alComando(2), objArqImpArq)
                    If lErro <> SUCESSO Then gError 189917
                
                Case TIPO_ARQ_PV
                
                    If UCase(right(objArqImpArq.sNomeArquivo, 4)) = ".XLS" Then
                    'Precisa dois dados do cliente já nas tabelas oficiais para poder validar
                    
                        For Each objTipoArq In colTipoArq
                            If objTipoArq.iCodigo = TIPO_ARQ_CLIENTE Then
                                lErro = Importacao_Inicia(alComando(8), alComando(9), objTipoArq)
                                If lErro <> SUCESSO Then gError 189972
                        
                                lErro = Importa_Dados_Cliente(alComandoExt(1), alComandoExt(2), alComandoExt(3), alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), alComando(26), alComando(27), objTipoArq)
                                If lErro <> SUCESSO Then gError 189920
                            End If
                        Next
                    
                    End If

                
                    lErro = Importa_Dados_PV_Arq(alComando(10), alComando(11), objArqImpArq)
                    If lErro <> SUCESSO Then gError 190035
                
                Case Else
                    gError 189918
                
            End Select
            
        Else
            TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "O arquivo " & objArqImpArq.sNomeArquivo & " já foi importado." & vbNewLine
        End If
               
        'MsgBox ("7")
        
        'Atualiza a barra de progressão
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 189919
    
    Next
    
    'MsgBox ("8")
    
    'Gera as informações da tabela temporária
    For Each objTipoArq In colTipoArq
    
        lErro = Importacao_Inicia(alComando(8), alComando(9), objTipoArq)
        If lErro <> SUCESSO Then gError 189972
    
        Select Case objTipoArq.iCodigo
        
            Case TIPO_ARQ_CLIENTE
            
                lErro = Importa_Dados_Cliente(alComandoExt(1), alComandoExt(2), alComandoExt(3), alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), alComando(26), alComando(27), objTipoArq)
                If lErro <> SUCESSO Then gError 189920
            
            Case TIPO_ARQ_PV
            
                lErro = Importa_Dados_PV(alComandoAd, alComandoExt(4), alComandoExt(5), alComandoExt(6), alComando(12), alComando(13), alComando(14), alComando(15), alComando(16), alComando(17), alComando(18), alComando(19), alComando(20), alComando(21), alComando(22), alComando(23), alComando(24), alComando(25), alComando(28), objTipoArq)
                If lErro <> SUCESSO Then gError 190036
            
            Case Else
                gError 189921
            
        End Select
    
    Next
        
    'MsgBox ("9")
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComandoAd) To UBound(alComandoAd)
        Call Comando_Fechar(alComandoAd(iIndice))
    Next
        
    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    Call Fecha_Conexoes

    lErro = CF("Config_Grava", "FATConfig", "IMPORTAR_PV_APP_PPL", giFilialEmpresa, "0")
    If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
    
    Importa_Dados = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados:

    Importa_Dados = gErr
     
    Select Case gErr
    
        Case 189912, 189915 To 189917, 189919, 189920, 189962, 189972, 190035, 190036

        Case 189913, 189914
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189918, 189921
        
        Case ERRO_SEM_MENSAGEM
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189922)
     
    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComandoAd) To UBound(alComandoAd)
        Call Comando_Fechar(alComandoAd(iIndice))
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    Call Fecha_Conexoes
    
    Exit Function

End Function

Private Function Importa_Dados_Cliente_Arq(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objArqImpArq As ClassArqImportacaoArq) As Long
        
Dim lErro As Long
Dim lSeq As Long
Dim lTransacao As Long
Dim objArqImportacao As New ClassArqImportacao

On Error GoTo Erro_Importa_Dados_Cliente_Arq

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 189923

    objArqImportacao.iTipoArq = objArqImpArq.objTipoArq.iCodigo
    objArqImportacao.sNomeArquivo = objArqImpArq.sNomeArquivo

    lErro = ArqImportacao_Insere(lComando1, objArqImportacao)
    If lErro <> SUCESSO Then gError 189924
    
    If UCase(right(objArqImpArq.sNomeArquivo, 4)) = ".XLS" Then
        
        lErro = Importa_Dados_Cliente_Arq_Xls(lComando2, objArqImpArq, objArqImportacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Else
    
        lErro = Importa_Dados_Cliente_Arq_Txt(lComando2, objArqImpArq, objArqImportacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 189928
    
    Importa_Dados_Cliente_Arq = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_Cliente_Arq:

    Importa_Dados_Cliente_Arq = gErr
     
    Select Case gErr
    
        Case 189923
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 189924
        
        Case 189928
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189929)
     
    End Select

    Call Transacao_Rollback
  
    Exit Function

End Function

Private Function Importa_Dados_Cliente_Arq_Txt(ByVal lComando, ByVal objArqImpArq As ClassArqImportacaoArq, ByVal objArqImportacao As ClassArqImportacao) As Long
        
Dim lErro As Long
Dim lSeq As Long
Dim lTotalRegistros As Long
Dim objIntegracaoCliente As ClassIntegracaoCliente
Dim objTS As TextStream
Dim objFSO As New FileSystemObject
Dim sTipoReg As String
Dim sRegistro As String
Dim sNomeFuncao As String
Dim bInterrompeLoop As Boolean

On Error GoTo Erro_Importa_Dados_Cliente_Arq_Txt

    Set objTS = objArqImpArq.objFile.OpenAsTextStream(ForReading)
    
    bInterrompeLoop = False
    
    lErro = CF("Importa_Dados_Cliente_Arq_Aux", sNomeFuncao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lSeq = 0
    Do While Not objTS.AtEndOfLine
        
        sRegistro = objTS.ReadLine
        
        lErro = CF(sNomeFuncao, lComando, objArqImpArq, objArqImportacao, sRegistro, objIntegracaoCliente, lSeq, bInterrompeLoop)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If bInterrompeLoop Then Exit Do
        
    Loop
         
    Importa_Dados_Cliente_Arq_Txt = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_Cliente_Arq_Txt:

    Importa_Dados_Cliente_Arq_Txt = gErr
     
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189929)
     
    End Select
  
    Exit Function

End Function

Private Function Importa_Dados_Cliente_Arq_Xls(ByVal lComando, ByVal objArqImpArq As ClassArqImportacaoArq, ByVal objArqImportacao As ClassArqImportacao) As Long

Dim lErro As Long
Dim sNomeFuncao As String

On Error GoTo Erro_Importa_Dados_Cliente_Arq_Xls

    lErro = CF("Importa_Dados_Cliente_Arq_Funcao_Cust", sNomeFuncao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Len(Trim(sNomeFuncao)) > 0 Then
    
        lErro = CF(sNomeFuncao, lComando, objArqImpArq, objArqImportacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    Importa_Dados_Cliente_Arq_Xls = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_Cliente_Arq_Xls:

    Importa_Dados_Cliente_Arq_Xls = gErr
     
    Exit Function
    
End Function

Public Function Importa_Dados_Cliente_Arq_Xls_Phar(ByVal lComando, ByVal objArqImpArq As ClassArqImportacaoArq, ByVal objArqImportacao As ClassArqImportacao) As Long
        
Dim lErro As Long
Dim lSeq As Long, bArqAberto As Boolean
Dim lTotalRegistros As Long, iLinha As Integer
Dim objIntegracaoCliente As ClassIntegracaoCliente
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
Dim objFile As File, sReg As String
Dim sVendedor As String, iVendedor As Integer
Dim sTabPreco As String, iTabPreco As Integer
Dim sTipo As String, iTipo As Integer
Dim alComando(1 To 3) As Long, iIndice As Integer, sCampo As String, sColuna As String
Dim objCliente As ClassCliente, iCodFilial As Integer, colEnderecos As colEndereco
Dim objExcelApp As New ClassExcelApp

On Error GoTo Erro_Importa_Dados_Cliente_Arq_Xls_Phar

    'abertura de comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211021
    Next
    
    Set objFile = objArqImpArq.objFile

'    'Abre o excel
'    lErro = CF("Excel_Abrir")
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    'Desabilita as mensagens do Excel
'    GL_objExcelSist.DisplayAlerts = False
'
'    Call GL_objExcelSist.Workbooks.Open(objFile.Path)
'
'    Set objPastaTrabalho = GL_objExcelSist.ActiveWorkBook
'
'    'Seleciona a planilha ativa na pasta de trabalho criada
'    Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
    
    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Abre o excel
    lErro = objExcelApp.Abrir_Planilha(objFile.Path)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sCampo = "Razão Social"
    sColuna = "E1"
    If UCase(objExcelApp.Obtem_Valor_Celula2(sColuna)) <> UCase(sCampo) Then gError 211022
    sCampo = "Número"
    sColuna = "L1"
    If UCase(objExcelApp.Obtem_Valor_Celula2(sColuna)) <> UCase(sCampo) Then gError 211023
    sCampo = "UF"
    sColuna = "Q1"
    If UCase(objExcelApp.Obtem_Valor_Celula2(sColuna)) <> UCase(sCampo) Then gError 211024
    sCampo = "Vendedor"
    sColuna = "U1"
    If UCase(objExcelApp.Obtem_Valor_Celula2(sColuna)) <> UCase(sCampo) Then gError 211025
    
    iLinha = 1
    Do While True
    
        iLinha = iLinha + 1
                   
        sReg = objExcelApp.Obtem_Valor_Celula2("A" & CStr(iLinha))
    
        If Len(Trim(sReg)) = 0 Then Exit Do
        
        Set objIntegracaoCliente = New ClassIntegracaoCliente
    
        objIntegracaoCliente.lNumIntArq = objArqImportacao.lNumIntArq
        objIntegracaoCliente.lSeqRegistro = lSeq
        objIntegracaoCliente.iTipoInt = TIPO_INTEGRACAO_IMPORTACAO
                
        objIntegracaoCliente.lCodCliente = 0
        objIntegracaoCliente.lCodExterno = 0
        objIntegracaoCliente.iAtivo = CLIENTE_ATIVO
        objIntegracaoCliente.iCodigoPais = PAIS_BRASIL
        objIntegracaoCliente.sRazaoSocial = left(Trim(objExcelApp.Obtem_Valor_Celula2("E" & CStr(iLinha))), STRING_CLIENTE_RAZAO_SOCIAL)
        objIntegracaoCliente.sContato = left(Trim(objExcelApp.Obtem_Valor_Celula2("F" & CStr(iLinha))), STRING_CONTATO)
        objIntegracaoCliente.sEmail = left(Trim(objExcelApp.Obtem_Valor_Celula2("G" & CStr(iLinha))), STRING_EMAIL)
        objIntegracaoCliente.sCgc = left(Trim(objExcelApp.Obtem_Valor_Celula2("H" & CStr(iLinha))), STRING_CGC)
        objIntegracaoCliente.sInscricaoEstadual = left(Trim(objExcelApp.Obtem_Valor_Celula2("I" & CStr(iLinha))), STRING_INSCR_EST)
        objIntegracaoCliente.sTipoLogradouro = Trim(objExcelApp.Obtem_Valor_Celula2("J" & CStr(iLinha)))
        objIntegracaoCliente.sLogradouro = Trim(objExcelApp.Obtem_Valor_Celula2("K" & CStr(iLinha)))
        
        If Len(Trim(objExcelApp.Obtem_Valor_Celula2("L" & CStr(iLinha)))) > 5 Then gError 211919
        lErro = Valor_Long_Critica(Trim(objExcelApp.Obtem_Valor_Celula2("L" & CStr(iLinha))))
        If lErro <> SUCESSO Then gError 211920
        
        objIntegracaoCliente.lNumero = objExcelApp.Obtem_Valor_Celula2("L" & CStr(iLinha))
        objIntegracaoCliente.sComplemento = Trim(objExcelApp.Obtem_Valor_Celula2("M" & CStr(iLinha)))
        objIntegracaoCliente.sBairro = left(Trim(objExcelApp.Obtem_Valor_Celula2("N" & CStr(iLinha))), STRING_BAIRRO)
        objIntegracaoCliente.sCEP = Trim(objExcelApp.Obtem_Valor_Celula2("O" & CStr(iLinha)))
        objIntegracaoCliente.sCidade = left(Trim(objExcelApp.Obtem_Valor_Celula2("P" & CStr(iLinha))), STRING_CIDADE)
        objIntegracaoCliente.sSiglaEstado = Trim(objExcelApp.Obtem_Valor_Celula2("Q" & CStr(iLinha)))
        
        If Len(Trim(objExcelApp.Obtem_Valor_Celula2("R" & CStr(iLinha)))) > 5 Then gError 211921
        lErro = Valor_Inteiro_Critica(Trim(objExcelApp.Obtem_Valor_Celula2("R" & CStr(iLinha))))
        If lErro <> SUCESSO Then gError 211922
        
        objIntegracaoCliente.iTelDDD1 = StrParaInt(objExcelApp.Obtem_Valor_Celula2("R" & CStr(iLinha)))
        
        objIntegracaoCliente.sTelNumero1 = Trim(objExcelApp.Obtem_Valor_Celula2("S" & CStr(iLinha)))
        objIntegracaoCliente.sFilialNome = "Matriz"
        
        objIntegracaoCliente.sTelefone1 = CStr(objIntegracaoCliente.iTelDDD1) & " " & objIntegracaoCliente.sTelNumero1
        
        objIntegracaoCliente.iCodigoPaisCobr = objIntegracaoCliente.iCodigoPais
        objIntegracaoCliente.iCodigoPaisEnt = objIntegracaoCliente.iCodigoPais
        
        objIntegracaoCliente.sSiglaEstadoCobr = objIntegracaoCliente.sSiglaEstado
        objIntegracaoCliente.sSiglaEstadoEnt = objIntegracaoCliente.sSiglaEstado
        
        If Len(Trim(objIntegracaoCliente.sCgc)) > 0 Then
            If Len(Trim(objIntegracaoCliente.sCgc)) <= STRING_CPF Then
                objIntegracaoCliente.sCgc = String(STRING_CPF - Len(Trim(objIntegracaoCliente.sCgc)), "0") & Trim(objIntegracaoCliente.sCgc)
            ElseIf Len(Trim(objIntegracaoCliente.sCgc)) <= STRING_CGC Then
                objIntegracaoCliente.sCgc = String(STRING_CGC - Len(Trim(objIntegracaoCliente.sCgc)), "0") & Trim(objIntegracaoCliente.sCgc)
            Else
                gError 211061
            End If
        Else
            gError 211062
        End If
        
        'Compatibiliza o tipo de logradouro
        Select Case UCase(objIntegracaoCliente.sTipoLogradouro)
        
            Case "R", "R.", "RUA"
                objIntegracaoCliente.sTipoLogradouro = "R"
                
            Case "AV", "AV.", "AVENIDA"
                objIntegracaoCliente.sTipoLogradouro = "AV"
            
            Case Else
        
                Set objCliente = New ClassCliente
                Set colEnderecos = New colEndereco
                objCliente.sCgc = objIntegracaoCliente.sCgc
                
                lErro = CF("Cliente_Le_Cgc", objCliente, iCodFilial)
                If lErro <> SUCESSO And lErro <> 6710 Then gError ERRO_SEM_MENSAGEM
                
                If lErro = SUCESSO Then
                
                    lErro = CF("Enderecos_Le_Cliente", colEnderecos, objCliente)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    If Len(Trim(colEnderecos.Item(1).sTipoLogradouro)) > 0 Then
                        objIntegracaoCliente.sTipoLogradouro = colEnderecos.Item(1).sTipoLogradouro
                    Else
                        'Se não identificou e não está cadastrado joga como parte do endereço
                        If Len(Trim(objIntegracaoCliente.sTipoLogradouro)) > 0 Then
                            objIntegracaoCliente.sTipoLogradouro = ""
                            objIntegracaoCliente.sLogradouro = objIntegracaoCliente.sTipoLogradouro & " " & objIntegracaoCliente.sLogradouro
                        End If
                    End If
                
                Else
                    'Se não identificou e não está cadastrado joga como parte do endereço
                    If Len(Trim(objIntegracaoCliente.sTipoLogradouro)) > 0 Then
                        objIntegracaoCliente.sTipoLogradouro = ""
                        objIntegracaoCliente.sLogradouro = objIntegracaoCliente.sTipoLogradouro & " " & objIntegracaoCliente.sLogradouro
                    End If
                End If
                
        End Select
        
        objIntegracaoCliente.sEndereco = objIntegracaoCliente.sTipoLogradouro
        objIntegracaoCliente.sEndereco = objIntegracaoCliente.sEndereco & IIf(Len(Trim(objIntegracaoCliente.sEndereco)) <> 0, " ", "") & objIntegracaoCliente.sLogradouro
        objIntegracaoCliente.sEndereco = objIntegracaoCliente.sEndereco & IIf(objIntegracaoCliente.lNumero <> 0, ", " & CStr(objIntegracaoCliente.lNumero), "")
        objIntegracaoCliente.sEndereco = objIntegracaoCliente.sEndereco & IIf(Len(Trim(objIntegracaoCliente.sComplemento)) <> 0, " - " & objIntegracaoCliente.sComplemento, "")
        objIntegracaoCliente.sEndereco = left(objIntegracaoCliente.sEndereco, STRING_ENDERECO)
        
        sTabPreco = objExcelApp.Obtem_Valor_Celula2("T" & CStr(iLinha))
        sVendedor = objExcelApp.Obtem_Valor_Celula2("U" & CStr(iLinha))
        sTipo = objExcelApp.Obtem_Valor_Celula2("D" & CStr(iLinha))
   
        'Localiza o vendedor
        iVendedor = 0
        Select Case sVendedor
            Case "Cassiano Grando", "Reginaldo Rocha"
                sVendedor = "Enio Medeiros"
            Case "Janaina Santiago"
                sVendedor = "José Pinheiro"
        End Select
        lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Vendedores WHERE NomeReduzido = ? OR Nome = ?", iVendedor, sVendedor, sVendedor)
        If lErro <> AD_SQL_SUCESSO Then gError 211026
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211027
        
        If lErro <> AD_SQL_SUCESSO Then gError 211028  'Não localizou o vendedor
        
        lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM TiposDeCliente WHERE Descricao = ?", iTipo, sTipo)
        If lErro <> AD_SQL_SUCESSO Then gError 211029
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211030
        
        If lErro <> AD_SQL_SUCESSO Then gError 211031  'Não localizou o tipo
    
        Select Case sTabPreco
            Case "Tabela Profissional"
                iTabPreco = 7
            Case "Tabela Farma"
                iTabPreco = 1
            Case "Tabela Hospitalar"
                iTabPreco = 16
            Case "Tabela Licitações"
                iTabPreco = 21
            Case Else
                 gError 211032  'Não localizou a tabela de preço
        End Select
        
        objIntegracaoCliente.iVendedor = iVendedor
        objIntegracaoCliente.iTabelaPreco = iTabPreco
        objIntegracaoCliente.sTipoCliente = sTipo
        objIntegracaoCliente.sNomeReduzido = left(objIntegracaoCliente.sRazaoSocial, STRING_CLIENTE_NOME_REDUZIDO)
    
        'VALIDA OS DADOS
        If Len(Trim(objIntegracaoCliente.sCEP)) <> 8 And Len(Trim(objIntegracaoCliente.sCEP)) <> 0 Then gError 211033
        If Len(objIntegracaoCliente.sSiglaEstado) <> 2 Then gError 213222
        If Len(objIntegracaoCliente.sTelefone1) > STRING_TELEFONE Then gError 213223
   
        lErro = IntegracaoCliente_Insere(lComando, objIntegracaoCliente)
        If lErro <> SUCESSO Then
            gError ERRO_SEM_MENSAGEM
        End If
    
    Loop
        
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha o Excel
    'Call CF("Excel_Fechar")
         
    lErro = objExcelApp.Fechar()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
         
    Importa_Dados_Cliente_Arq_Xls_Phar = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_Cliente_Arq_Xls_Phar:

    Importa_Dados_Cliente_Arq_Xls_Phar = gErr
     
    Select Case gErr
    
        Case 211020
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 211022 To 211025
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_FORMATO_DIF_ESPERADO", gErr, objFile.Name, sColuna, sCampo)
    
        Case 211026, 211027
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)
    
        Case 211028
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDOR1", gErr, sVendedor)
    
        Case 211029, 211030
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDECLIENTE", gErr)
    
        Case 211031
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_ENCONTRADO2", gErr, sTipo)
    
        Case 211032
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_CADASTRADA", gErr, sTabPreco)
            
        Case 211033
            Call Rotina_Erro(vbOKOnly, "ERRO_CEP_INVALIDO", gErr, objIntegracaoCliente.sCEP)
            
        Case 211061
            Call Rotina_Erro(vbOKOnly, "ERRO_CGC_INVALIDO", gErr, objIntegracaoCliente.sCgc)
            
        Case 211062
            Call Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_INFORMADO", gErr)
            
        Case 211919, 211920
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_END_INVALIDO_LINHA", gErr, iLinha)
    
        Case 211921, 211922
            Call Rotina_Erro(vbOKOnly, "ERRO_DDD_INVALIDO_LINHA", gErr, iLinha)
            
        Case 213222
            Call Rotina_Erro(vbOKOnly, "ERRO_SIGLA_ESTADO_NAO_CADASTRADA", gErr, objIntegracaoCliente.sSiglaEstado)
        
        Case 213223
            Call Rotina_Erro(vbOKOnly, "ERRO_TELEFONE1_MAIOR_MAXIMO", gErr, Len(objIntegracaoCliente.sTelefone1), STRING_TELEFONE)
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211034)
     
    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Call CF("Excel_Fechar")
    Call objExcelApp.Fechar
  
    Exit Function

End Function

Public Function Importa_Dados_Cliente_Arq_Aux1(ByVal lComando2 As Long, ByVal objArqImpArq As ClassArqImportacaoArq, ByVal objArqImportacao As ClassArqImportacao, ByVal sRegistro As String, objIntegracaoCliente As ClassIntegracaoCliente, lSeq As Long, bInterrompeLoop As Boolean) As Long
'Importação padrão de clientes

Dim lErro As Long
Dim lTotalRegistros As Long
Dim sTipoReg As String

On Error GoTo Erro_Importa_Dados_Cliente_Arq_Aux1

    Set objIntegracaoCliente = New ClassIntegracaoCliente

    sTipoReg = left(sRegistro, 2)
    
    If sTipoReg = "01" Then
        
        lSeq = lSeq + 1
    
        lErro = objIntegracaoCliente.Desmonta_Registro_Cliente(sRegistro)
        If lErro <> SUCESSO Then gError 189925
        
        objIntegracaoCliente.lNumIntArq = objArqImportacao.lNumIntArq
        objIntegracaoCliente.lSeqRegistro = lSeq
        objIntegracaoCliente.iTipoInt = TIPO_INTEGRACAO_IMPORTACAO
        
        lErro = IntegracaoCliente_Insere(lComando2, objIntegracaoCliente)
        If lErro <> SUCESSO Then gError 189926
        
    Else
    
        lTotalRegistros = StrParaLong(Mid(sRegistro, 3, 6))
        
        If lTotalRegistros <> lSeq Then gError 189927
        
        bInterrompeLoop = True
    
    End If
        
    Exit Function
    
Erro_Importa_Dados_Cliente_Arq_Aux1:

    Importa_Dados_Cliente_Arq_Aux1 = gErr
     
    Select Case gErr
    
        Case 189925 To 189926

        Case 189927
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_REG_ARQ_DIF_INF", gErr, objArqImpArq.sNomeArquivo, lTotalRegistros, lSeq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189929)
     
    End Select

    Exit Function
    
End Function

Public Function Importa_Dados_Cliente_Arq_Aux2(ByVal lComando2 As Long, ByVal objArqImpArq As ClassArqImportacaoArq, ByVal objArqImportacao As ClassArqImportacao, ByVal sRegistro As String, objIntegracaoCliente As ClassIntegracaoCliente, lSeq As Long, bInterrompeLoop As Boolean) As Long
'Importação de arquivos VW

Dim lErro As Long
Dim lTotalRegistros As Long
Dim sTipoReg As String
Dim iPOS As Integer

On Error GoTo Erro_Importa_Dados_Cliente_Arq_Aux2

    sTipoReg = left(sRegistro, 3)
    iPOS = 4
    
    lSeq = lSeq + 1
    
    If sTipoReg = "CL1" Then
        
        Set objIntegracaoCliente = New ClassIntegracaoCliente
        
        objIntegracaoCliente.lNumIntArq = objArqImportacao.lNumIntArq
        objIntegracaoCliente.lSeqRegistro = lSeq
        objIntegracaoCliente.iTipoInt = TIPO_INTEGRACAO_IMPORTACAO
                
        iPOS = iPOS + 1
        objIntegracaoCliente.lCodCliente = StrParaLong(Mid(sRegistro, iPOS, 4))
        iPOS = iPOS + 4
        objIntegracaoCliente.iCodFilial = 1
        objIntegracaoCliente.sRazaoSocial = Trim(Mid(sRegistro, iPOS, 25))
        iPOS = iPOS + 25
        objIntegracaoCliente.sNomeReduzido = left(objIntegracaoCliente.sRazaoSocial, STRING_CLIENTE_NOME_REDUZIDO)
        objIntegracaoCliente.sEndereco = Trim(Mid(sRegistro, iPOS, 30))
        iPOS = iPOS + 30
        objIntegracaoCliente.sCidade = Trim(Mid(sRegistro, iPOS, 15))
        iPOS = iPOS + 15
        objIntegracaoCliente.sBairro = Trim(Mid(sRegistro, iPOS, 15))
        iPOS = iPOS + 15
        objIntegracaoCliente.sCEP = Trim(Mid(sRegistro, iPOS, 8))
        iPOS = iPOS + 8
        objIntegracaoCliente.sSiglaEstado = Trim(Mid(sRegistro, iPOS, 2))
        iPOS = iPOS + 2
        iPOS = iPOS + 2
        objIntegracaoCliente.iCodigoPais = PAIS_BRASIL
        objIntegracaoCliente.sCgc = Trim(Mid(sRegistro, iPOS, 14))
        iPOS = iPOS + 14
        objIntegracaoCliente.lCodExterno = StrParaLong(Mid(sRegistro, iPOS, 6))
        iPOS = iPOS + 6
        objIntegracaoCliente.sObservacao = "CL1-CAT-VW: " & Trim(Mid(sRegistro, iPOS, 1))
        iPOS = iPOS + 1
        
    ElseIf sTipoReg = "CL2" Then
    
        iPOS = iPOS + 23
        objIntegracaoCliente.dLimiteCredito = StrParaDbl(Mid(sRegistro, iPOS, 15)) / 100
        iPOS = iPOS + 15
        'objIntegracaoCliente.sObservacao = objIntegracaoCliente.sObservacao & "; CL2-CONDPGTO: " & Trim(Mid(sRegistro, iPOS, 3))
        iPOS = iPOS + 3
        'objIntegracaoCliente.sObservacao = objIntegracaoCliente.sObservacao & "; CL2-DESCRPGTO: " & Trim(Mid(sRegistro, iPOS, 30))
        objIntegracaoCliente.iCondicaoPagto = 7
        iPOS = iPOS + 30
        iPOS = iPOS + 7
        objIntegracaoCliente.sInscricaoEstadual = Trim(Mid(sRegistro, iPOS, 15))
        iPOS = iPOS + 15
        objIntegracaoCliente.sTelefone1 = left(Trim(Mid(sRegistro, iPOS, 20)), STRING_TELEFONE)
        iPOS = iPOS + 20
        iPOS = iPOS + 3
        objIntegracaoCliente.sObservacao = objIntegracaoCliente.sObservacao & "; CL2-COD-VW: " & Trim(Mid(sRegistro, iPOS, 6))
        iPOS = iPOS + 6
        objIntegracaoCliente.sObservacao = objIntegracaoCliente.sObservacao & "; CL2-CAT-VW: " & Trim(Mid(sRegistro, iPOS, 1))
        iPOS = iPOS + 1
        
        objIntegracaoCliente.sObservacao = left(objIntegracaoCliente.sObservacao, STRING_CLIENTE_OBSERVACAO)
        
        If objIntegracaoCliente.sCgc = "00000000000000" Then objIntegracaoCliente.sCgc = ""
        If objIntegracaoCliente.sInscricaoEstadual = "000000000000000" Then objIntegracaoCliente.sInscricaoEstadual = ""
        If objIntegracaoCliente.sCEP = "00000000" Then objIntegracaoCliente.sCEP = ""
        If objIntegracaoCliente.sSiglaEstado = "99" Then objIntegracaoCliente.sSiglaEstado = ""
    
        objIntegracaoCliente.sObservacaoFilial = ""
        objIntegracaoCliente.sTipoCliente = ""
    
        objIntegracaoCliente.iTipo = 1
        objIntegracaoCliente.iCondicaoPagto = 0
        objIntegracaoCliente.iAtivo = CLIENTE_ATIVO
        objIntegracaoCliente.sFilialNome = "Matriz"
        objIntegracaoCliente.sRG = ""
        objIntegracaoCliente.sInscricaoMunicipal = ""
        objIntegracaoCliente.sTelefone2 = ""
        objIntegracaoCliente.sEmail = ""
        objIntegracaoCliente.sFax = ""
        objIntegracaoCliente.sContato = ""
        objIntegracaoCliente.sEnderecoEnt = objIntegracaoCliente.sEndereco
        objIntegracaoCliente.sBairroEnt = objIntegracaoCliente.sBairro
        objIntegracaoCliente.sCidadeEnt = objIntegracaoCliente.sCidade
        objIntegracaoCliente.sSiglaEstadoEnt = objIntegracaoCliente.sSiglaEstado
        objIntegracaoCliente.iCodigoPaisEnt = objIntegracaoCliente.iCodigoPais
        objIntegracaoCliente.sCEPEnt = objIntegracaoCliente.sCEP
        objIntegracaoCliente.sTelefone1Ent = objIntegracaoCliente.sTelefone1
        objIntegracaoCliente.sTelefone2Ent = objIntegracaoCliente.sTelefone2
        objIntegracaoCliente.sEmailEnt = objIntegracaoCliente.sEmail
        objIntegracaoCliente.sFaxEnt = objIntegracaoCliente.sFax
        objIntegracaoCliente.sContatoEnt = objIntegracaoCliente.sContato
        objIntegracaoCliente.sEnderecoCobr = objIntegracaoCliente.sEndereco
        objIntegracaoCliente.sBairroCobr = objIntegracaoCliente.sBairro
        objIntegracaoCliente.sCidadeCobr = objIntegracaoCliente.sCidade
        objIntegracaoCliente.sSiglaEstadoCobr = objIntegracaoCliente.sSiglaEstado
        objIntegracaoCliente.iCodigoPaisCobr = objIntegracaoCliente.iCodigoPais
        objIntegracaoCliente.sCEPCobr = objIntegracaoCliente.sCEP
        objIntegracaoCliente.sTelefone1Cobr = objIntegracaoCliente.sTelefone1
        objIntegracaoCliente.sTelefone2Cobr = objIntegracaoCliente.sTelefone2
        objIntegracaoCliente.sEmailCobr = objIntegracaoCliente.sEmail
        objIntegracaoCliente.sFaxCobr = objIntegracaoCliente.sFax
        objIntegracaoCliente.sContatoCobr = objIntegracaoCliente.sContato
        objIntegracaoCliente.iVendedor = 0
    
        lErro = IntegracaoCliente_Insere(lComando2, objIntegracaoCliente)
        If lErro <> SUCESSO Then
            gError 189926
        End If
    
    ElseIf sTipoReg = "FTP" Then
    
        lTotalRegistros = StrParaLong(Mid(sRegistro, 9, 9))
        
        If lTotalRegistros <> lSeq Then gError 189927
        
        bInterrompeLoop = True
    
    End If
        
    Exit Function
    
Erro_Importa_Dados_Cliente_Arq_Aux2:

    Importa_Dados_Cliente_Arq_Aux2 = gErr
     
    Select Case gErr
    
        Case 189925 To 189926

        Case 189927
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_REG_ARQ_DIF_INF", gErr, objArqImpArq.sNomeArquivo, lTotalRegistros, lSeq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189929)
     
    End Select

    Exit Function
    
End Function

Private Function Importa_Dados_PV_Arq(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objArqImpArq As ClassArqImportacaoArq) As Long
        
Dim lErro As Long
Dim lTransacao As Long
Dim objArqImportacao As New ClassArqImportacao

On Error GoTo Erro_Importa_Dados_PV_Arq

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190037

    objArqImportacao.iTipoArq = objArqImpArq.objTipoArq.iCodigo
    objArqImportacao.sNomeArquivo = objArqImpArq.sNomeArquivo

    lErro = ArqImportacao_Insere(lComando1, objArqImportacao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If UCase(right(objArqImpArq.sNomeArquivo, 4)) = ".XLS" Then
        
        lErro = Importa_Dados_PV_Arq_Xls(lComando2, objArqImpArq, objArqImportacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Else
    
        lErro = Importa_Dados_PV_Arq_Txt(lComando2, objArqImpArq, objArqImportacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190042
    
    Importa_Dados_PV_Arq = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_PV_Arq:

    Importa_Dados_PV_Arq = gErr
     
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 190037
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 190042
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190043)
     
    End Select

    Call Transacao_Rollback
  
    Exit Function

End Function

Private Function Importa_Dados_PV_Arq_Txt(ByVal lComando As Long, ByVal objArqImpArq As ClassArqImportacaoArq, ByVal objArqImportacao As ClassArqImportacao) As Long
        
Dim lErro As Long
Dim lSeq As Long
Dim lTotalRegistros As Long
Dim objIntegracaoPV As ClassIntegracaoPV
Dim objTS As TextStream
Dim objFSO As New FileSystemObject
Dim sTipoReg As String
Dim sRegistro As String

On Error GoTo Erro_Importa_Dados_PV_Arq_Txt

    Set objTS = objArqImpArq.objFile.OpenAsTextStream(ForReading)
    
    lSeq = 0
    Do While Not objTS.AtEndOfLine
                
        Set objIntegracaoPV = New ClassIntegracaoPV
    
        sRegistro = objTS.ReadLine
        
        sTipoReg = left(sRegistro, 2)
        
        If sTipoReg = "01" Then
            
            lSeq = lSeq + 1
        
            lErro = objIntegracaoPV.Desmonta_Registro_PV(sRegistro)
            If lErro <> SUCESSO Then gError 190039
            
            objIntegracaoPV.lNumIntArq = objArqImportacao.lNumIntArq
            objIntegracaoPV.lSeqRegistro = lSeq
            objIntegracaoPV.iTipoInt = TIPO_INTEGRACAO_IMPORTACAO
            
            lErro = IntegracaoPV_Insere(lComando, objIntegracaoPV)
            If lErro <> SUCESSO Then gError 190040
            
        Else
        
            lTotalRegistros = StrParaLong(Mid(sRegistro, 3, 6))
            
            If lTotalRegistros <> lSeq Then gError 190041
            
            Exit Do
        
        End If
        
    Loop
    
    Importa_Dados_PV_Arq_Txt = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_PV_Arq_Txt:

    Importa_Dados_PV_Arq_Txt = gErr
     
    Select Case gErr
    
        Case 190038 To 190040

        Case 190041
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_REG_ARQ_DIF_INF", gErr, objArqImpArq.sNomeArquivo, lTotalRegistros, lSeq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190043)
     
    End Select
  
    Exit Function

End Function

Private Function Importa_Dados_PV_Arq_Xls(ByVal lComando As Long, ByVal objArqImpArq As ClassArqImportacaoArq, ByVal objArqImportacao As ClassArqImportacao) As Long

Dim lErro As Long
Dim sNomeFuncao As String

On Error GoTo Erro_Importa_Dados_PV_Arq_Xls

    lErro = CF("Importa_Dados_PV_Arq_Funcao_Cust", sNomeFuncao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Len(Trim(sNomeFuncao)) > 0 Then
    
        lErro = CF(sNomeFuncao, lComando, objArqImpArq, objArqImportacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    Importa_Dados_PV_Arq_Xls = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_PV_Arq_Xls:

    Importa_Dados_PV_Arq_Xls = gErr
     
    Exit Function
    
End Function

Public Function Importa_Dados_PV_Arq_Xls_Phar(ByVal lComando As Long, ByVal objArqImpArq As ClassArqImportacaoArq, ByVal objArqImportacao As ClassArqImportacao) As Long
        
Dim lErro As Long, alComando(1 To 3) As Long
Dim lSeq As Long, iLinha As Integer, bAchou As Boolean
Dim lTotalRegistros As Long, iIndice As Integer
Dim objIntegracaoPV As ClassIntegracaoPV
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
Dim objFile As File, sReg As String, bArqAberto As Boolean
Dim colProdutos As New Collection, objProduto As ClassProduto, colProdCol As New Collection
Dim colItensPV As Collection, objItemPV As ClassItemPedido
Dim iColuna As Integer, iPOS As Integer, bNaoID As Boolean
Dim iCol_Codigo As Integer, iCol_Status As Integer, iCol_CGC As Integer
Dim iCol_CondPagto As Integer, iCol_PedCli As Integer, iCol_Txt1 As Integer
Dim iCol_Txt2 As Integer, iCol_Obs As Integer, iCol_Frete As Integer, iCol_Desconto As Integer
Dim iNumCols As Integer, sStatus As String, sColuna As String, iCondPagto As Integer
Dim objCliente As ClassCliente, iCodFilial As Integer, sProdMask As String
Dim objTabelaPrecoItem As ClassTabelaPrecoItem, iItem As Integer, dPercDesc As Double
Dim objExcelApp As New ClassExcelApp

On Error GoTo Erro_Importa_Dados_PV_Arq_Xls_Phar

    'abertura de comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211035
    Next
    
    Set objFile = objArqImpArq.objFile

    lErro = CF("Produto_Le_Faixa", "", "", colProdutos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'    'Abre o excel
'    lErro = CF("Excel_Abrir")
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    'Desabilita as mensagens do Excel
'    GL_objExcelSist.DisplayAlerts = False
'
'    Call GL_objExcelSist.Workbooks.Open(objFile.Path)
'
'    Set objPastaTrabalho = GL_objExcelSist.ActiveWorkBook
'
'    'Seleciona a planilha ativa na pasta de trabalho criada
'    Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet

    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Abre o excel
    lErro = objExcelApp.Abrir_Planilha(objFile.Path)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iCol_Codigo = 1
    iColuna = 1
    Do While True
        iColuna = iColuna + 1
        sReg = objExcelApp.Obtem_Valor_Celula(1, iColuna)
        If Len(Trim(sReg)) = 0 Then Exit Do
        iPOS = InStr(1, sReg, "|")
        If iPOS > 0 Then sReg = left(sReg, iPOS - 1)
        Select Case Trim(UCase(sReg))
        
            Case UCase("Status")
                iCol_Status = iColuna
                
            Case UCase("CPF/CNPJ")
                iCol_CGC = iColuna
            
            Case UCase("Condição de pagamento")
                iCol_CondPagto = iColuna
        
            Case UCase("Frete")
                iCol_Frete = iColuna
        
            Case UCase("Desconto")
                iCol_Desconto = iColuna
        
            Case UCase("Pedido cliente")
                iCol_PedCli = iColuna
        
            Case UCase("Texto linha 1")
                iCol_Txt1 = iColuna
            
            Case UCase("Texto linha 2")
                iCol_Txt2 = iColuna
            
            Case UCase("Observações")
                iCol_Obs = iColuna
        
            Case Else 'Verifica se é um produto
                
                bAchou = False
                For Each objProduto In colProdutos
                    If Trim(UCase(sReg)) = UCase(objProduto.sNomeReduzido) Then
                        objProduto.iNivel = iColuna 'Guarda a coluna em que o produto se encontra
                        colProdCol.Add objProduto
                        bAchou = True
                        Exit For
                    End If
                Next
                'Se tem | é porque é uma coluna de produto
                'Se não achar tem que dar o erro
                sColuna = sReg
                If Not bAchou And iPOS <> 0 Then gError 211055
        
        End Select
        
    Loop
    iNumCols = iColuna
    
    sColuna = "Status"
    If iCol_Status = 0 Then gError 211036
    sColuna = "CPF/CNPJ"
    If iCol_CGC = 0 Then gError 211037
    sColuna = "Condição de pagamento"
    If iCol_CondPagto = 0 Then gError 211038
    sColuna = "Frete"
    If iCol_Frete = 0 Then gError 211039
    sColuna = "Desconto"
    If iCol_Desconto = 0 Then gError 211040
    sColuna = "Pedido cliente"
    If iCol_PedCli = 0 Then gError 211041
    sColuna = "Texto linha 1"
    If iCol_Txt1 = 0 Then gError 211042
    sColuna = "Texto linha 2"
    If iCol_Txt2 = 0 Then gError 211043
    sColuna = "Observações"
    If iCol_Obs = 0 Then gError 211044
    
    iLinha = 1
    Do While True
    
        Set colItensPV = New Collection
        Set objCliente = New ClassCliente
    
        iLinha = iLinha + 1
                   
        sReg = objExcelApp.Obtem_Valor_Celula2("A" & CStr(iLinha))
    
        If Len(Trim(sReg)) = 0 Then Exit Do
        
        sStatus = objExcelApp.Obtem_Valor_Celula(iLinha, iCol_Status)
        
        If Trim(UCase(sStatus)) = UCase("Confirmed") Then
             
             Set objIntegracaoPV = New ClassIntegracaoPV
         
             objIntegracaoPV.lNumIntArq = objArqImportacao.lNumIntArq
             objIntegracaoPV.lSeqRegistro = lSeq
             objIntegracaoPV.iTipoInt = TIPO_INTEGRACAO_IMPORTACAO
             
             'MONTA A COLEÇÃO PRODUTO A PRODUTO
             For Each objProduto In colProdCol
                 sReg = objExcelApp.Obtem_Valor_Celula(iLinha, objProduto.iNivel)
                 If Len(Trim(sReg)) > 0 Then
                     Set objItemPV = New ClassItemPedido
                     objItemPV.sProduto = objProduto.sCodigo
                     
                     sColuna = objProduto.sNomeReduzido
                     If Not IsNumeric(sReg) Then gError 211045
                     
                     objItemPV.dQuantidade = StrParaDbl(sReg)
                     objItemPV.sUnidadeMed = objProduto.sSiglaUMVenda
                     
                     colItensPV.Add objItemPV
                 End If
             Next
             
             objIntegracaoPV.dtDataEmissao = Date
             objIntegracaoPV.dtDataEntrega = DATA_NULA
             
             sColuna = "Desconto"
             sReg = Replace(objExcelApp.Obtem_Valor_Celula(iLinha, iCol_Desconto), "%", "")
             If Not IsNumeric(sReg) Then gError 211046
             dPercDesc = StrParaDbl(sReg) / 100
             
             sColuna = "Frete"
             sReg = Replace(objExcelApp.Obtem_Valor_Celula(iLinha, iCol_Frete), "R$", "")
             If Not IsNumeric(sReg) Then gError 211047
             objIntegracaoPV.dValorFretePedido = StrParaDbl(sReg)
             
             objIntegracaoPV.sCGCCliente = Trim(objExcelApp.Obtem_Valor_Celula(iLinha, iCol_CGC))
             objIntegracaoPV.lCodPedidoExterno = StrParaLong(objExcelApp.Obtem_Valor_Celula(iLinha, iCol_Codigo))
             objIntegracaoPV.sNomeCondPagtoExterno = objExcelApp.Obtem_Valor_Celula(iLinha, iCol_CondPagto)
             objIntegracaoPV.sMensagemPedido = objExcelApp.Obtem_Valor_Celula(iLinha, iCol_Txt1)
             If Len(Trim(objExcelApp.Obtem_Valor_Celula(iLinha, iCol_Txt2))) > 0 Then
                If Len(Trim(objIntegracaoPV.sMensagemPedido)) <> 0 Then objIntegracaoPV.sMensagemPedido = objIntegracaoPV.sMensagemPedido & " "
                objIntegracaoPV.sMensagemPedido = objIntegracaoPV.sMensagemPedido & objExcelApp.Obtem_Valor_Celula(iLinha, iCol_Txt2)
             End If
             If Len(Trim(objExcelApp.Obtem_Valor_Celula(iLinha, iCol_Obs))) > 0 Then
                If Len(Trim(objIntegracaoPV.sMensagemPedido)) <> 0 Then objIntegracaoPV.sMensagemPedido = objIntegracaoPV.sMensagemPedido & " "
                objIntegracaoPV.sMensagemPedido = objIntegracaoPV.sMensagemPedido & objExcelApp.Obtem_Valor_Celula(iLinha, iCol_Obs)
             End If
            objIntegracaoPV.sMensagemPedido = left(objIntegracaoPV.sMensagemPedido, STRING_PEDIDOVENDA_MENSAGEM_NOTA)
                        
            If Len(Trim(objIntegracaoPV.sCGCCliente)) > 0 Then
                If Len(Trim(objIntegracaoPV.sCGCCliente)) <= STRING_CPF Then
                    objIntegracaoPV.sCGCCliente = String(STRING_CPF - Len(Trim(objIntegracaoPV.sCGCCliente)), "0") & Trim(objIntegracaoPV.sCGCCliente)
                ElseIf Len(Trim(objIntegracaoPV.sCGCCliente)) <= STRING_CGC Then
                    objIntegracaoPV.sCGCCliente = String(STRING_CGC - Len(Trim(objIntegracaoPV.sCGCCliente)), "0") & Trim(objIntegracaoPV.sCGCCliente)
                Else
                    gError 211061
                End If
            Else
                gError 211062
            End If
            
            objCliente.sCgc = objIntegracaoPV.sCGCCliente
             
            lErro = CF("Cliente_Le_Cgc", objCliente, iCodFilial)
            If lErro <> SUCESSO And lErro <> 6710 Then gError ERRO_SEM_MENSAGEM
             
            If lErro <> SUCESSO Then gError 211048
             
            lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM CondicoesPagto WHERE DescReduzida = ? OR Descricao = ?", iCondPagto, objIntegracaoPV.sNomeCondPagtoExterno, objIntegracaoPV.sNomeCondPagtoExterno)
            If lErro <> AD_SQL_SUCESSO Then gError 211051
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211052
            
            If lErro <> AD_SQL_SUCESSO Then gError 211053  'Não localizou o vendedor
            
            objIntegracaoPV.sPedRepr = left(objExcelApp.Obtem_Valor_Celula(iLinha, iCol_PedCli), STRING_PEDIDOVENDA_PEDIDO_REPRESENTANTE)
            
            objIntegracaoPV.iCodFilialCorporator = iCodFilial
            objIntegracaoPV.iCodTabelaPreco = objCliente.iTabelaPreco
            objIntegracaoPV.lCodClienteCorporator = objCliente.lCodigo
            objIntegracaoPV.iFilialEmpresaFat = giFilialEmpresa
            objIntegracaoPV.iCodVendedor = objCliente.iVendedor
            objIntegracaoPV.iCodCondPagto = iCondPagto
                    
            objIntegracaoPV.iTrazerCanalVendaAuto = MARCADO
            objIntegracaoPV.iTrazerComissaoAuto = MARCADO
            objIntegracaoPV.iTrazerDescricaoAuto = MARCADO
            objIntegracaoPV.iTrazerMensagemAuto = DESMARCADO
            objIntegracaoPV.iTrazerPesoAuto = MARCADO
            objIntegracaoPV.iTrazerReservaAuto = MARCADO
            objIntegracaoPV.iTrazerTranspAuto = MARCADO
        
             'VALIDA OS DADOS
             
             If colItensPV.Count = 0 Then gError 211054
             
             iItem = 0
             For Each objItemPV In colItensPV
             
                Set objTabelaPrecoItem = New ClassTabelaPrecoItem
             
                objTabelaPrecoItem.iCodTabela = objIntegracaoPV.iCodTabelaPreco
                objTabelaPrecoItem.iFilialEmpresa = objIntegracaoPV.iFilialEmpresaFat
                objTabelaPrecoItem.sCodProduto = objItemPV.sProduto
                
                lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
                
                If lErro <> SUCESSO Then
                    objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
                
                    lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                    If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
                
                    If lErro <> SUCESSO Then gError 211049 'sem preço
                End If
            
                iItem = iItem + 1
                
                Call Mascara_RetornaProdutoTela(objItemPV.sProduto, sProdMask)
                
                objIntegracaoPV.sCodProduto = objItemPV.sProduto
                objIntegracaoPV.sCodProdutoExterno = sProdMask
                objIntegracaoPV.dQuantidadePedida = objItemPV.dQuantidade
                objIntegracaoPV.sUM = objItemPV.sUnidadeMed
                objIntegracaoPV.iItem = iItem
                objIntegracaoPV.dPrecoUnitario = objTabelaPrecoItem.dPreco
                
                objIntegracaoPV.dValorDescontoItem = objIntegracaoPV.dPrecoUnitario * objIntegracaoPV.dQuantidadePedida * dPercDesc
                
                lErro = IntegracaoPV_Insere(lComando, objIntegracaoPV)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
             Next
             
        End If
    
    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha o Excel
    'Call CF("Excel_Fechar")
         
    lErro = objExcelApp.Fechar()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Importa_Dados_PV_Arq_Xls_Phar = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_PV_Arq_Xls_Phar:

    Importa_Dados_PV_Arq_Xls_Phar = gErr
     
    Select Case gErr
    
        Case 211034
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 211035 To 211044
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_COLUNA_NAO_ENCONTRADA", gErr, objFile.Name, sColuna)
    
        Case 211045 To 211047
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_COLUNA_NAO_NUMERICA", gErr, iLinha, sColuna, objFile.Name)
    
        Case 211048
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_CGC_NAO_ENCONTRADA", gErr, objCliente.sCgc)
    
        Case 211049
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITEM_INEXISTENTE", gErr, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto)
            
        Case 211051, 211052
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONDICOESPAGTO", gErr)
        
        Case 211053
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA1", gErr, objIntegracaoPV.sNomeCondPagtoExterno)
    
        Case 211054
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_PV_SEM_PRODUTO", gErr, objFile.Name)
    
        Case 211055
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_PV_PRODUTO_NAO_LOC", gErr, objFile.Name, sColuna)
    
        Case 211061
            Call Rotina_Erro(vbOKOnly, "ERRO_CGC_INVALIDO", gErr, objIntegracaoPV.sCGCCliente)
            
        Case 211062
            Call Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_INFORMADO", gErr)
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211050)
     
    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Call CF("Excel_Fechar")
    Call objExcelApp.Fechar
  
    Exit Function

End Function

Private Function Importa_Dados_Cliente(ByVal lComandoExt1 As Long, ByVal lComandoExt2 As Long, ByVal lComandoExt3 As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal objTipoArq As ClassTipoArqIntegracao) As Long
        
Dim lErro As Long
Dim lSeq As Long
Dim lTransacao As Long
Dim objIntegracaoCliente As ClassIntegracaoCliente
Dim tIntegracaoCliente As typeIntegracaoCliente
Dim colEnderecos As colEndereco
Dim colEndereco As Collection
Dim objEndereco As ClassEndereco
Dim objCliente As ClassCliente
Dim objEnderecoBD As ClassEndereco
Dim objClienteBD As ClassCliente
Dim colEnderecosBD As colEndereco
Dim colEnderecoBD As Collection
Dim objFilialCliente As ClassFilialCliente
Dim colCategoriaItem As New Collection
Dim colEnderecoFilial As Collection
Dim lNumIntArq As Long
Dim bErroNoRegistro As Boolean
Dim bErroNoArquivo As Boolean
Dim objLogBase As ClassIntegracaoLog
Dim lCliente As Long
Dim iFilial As Integer
Dim bAchouCliente As Boolean
Dim lTransacaoExt As Long
Dim iIndice As Integer
Dim lTotalRegErr As Long
Dim lTotalRegCertos As Long
Dim objCidades As ClassCidades
Dim iCodigoCidade As Integer
Dim bExisteRegistro As Boolean
Dim objArqImportacao As ClassArqImportacao
Dim sParteCodigo As String
Dim bProcuraPeloCGC As Boolean
Dim bSemDados As Boolean
Dim bForca2 As Boolean
Dim sConteudo As String

On Error GoTo Erro_Importa_Dados_Cliente

    'Abertura de transação
    lTransacaoExt = Transacao_AbrirExt(glConexao1)
    If lTransacaoExt = 0 Then gError 189958

    lErro = Comando_ExecutarPos(lComandoExt1, "SELECT NumIntArq FROM ArqImportacao WHERE DataAtualizacao = ? AND TipoArq = ? ORDER BY NumIntArq", 0, lNumIntArq, DATA_NULA, objTipoArq.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 189940
    
    lErro = Comando_BuscarProximo(lComandoExt1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189941
    
    If lErro <> AD_SQL_SUCESSO Then
        bSemDados = True
    Else
        bSemDados = False
    End If
    
    lErro = CF("Config_Le", "FATConfig", "IMPORTAR_PV_APP_PPL", giFilialEmpresa, sConteudo)
    If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
    
    If gobjFAT.iImpDadosInterno = MARCADO Or StrParaInt(sConteudo) = MARCADO Then
        bForca2 = True 'pega também os que não tem arquivo texto NumIntArq = 0
    Else
        bForca2 = False
    End If
    Do While lErro <> AD_SQL_SEM_DADOS Or bForca2
    
        bForca2 = False
    
        bErroNoArquivo = False
        
        Set objIntegracaoCliente = New ClassIntegracaoCliente
    
        objIntegracaoCliente.lNumIntArq = lNumIntArq

        lErro = IntegracaoCliente_Le_ImpSemAtu(lComandoExt2, tIntegracaoCliente, objIntegracaoCliente)
        If lErro <> SUCESSO Then gError 189942
        
        lErro = IntegracaoCliente_Le_Proximo(lComandoExt2, tIntegracaoCliente, objIntegracaoCliente)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189943
        
        Do While lErro <> ERRO_LEITURA_SEM_DADOS
        
            Set colEnderecos = New colEndereco
            Set colEndereco = New Collection
            Set colEnderecosBD = New colEndereco
            Set colEnderecoBD = New Collection
            Set colEnderecoFilial = New Collection
            Set objCliente = New ClassCliente
            Set objClienteBD = New ClassCliente
            Set objFilialCliente = New ClassFilialCliente
            Set colCategoriaItem = New Collection

            bErroNoRegistro = False
            bExisteRegistro = False

            Set objLogBase = New ClassIntegracaoLog

            objLogBase.iEtapa = INT_ETAPA_IMP_ATUALI
            objLogBase.iTipoArq = objTipoArq.iCodigo
            objLogBase.sChave = "Cliente: " & CStr(objIntegracaoCliente.lCodCliente) & " Filial: " & CStr(objIntegracaoCliente.iCodFilial) & " CodExterno: " & CStr(objIntegracaoCliente.lCodExterno) & " CGC: " & objIntegracaoCliente.sCgc
            objLogBase.lNumIntDoc = objIntegracaoCliente.lNumIntDoc

            bAchouCliente = False
            
            'Se não veio o nosso código pesquisa pelo CodExterno e Pelo CGC
            If objIntegracaoCliente.lCodCliente = 0 Or objIntegracaoCliente.iCodFilial = 0 Then

                Call Log_Adiciona("O Cliente não possui o código no corporator.", objLogBase, True)

                If objIntegracaoCliente.lCodExterno <> 0 Then

                    lErro = Comando_Executar(lComando4, "SELECT CodCliente, CodFilial FROM FiliaisClientes WHERE CodExterno = ? ", lCliente, iFilial, objIntegracaoCliente.lCodExterno)
                    If lErro <> AD_SQL_SUCESSO Then gError 189944

                    lErro = Comando_BuscarProximo(lComando4)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189945

                    If lErro = AD_SQL_SUCESSO Then
                        objIntegracaoCliente.lCodCliente = lCliente
                        objIntegracaoCliente.iCodFilial = iFilial
                        bAchouCliente = True
                        Call Log_Adiciona("O Cliente foi localizado pelo código externo. " & "Cliente: " & CStr(objIntegracaoCliente.lCodCliente) & " Filial: " & CStr(objIntegracaoCliente.iCodFilial), objLogBase, True)
                    Else
                        Call Log_Adiciona("O código externo do cliente não foi localizado no corporator.", objLogBase, True)
                    End If

                End If
                
                bProcuraPeloCGC = True
                
                lErro = CF("Importa_Dados_Cliente_ProcuraCGC", bProcuraPeloCGC)
                If lErro <> SUCESSO Then gError 200264

                If Len(Trim(objIntegracaoCliente.sCgc)) <> 0 And Not bAchouCliente And bProcuraPeloCGC Then

                    lErro = Comando_Executar(lComando5, "SELECT CodCliente, CodFilial FROM FiliaisClientes WHERE CGC = ? ", lCliente, iFilial, objIntegracaoCliente.sCgc)
                    If lErro <> AD_SQL_SUCESSO Then gError 189946

                    lErro = Comando_BuscarProximo(lComando5)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189947

                    If lErro = AD_SQL_SUCESSO Then
                        objIntegracaoCliente.lCodCliente = lCliente
                        objIntegracaoCliente.iCodFilial = iFilial
                        bAchouCliente = True
                        Call Log_Adiciona("O Cliente foi localizado pelo CGC. " & "Cliente: " & CStr(objIntegracaoCliente.lCodCliente) & " Filial: " & CStr(objIntegracaoCliente.iCodFilial), objLogBase, True)
                    Else
                        Call Log_Adiciona("O CGC do cliente não foi localizado no corporator.", objLogBase, True)
                    End If

                End If
                
                'Se foi i,mportado do app do Lojista tenta localizar o cliente pelo email
                If objIntegracaoCliente.lNumIntGer = -1 And Len(Trim(objIntegracaoCliente.sEmail)) <> 0 Then
                
                    lErro = Comando_Executar(lComando7, "SELECT F.CodCliente, F.CodFilial FROM FiliaisClientes AS F, Enderecos As E WHERE E.Codigo = F.Endereco AND (E.Email = ? OR E.Email2 = ?)", lCliente, iFilial, objIntegracaoCliente.sEmail, objIntegracaoCliente.sEmail)
                    If lErro <> AD_SQL_SUCESSO Then gError 189944

                    lErro = Comando_BuscarProximo(lComando7)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189945

                    If lErro = AD_SQL_SUCESSO Then
                        objIntegracaoCliente.lCodCliente = lCliente
                        objIntegracaoCliente.iCodFilial = iFilial
                        bAchouCliente = True
                        Call Log_Adiciona("O Cliente foi localizado pelo Email. " & "Cliente: " & CStr(objIntegracaoCliente.lCodCliente) & " Filial: " & CStr(objIntegracaoCliente.iCodFilial), objLogBase, True)
                    Else
                        Call Log_Adiciona("O email do cliente não foi localizado no corporator.", objLogBase, True)
                    End If
                
                End If

            End If
            
            objCliente.lCodigo = objIntegracaoCliente.lCodCliente
            
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then
                bErroNoRegistro = True
                Call Log_Adiciona("Não foi possível ler o cliente.", objLogBase)
            End If

            objClienteBD.lCodigo = objIntegracaoCliente.lCodCliente
            
            lErro = CF("Cliente_Le", objClienteBD)
            If lErro <> SUCESSO And lErro <> 12293 Then
                bErroNoRegistro = True
                Call Log_Adiciona("Não foi possível ler o cliente.", objLogBase)
            End If

            If lErro = SUCESSO Then
            
                bExisteRegistro = True

                'Lê os dados dos tres tipos de enderecos
                lErro = CF("Enderecos_Le_Cliente", colEnderecos, objCliente)
                If lErro <> SUCESSO Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Não foi possível ler os endereços do cliente.", objLogBase)
                End If

                For Each objEndereco In colEnderecos
                    colEndereco.Add objEndereco
                Next
                
                'Lê os dados dos tres tipos de enderecos
                lErro = CF("Enderecos_Le_Cliente", colEnderecosBD, objClienteBD)
                If lErro <> SUCESSO Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Não foi possível ler os endereços do cliente.", objLogBase)
                End If

                For Each objEnderecoBD In colEnderecosBD
                    colEnderecoBD.Add objEnderecoBD
                Next

                'Se não encontrou o cliente verifica se o nome reduzido já existe,
                'se já existir concatena com o nome
                lErro = Comando_Executar(lComando6, "SELECT Codigo FROM Clientes WHERE NomeReduzido = ? AND Codigo <> ? ", lCliente, objIntegracaoCliente.sNomeReduzido, objIntegracaoCliente.lCodCliente)
                If lErro <> AD_SQL_SUCESSO Then gError 189944

                lErro = Comando_BuscarProximo(lComando6)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189945
            
                If lErro = AD_SQL_SUCESSO Then
                    If objIntegracaoCliente.lCodExterno <> 0 Then
                        sParteCodigo = "(" & CStr(objIntegracaoCliente.lCodExterno) & ")"
                    Else
                        sParteCodigo = "(" & CStr(objIntegracaoCliente.lCodCliente) & ")"
                    End If
                    objIntegracaoCliente.sNomeReduzido = left(objIntegracaoCliente.sNomeReduzido, STRING_CLIENTE_NOME_REDUZIDO - Len(sParteCodigo)) & sParteCodigo
                End If
                
            Else
                For iIndice = 1 To 3
                    Set objEndereco = New ClassEndereco
                    colEndereco.Add objEndereco
                Next
                Call Log_Adiciona("O código do cliente não foi localizado no corporator.", objLogBase, True)
             
                'Se não encontrou o cliente verifica se o nome reduzido já existe,
                'se já existir concatena com o nome
                lErro = Comando_Executar(lComando6, "SELECT Codigo FROM Clientes WHERE NomeReduzido = ? ", lCliente, objIntegracaoCliente.sNomeReduzido)
                If lErro <> AD_SQL_SUCESSO Then gError 189944

                lErro = Comando_BuscarProximo(lComando6)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189945
            
                If lErro = AD_SQL_SUCESSO Then
                    If objIntegracaoCliente.lCodExterno <> 0 Then
                        sParteCodigo = "(" & CStr(objIntegracaoCliente.lCodExterno) & ")"
                    Else
                        sParteCodigo = "(" & CStr(objIntegracaoCliente.lCodCliente) & ")"
                    End If
                    objIntegracaoCliente.sNomeReduzido = left(objIntegracaoCliente.sNomeReduzido, STRING_CLIENTE_NOME_REDUZIDO - Len(sParteCodigo)) & sParteCodigo
                End If
            
            End If

            lErro = Importa_Dados_Cliente_Preenche_Cliente(objIntegracaoCliente, objCliente, colEndereco)
            If lErro <> SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Não foi possível atualizar os objetos da cliente e seus endereços.", objLogBase)
            End If

            'verifica preenchimento de campos de acordo com configuracao
            lErro = CF("Cliente_Verifica_Preenchimento", objCliente)
            If lErro <> SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Informações incompletas para gravação do Cliente.", objLogBase)
            End If

            objFilialCliente.lCodCliente = objCliente.lCodigo
            objFilialCliente.iCodFilial = objIntegracaoCliente.iCodFilial

            'Lê a filial do cliente
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then
                bErroNoRegistro = True
                Call Log_Adiciona("Não foi possível ler a filial do cliente.", objLogBase)
            End If
            If lErro = SUCESSO Then

                'Lê as Categoria da Filial do Cliente
                lErro = CF("FilialCliente_Le_Categorias", objFilialCliente, colCategoriaItem)
                If lErro <> SUCESSO Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Não foi possível ler as categorias da filial do cliente.", objLogBase)
                End If

                Set objFilialCliente.colCategoriaItem = colCategoriaItem

                'Lê os dados dos tres tipos de enderecos
                lErro = CF("Enderecos_Le_FiliaisClientes", colEnderecoFilial, objFilialCliente)
                If lErro <> SUCESSO Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Não foi possível ler os endereços das filiais do cliente.", objLogBase)
                End If

            Else
                For iIndice = 1 To 3
                    Set objEndereco = New ClassEndereco
                    colEnderecoFilial.Add objEndereco
                Next
                Call Log_Adiciona("O código da filial cliente não foi localizado no corporator.", objLogBase, True)
            End If

            lErro = Importa_Dados_Cliente_Preenche_Filial(objIntegracaoCliente, objFilialCliente, colEnderecoFilial)
            If lErro <> SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Não foi possível atualizar os objetos da filial do cliente e seus endereços.", objLogBase)
            End If

            'Se o CGC estiver Preenchido
            If Len(Trim(objFilialCliente.sCgc)) > 0 And bProcuraPeloCGC Then
                'Verifica se tem outro Cliente com o mesmo CGC e dá aviso
                lErro = CF("FilialCliente_Testa_CGC", objFilialCliente.lCodCliente, 0, objFilialCliente.sCgc)
                If lErro <> SUCESSO Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("O CGC do cliente já existe.", objLogBase)
                End If
            End If

            'verifica preenchimento de campos de acordo com configuracao
            lErro = CF("FilialCliente_Verifica_Preenchimento", objFilialCliente)
            If lErro <> SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Informações incompletas para gravação da Filial do Cliente.", objLogBase)
            End If
            
            Set objArqImportacao = New ClassArqImportacao
            
            If giValidacaoManual = MARCADO And bExisteRegistro And (Not bErroNoRegistro) Then
                
                Load EscCliente
                lErro = EscCliente.Trata_Parametros(objArqImportacao, objClienteBD, colEnderecoBD.Item(1), objCliente, colEndereco.Item(1))
                If lErro <> SUCESSO Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Não conseguiu abrir a tela de comparação.", objLogBase)
                End If
                EscCliente.Show vbModal
                
            End If

            'Abertura de transação
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 189948
            
            If objArqImportacao.iManter <> MARCADO Then
            
                For iIndice = 1 To 3
                
                    Set objCidades = New ClassCidades
                    
                    objCidades.sDescricao = colEnderecoFilial.Item(iIndice).sCidade
                    
                    If Len(Trim(objCidades.sDescricao)) > 0 Then
                    
                        lErro = CF("Cidade_Le_Nome", objCidades)
                        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then
                            bErroNoRegistro = True
                            Call Log_Adiciona("Não foi possível ler a cidade do endereço.", objLogBase)
                        End If
                        
                        If lErro <> SUCESSO Then
                        
                            If Not bErroNoRegistro Then
                        
                                lErro = CF("Config_Obter_Inteiro_Automatico_EmTrans", "FATConfig", "NUM_PROX_CIDADECADASTRO", "Cidades", "Codigo", iCodigoCidade)
                                If lErro <> SUCESSO Then
                                    bErroNoRegistro = True
                                    Call Log_Adiciona("Não obter o próximo código de cidades.", objLogBase)
                                End If
                                
                            End If
                            
                            If Not bErroNoRegistro Then
                    
                                objCidades.iCodigo = iCodigoCidade
                                objCidades.sDescricao = colEnderecoFilial.Item(iIndice).sCidade
                                
                                lErro = CF("Cidade_Grava_EmTrans", objCidades)
                                If lErro <> SUCESSO Then
                                    bErroNoRegistro = True
                                    Call Log_Adiciona("Não foi possível gravar a cidade " & colEnderecoFilial.Item(iIndice).sCidade & ".", objLogBase)
                                End If
                                
                            End If
                            
                        End If
                    
                    End If
                
                Next
    
                If Not bErroNoRegistro Then
    
                    'Grava o Cliente no BD
                    lErro = CF("Cliente_Grava_EmTrans", objCliente, colEndereco)
                    If lErro <> SUCESSO Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Não foi possível gravar o Cliente.", objLogBase)
                    End If
    
                End If
    
                If Not bErroNoRegistro Then
    
                    lErro = CF("FiliaisClientes_Grava_EmTrans", objFilialCliente, colEnderecoFilial)
                    If lErro <> SUCESSO Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Não foi possível gravar a filial do Cliente.", objLogBase)
                    End If
    
                End If
    
                If Not bErroNoRegistro Then
    
                    lErro = Comando_Executar(lComando3, "UPDATE FiliaisClientes SET CodExterno = ? WHERE CodCliente = ? AND CodFilial = ? ", objIntegracaoCliente.lCodExterno, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
                    If lErro <> SUCESSO Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Não foi possível gravar o código externo na filial do Cliente.", objLogBase)
                    End If
    
                End If
                
            End If

            Call Log_Insere(objLogBase)

            If bErroNoRegistro Then

                bErroNoArquivo = True

                lTotalRegErr = lTotalRegErr + 1
                Call Transacao_Rollback

                'Abertura de transação só para marcar que houve erro
                lTransacao = Transacao_Abrir()
                If lTransacao = 0 Then gError 189949

                'Atualiza o registros
                lErro = Comando_Executar(lComando1, "UPDATE IntegracaoCliente SET ComErro = ? WHERE NumIntDoc = ? ", MARCADO, objIntegracaoCliente.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 189950

            Else

                lTotalRegCertos = lTotalRegCertos + 1

                'Atualiza o registros
                lErro = Comando_Executar(lComando1, "UPDATE IntegracaoCliente SET ComErro = ? , DataAtualizacao = ?, CodCliente = ?, CodFilial = ? WHERE NumIntDoc = ? ", DESMARCADO, Date, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, objIntegracaoCliente.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 189951

            End If

            'Confirma a transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 189952

            lErro = Rotina_Atualizar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError 189953

            lErro = IntegracaoCliente_Le_Proximo(lComandoExt2, tIntegracaoCliente, objIntegracaoCliente)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189954
        
        Loop
        
         'Não atualiza o Arquivo e não busca o próximo
        If bErroNoArquivo Then Exit Do
        
        If Not bSemDados Then
            lErro = ArqImportacao_Atualiza(lComandoExt3, lComandoExt1)
            If lErro <> SUCESSO Then gError 189955
            
            lErro = Comando_BuscarProximo(lComandoExt1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189956
        Else
            lErro = AD_SQL_SEM_DADOS
        End If
    Loop
    
    'Confirma a transação
    lErro = Transacao_CommitExt(lTransacaoExt)
    If lErro <> AD_SQL_SUCESSO Then gError 189959
    
    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Os Clientes foram atualizados. " & CStr(lTotalRegCertos) & " registros com sucesso e " & CStr(lTotalRegErr) & " registros com erro." & vbNewLine
    
    Importa_Dados_Cliente = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_Cliente:

    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Os Clientes foram atualizados. " & CStr(lTotalRegCertos) & " registros com sucesso e " & CStr(lTotalRegErr) & " registros com erro." & vbNewLine

    Importa_Dados_Cliente = gErr
     
    Select Case gErr
    
        Case 189940, 189941, 189956
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQIMPORTACAO", gErr)
            
        Case 189942, 189943, 189954
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INTEGRACAOCLIENTE", gErr)
            
        Case 189944 To 189947
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES2", gErr)

        Case 189948, 189949, 189958
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 189950, 189951
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_INTEGRACAOCLIENTE", gErr)
        
        Case 189952, 189959
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 189953 To 189955, 200264
        
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189957)
     
    End Select

    Call Transacao_Rollback
    Call Transacao_RollbackExt(lTransacaoExt)
  
    If gcolLog.Count > 0 Then Call Log_Insere(objLogBase)
  
    Exit Function

End Function

Private Function Importa_Dados_PV(alComando() As Long, ByVal lComandoExt1 As Long, ByVal lComandoExt2 As Long, ByVal lComandoExt3 As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal lComando8 As Long, ByVal lComando9 As Long, ByVal lComando10 As Long, ByVal lComando11 As Long, ByVal lComando12 As Long, ByVal lComando13 As Long, ByVal lComando14 As Long, ByVal lComando15 As Long, ByVal objTipoArq As ClassTipoArqIntegracao) As Long
        
Dim lErro As Long
Dim lSeq As Long
Dim lTransacao As Long
Dim objIntegracaoPV As ClassIntegracaoPV
Dim tIntegracaoPV As typeIntegracaoPV
Dim lNumIntArq As Long
Dim bErroNoRegistro As Boolean
Dim bErroNoArquivo As Boolean
Dim objLogBase As ClassIntegracaoLog
Dim lTransacaoExt As Long
Dim iIndice As Integer
Dim lTotalRegErr As Long
Dim lTotalRegCertos As Long
Dim objPV As ClassPedidoDeVenda
Dim objItemPV As ClassItemPedido
Dim lCodPedido As Long
Dim iFilialEmpresa As Integer
Dim iStatus As Integer
Dim lCodExternoAnt As Long
Dim bForca As Boolean
Dim bForca2 As Boolean
Dim colIntegracao As New Collection
Dim bCancelado As Boolean
Dim objIntegracaoPVAux As ClassIntegracaoPV
Dim objTelaPedido As Object
Dim objTelaPedidoOCX As Object
Dim bLoad As Boolean
Dim iCancel As Integer
Dim bPedidoNovo As Boolean
Dim bSemDados As Boolean
Dim dtVencimento As Date, dValorParc  As Double
Dim sConteudo As String

On Error GoTo Erro_Importa_Dados_PV

    Set objTelaPedidoOCX = CreateObject("TelasFAT.PedidoVendaOCX")
    Set objTelaPedido = objTelaPedidoOCX.objCTtela

    bLoad = False
    Call objTelaPedidoOCX.Form_Load
    bLoad = True
    Call objTelaPedido.Form_Activate
    
    objTelaPedido.iImportacao = MARCADO

    'Abertura de transação
    lTransacaoExt = Transacao_AbrirExt(glConexao1)
    If lTransacaoExt = 0 Then gError 190068

    lErro = Comando_ExecutarPos(lComandoExt1, "SELECT NumIntArq FROM ArqImportacao WHERE DataAtualizacao = ? AND TipoArq = ? ORDER BY NumIntArq", 0, lNumIntArq, DATA_NULA, objTipoArq.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 190069
    
    lErro = Comando_BuscarProximo(lComandoExt1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190070
    
    If lErro <> AD_SQL_SUCESSO Then
        bSemDados = True
    Else
        bSemDados = False
    End If
    
    lErro = CF("Config_Le", "FATConfig", "IMPORTAR_PV_APP_PPL", giFilialEmpresa, sConteudo)
    If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
    
    If gobjFAT.iImpDadosInterno = MARCADO Or StrParaInt(sConteudo) = MARCADO Then
        bForca2 = True 'pega também os que não tem arquivo texto NumIntArq = 0
    Else
        bForca2 = False
    End If
    Do While lErro <> AD_SQL_SEM_DADOS Or bForca2
    
        bForca2 = False
    
        bErroNoArquivo = False
        
        Set objIntegracaoPV = New ClassIntegracaoPV
    
        objIntegracaoPV.lNumIntArq = lNumIntArq

        lErro = IntegracaoPV_Le_ImpSemAtu(lComandoExt2, tIntegracaoPV, objIntegracaoPV)
        If lErro <> SUCESSO Then gError 190071
        
        lErro = IntegracaoPV_Le_Proximo(lComandoExt2, tIntegracaoPV, objIntegracaoPV)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190072
        
        lCodExternoAnt = 0
        
        bForca = False
        Do While lErro <> ERRO_LEITURA_SEM_DADOS Or bForca
        
            Set objIntegracaoPV = objIntegracaoPV
        
            If objIntegracaoPV.lCodPedidoExterno <> lCodExternoAnt Then
            
                If Not (objPV Is Nothing) Then
                
                    lErro = Comando_Executar(alComando(1), "SELECT Vencimento, Valor FROM IntegracaoPVParc WHERE NumIntArq = ? AND Pedido = ? AND FilialEmpresa = ? ORDER BY Parcela", dtVencimento, dValorParc, lNumIntArq, lCodExternoAnt, giFilialEmpresa)
                    If lErro <> AD_SQL_SUCESSO Then gError 190079
        
                    lErro = Comando_BuscarProximo(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190080
                    
                    Do While lErro <> AD_SQL_SEM_DADOS
                    
                        objPV.colParcelas.Add dValorParc, dtVencimento, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, DATA_NULA, 0, 0
                    
                        lErro = Comando_BuscarProximo(alComando(1))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190080
                    
                    Loop
                
                    lErro = Importa_Dados_PV_Preenche_PV(lComando4, lComando5, lComando6, lComando7, lComando8, lComando9, lComando10, lComando11, lComando12, lComando13, colIntegracao, objPV, objLogBase, bErroNoRegistro, bCancelado, objTelaPedido, objTelaPedidoOCX, lComando14, lComando15)
                    If lErro <> SUCESSO Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Não foi possível preencher o objeto do pedido de venda.", objLogBase)
                    End If
                
                    'Abertura de transação
                    lTransacao = Transacao_Abrir()
                    If lTransacao = 0 Then gError 190073
                    
                    If Not bErroNoRegistro And bPedidoNovo Then
                        lErro = CF("PedidoVenda_Grava_EmTrans", objPV)
                        If lErro <> SUCESSO Then
                            bErroNoRegistro = True
                            Call Log_Adiciona("Não foi possível gravar o pedido de venda.", objLogBase)
                        End If
                    End If
                    
                    If Not bErroNoRegistro And bCancelado Then
                        
                        For Each objItemPV In objPV.colItensPedido
                            objItemPV.dQuantCancelada = objItemPV.dQuantidade
                            objItemPV.dQuantReservada = 0
                            Set objItemPV.ColReserva = New ColReserva
                        Next
                    
                        lErro = CF("PedidoVenda_Grava_EmTrans", objPV)
                        If lErro <> SUCESSO Then
                            bErroNoRegistro = True
                            Call Log_Adiciona("Não foi possível gravar o pedido de venda.", objLogBase)
                        End If
                    End If
                    
'                    If Not bErroNoRegistro Then
'                        If bCancelado Then
'                            'baixa o pedido
'                            lErro = CF("PedidoDeVenda_Baixar", alComando, objPV)
'                            If lErro <> SUCESSO Then
'                                bErroNoRegistro = True
'                                Call Log_Adiciona("Não foi possível baixar o pedido de venda.", objLogBase)
'                            End If
'                        End If
'                    End If
        
                    If Not bErroNoRegistro Then
        
                        lErro = Comando_Executar(lComando3, "UPDATE PedidosDeVenda SET CodExterno = ? WHERE Codigo = ? AND FilialEmpresa = ? ", colIntegracao.Item(1).lCodPedidoExterno, objPV.lCodigo, objPV.iFilialEmpresa)
                        If lErro <> SUCESSO Then
                            bErroNoRegistro = True
                            Call Log_Adiciona("Não foi possível gravar o código externo na filial do Cliente.", objLogBase)
                        End If
        
                    End If
        
                    Call Log_Insere(objLogBase)
        
                    If bErroNoRegistro Then
        
                        bErroNoArquivo = True
        
                        For Each objIntegracaoPVAux In colIntegracao
                            lTotalRegErr = lTotalRegErr + 1
                        Next
                        
                        Call Transacao_Rollback
        
                        'Abertura de transação só para marcar que houve erro
                        lTransacao = Transacao_Abrir()
                        If lTransacao = 0 Then gError 190074
        
                        'Atualiza o registros
                        For Each objIntegracaoPVAux In colIntegracao
                            lErro = Comando_Executar(lComando1, "UPDATE IntegracaoPV SET ComErro = ? WHERE NumIntDoc = ? ", MARCADO, objIntegracaoPVAux.lNumIntDoc)
                            If lErro <> AD_SQL_SUCESSO Then gError 190075
                        Next
        
                    Else
        
                        For Each objIntegracaoPVAux In colIntegracao
                            lTotalRegCertos = lTotalRegCertos + 1
                        Next
        
                        'Atualiza o registros
                        For Each objIntegracaoPVAux In colIntegracao
                            lErro = Comando_Executar(lComando1, "UPDATE IntegracaoPV SET ComErro = ? , DataAtualizacao = ? WHERE NumIntDoc = ? ", DESMARCADO, Date, objIntegracaoPVAux.lNumIntDoc)
                            If lErro <> AD_SQL_SUCESSO Then gError 190076
                        Next
        
                    End If
        
                    'Confirma a transação
                    lErro = Transacao_Commit()
                    If lErro <> AD_SQL_SUCESSO Then gError 190077
        
                    For Each objIntegracaoPVAux In colIntegracao
                        lErro = Rotina_Atualizar_AtualizaTelaBatch
                        If lErro <> SUCESSO Then gError 190078
                    Next

                End If
            
                If objIntegracaoPV.lNumIntDoc <> 0 Then  'Se só rodou para gravar a coleção anterior não tem mais dados novos em objIntegracaoPV
            
                    Set colIntegracao = New Collection
                
                    lCodExternoAnt = objIntegracaoPV.lCodPedidoExterno
                    
                    bErroNoRegistro = False
        
                    Set objLogBase = New ClassIntegracaoLog
        
                    objLogBase.iEtapa = INT_ETAPA_IMP_ATUALI
                    objLogBase.iTipoArq = objTipoArq.iCodigo
                    objLogBase.sChave = "Pedido código: " & CStr(objIntegracaoPV.lCodPedidoExterno)
                    objLogBase.lNumIntDoc = objIntegracaoPV.lNumIntDoc
                    
                    Set objPV = New ClassPedidoDeVenda
                
                    lErro = Comando_Executar(lComando2, "SELECT Codigo, FilialEmpresa FROM PedVenTodos WHERE PedidoCliente = ? ", lCodPedido, iFilialEmpresa, CStr(objIntegracaoPV.lCodPedidoExterno))
                    If lErro <> AD_SQL_SUCESSO Then gError 190079
        
                    lErro = Comando_BuscarProximo(lComando2)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190080
        
                    If lErro = AD_SQL_SUCESSO Then
                        
                        objPV.lCodigo = lCodPedido
                        objPV.iFilialEmpresa = iFilialEmpresa
                        
                        bPedidoNovo = False
                        
                        Call Log_Adiciona("O pedido foi localizado pelo código externo. " & "Pedido: " & CStr(objPV.lCodigo) & " Filial Empresa: " & CStr(objPV.iFilialEmpresa), objLogBase, True)
                    Else
                    
                        bPedidoNovo = True
                        
                        Call Log_Adiciona("O código externo do pedido não foi localizado no corporator.", objLogBase, True)
                    End If
                    
                End If
                        
            End If
            
            If bForca Then Exit Do
            
            colIntegracao.Add objIntegracaoPV
        
            lErro = IntegracaoPV_Le_Proximo(lComandoExt2, tIntegracaoPV, objIntegracaoPV)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190081
            
            If lErro = ERRO_LEITURA_SEM_DADOS Then bForca = True
                    
        Loop
        
         'Não atualiza o Arquivo e não busca o próximo
        If bErroNoArquivo Then Exit Do
        
        If Not bSemDados Then
            lErro = ArqImportacao_Atualiza(lComandoExt3, lComandoExt1)
            If lErro <> SUCESSO Then gError 190082
        
            lErro = Comando_BuscarProximo(lComandoExt1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190083
        Else
            lErro = AD_SQL_SEM_DADOS
        End If

    Loop
    
    'Confirma a transação
    lErro = Transacao_CommitExt(lTransacaoExt)
    If lErro <> AD_SQL_SUCESSO Then gError 190084
    
    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Os pedidos de venda foram atualizados. " & CStr(lTotalRegCertos) & " registros com sucesso e " & CStr(lTotalRegErr) & " registros com erro." & vbNewLine

    'Call objTelaPedidoOCX.Form_Deactivate

    Call objTelaPedidoOCX.Form_Unload(iCancel)
    bLoad = False

    Set objTelaPedidoOCX = Nothing
    Set objTelaPedido = Nothing

    Importa_Dados_PV = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_PV:

    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Os pedidos de venda foram atualizados. " & CStr(lTotalRegCertos) & " registros com sucesso e " & CStr(lTotalRegErr) & " registros com erro." & vbNewLine

    Importa_Dados_PV = gErr
     
    Select Case gErr
    
        Case 190068, 190073, 190074
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 190069, 190070, 190083
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQIMPORTACAO", gErr)
            
        Case 190071, 190072, 190081
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INTEGRACAOPV", gErr)
            
        Case 190075, 190076
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_INTEGRACAOPV", gErr)
        
        Case 190077, 190084
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 190078, 190082
        
        Case 190079, 190080
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", gErr)
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190085)
     
    End Select

    Call Transacao_Rollback
    Call Transacao_RollbackExt(lTransacaoExt)
  
    If gcolLog.Count > 0 Then Call Log_Insere(objLogBase)
    
    If bLoad Then Call objTelaPedidoOCX.Form_Unload
    
    Set objTelaPedidoOCX = Nothing
    Set objTelaPedido = Nothing
  
    Exit Function

End Function

Public Function Exporta_Dados(ByVal objArqExportacaoAux As ClassArqExportacaoAux) As Long
        
Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 6) As Long
Dim alComandoExt(1 To 2) As Long
Dim objTipoArq As ClassTipoArqIntegracao
Dim sNomeArq As String
Dim sDiretorio As String
Dim lRetorno As Long

On Error GoTo Erro_Exporta_Dados

    lErro = Abre_Conexoes
    If lErro <> SUCESSO Then gError 189810

    'abertura de comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 189811
    Next
    
    'abertura de comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
        If alComandoExt(iIndice) = 0 Then gError 189812
    Next
    
    'Obtém o diretório onde estão os arquivos
    sDiretorio = String(512, 0)
    lRetorno = GetPrivateProfileString("Geral", "dirArqExportTemp", "c:\", sDiretorio, 512, "ADM100.INI")
    sDiretorio = left(sDiretorio, lRetorno)
    
    gsDiretorioExportTemp = sDiretorio
    gsDiretorioExport = objArqExportacaoAux.sDiretorio
    
    lErro = ArqExportacaoAux_Insere(alComando(1), objArqExportacaoAux)
    If lErro <> SUCESSO Then gError 189813
    
    'Gera as informações da tabela temporária
    For Each objTipoArq In objArqExportacaoAux.colTiposArq
        
        Select Case objTipoArq.iCodigo
        
            Case TIPO_ARQ_CLIENTE
            
                lErro = Exporta_Dados_Cliente(alComandoExt(1), alComandoExt(2), alComando(2), alComando(3), objArqExportacaoAux)
                If lErro <> SUCESSO Then gError 189814
            
            Case TIPO_ARQ_SLDPROD
            
                lErro = Exporta_Dados_SldProd(alComandoExt(1), alComandoExt(2), alComando(2), alComando(3), objArqExportacaoAux)
                If lErro <> SUCESSO Then gError 190091
            
            Case Else
                gError 189815
            
        End Select
    
    Next
    
    lErro = GeraArq_Inicia(objArqExportacaoAux.colTiposArq.Count)
    If lErro <> SUCESSO Then gError 189855
    
    'Gera os arquivos
    For Each objTipoArq In objArqExportacaoAux.colTiposArq
    
        Call Exportacao_Obtem_NomeArq(objTipoArq, sNomeArq)
        
        TelaAcompanhaBatch.NomeArqAtu = sNomeArq
        
        DoEvents
    
        Select Case objTipoArq.iCodigo
        
            Case TIPO_ARQ_CLIENTE
            
                lErro = Exporta_Dados_Cliente_Arq(alComando(4), alComando(5), alComando(6), sNomeArq)
                If lErro <> SUCESSO Then gError 189848
            
            Case TIPO_ARQ_SLDPROD
            
                lErro = Exporta_Dados_SldProd_Arq(alComando(4), alComando(5), alComando(6), sNomeArq)
                If lErro <> SUCESSO Then gError 190092
            
            Case Else
                gError 189849
            
        End Select
    
        'Atualiza a barra de progressão
        lErro = Rotina_Atualizar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 189856
    
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Fecha_Conexoes
    
    Exporta_Dados = SUCESSO
     
    Exit Function
    
Erro_Exporta_Dados:

    Exporta_Dados = gErr
     
    Select Case gErr
    
        Case 189810, 189813, 189814, 189848, 189855, 189856, 190091, 190092
    
        Case 189811, 189812
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 189815, 189849
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOARQINTEGRACAO_NAO_TRATADO", gErr, objTipoArq.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189816)
     
    End Select

    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Fecha_Conexoes
       
    Exit Function

End Function

Private Function Exporta_Dados_Cliente(ByVal lComandoExt1 As Long, ByVal lComandoExt2 As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objArqExportacaoAux As ClassArqExportacaoAux) As Long
        
Dim lErro As Long
Dim lTotalRegCertos As Long
Dim lTotalRegErr As Long
Dim lCliente As Long
Dim iFilial As Long
Dim lCodExterno As Long
Dim objIntegracaoCliente As ClassIntegracaoCliente
Dim lSeq As Long
Dim lTransacao As Long
Dim lTotalRegistros As Long
Dim objLogBase As ClassIntegracaoLog
Dim bErro As Boolean
Dim bErroNoRegistro As Boolean

On Error GoTo Erro_Exporta_Dados_Cliente

    bErro = False
   
    lErro = Exporta_Dados_Cliente_Le(lComandoExt1, lComandoExt2, objArqExportacaoAux, lCliente, iFilial, lTotalRegistros, lCodExterno)
    If lErro <> SUCESSO Then gError 189817
    
    lErro = Exportacao_Inicia("Clientes", lTotalRegistros)
    If lErro <> SUCESSO Then gError 189818
    
    lErro = Comando_BuscarPrimeiro(lComandoExt1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189819
    
    lSeq = 0
    Do While lErro <> AD_SQL_SEM_DADOS
    
        bErroNoRegistro = False
        
        Set objLogBase = New ClassIntegracaoLog
        
        objLogBase.iEtapa = INT_ETAPA_EXP_TABINT
        objLogBase.iTipoArq = TIPO_ARQ_CLIENTE
        objLogBase.sChave = "Cliente: " & CStr(lCliente) & " Filial: " & CStr(iFilial)

        lSeq = lSeq + 1
        Set objIntegracaoCliente = New ClassIntegracaoCliente
        
        lErro = Exporta_Dados_Cliente_Preenche_Obj(lCliente, iFilial, objIntegracaoCliente)
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Erro na leitura dos dados da filial do cliente", objLogBase)
        End If
        
        objIntegracaoCliente.lCodExterno = lCodExterno
        
        objIntegracaoCliente.lSeqRegistro = lSeq
        objIntegracaoCliente.lNumIntGer = objArqExportacaoAux.lNumIntGeracao
        objIntegracaoCliente.iTipoInt = TIPO_INTEGRACAO_EXPORTACAO
        
        'Abertura de transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 189820
        
        lErro = IntegracaoCliente_Insere(lComando1, objIntegracaoCliente)
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Não foi possível inserir as informações do cliente na tabela temporária.", objLogBase)
        End If
        
        Call Log_Insere(objLogBase)
        
        If bErroNoRegistro Then
        
            lTotalRegErr = lTotalRegErr + 1
            Call Transacao_Rollback
            
        Else
        
            lTotalRegCertos = lTotalRegCertos + 1
        
            lErro = Comando_Executar(lComando2, "UPDATE FiliaisClientes SET Exportado = ? WHERE CodCliente = ? AND CodFilial = ? ", MARCADO, lCliente, iFilial)
            If lErro <> AD_SQL_SUCESSO Then gError 189821

            'Confirma a transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 189822

        End If
        
        'Atualiza a barra de progressão
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 189823
    
        lErro = Comando_BuscarProximo(lComandoExt1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189824

    Loop
    
    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Os Clientes foram exportados. " & CStr(lTotalRegCertos) & " registros com sucesso e " & CStr(lTotalRegErr) & " registros com erro." & vbNewLine
    
    Exporta_Dados_Cliente = SUCESSO
     
    Exit Function
    
Erro_Exporta_Dados_Cliente:

    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Os Clientes foram exportados. " & CStr(lTotalRegCertos) & " registros com sucesso e " & CStr(lTotalRegErr) & " registros com erro." & vbNewLine

    Exporta_Dados_Cliente = gErr
     
    Select Case gErr
    
        Case 189817, 189818, 189823
        
        Case 189819, 189824
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES2", gErr)

        Case 189820
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 189821
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_FILIAISCLIENTES", gErr)
            
        Case 189822
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189825)
     
    End Select

    Call Transacao_Rollback
  
    If gcolLog.Count > 0 Then Call Log_Insere(objLogBase)
  
    Exit Function

End Function

Private Function Exporta_Dados_SldProd(ByVal lComandoExt1 As Long, ByVal lComandoExt2 As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objArqExportacaoAux As ClassArqExportacaoAux) As Long
        
Dim lErro As Long
Dim lTotalRegCertos As Long
Dim lTotalRegErr As Long
Dim sProduto As String
Dim iAlmoxarifado As Long
Dim objIntegracaoSldProd As ClassIntegracaoSldProd
Dim lSeq As Long
Dim lTransacao As Long
Dim lTotalRegistros As Long
Dim objLogBase As ClassIntegracaoLog
Dim bErro As Boolean
Dim bErroNoRegistro As Boolean
Dim dSaldo As Double
Dim bAlmValido As Boolean

On Error GoTo Erro_Exporta_Dados_SldProd

    bErro = False
   
    lErro = Exporta_Dados_SldProd_Le(lComandoExt1, lComandoExt2, objArqExportacaoAux, sProduto, iAlmoxarifado, dSaldo, lTotalRegistros)
    If lErro <> SUCESSO Then gError 190093
    
    lErro = Exportacao_Inicia("Saldo Disponível", lTotalRegistros)
    If lErro <> SUCESSO Then gError 190094
    
    lErro = Comando_BuscarPrimeiro(lComandoExt1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190095
    
    lSeq = 0
    Do While lErro <> AD_SQL_SEM_DADOS
    
        bErroNoRegistro = False
        
        Set objLogBase = New ClassIntegracaoLog
        
        objLogBase.iEtapa = INT_ETAPA_EXP_TABINT
        objLogBase.iTipoArq = TIPO_ARQ_SLDPROD
        objLogBase.sChave = "Produto: " & sProduto & " Almoxarifado: " & CStr(iAlmoxarifado)

        bAlmValido = True
        lErro = CF("Exporta_Dados_SldProd_AlmValido_Cust", iAlmoxarifado, bAlmValido)
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Erro na verificação do almoxarifado.", objLogBase)
        End If
        
        If bAlmValido Then

            lSeq = lSeq + 1
            Set objIntegracaoSldProd = New ClassIntegracaoSldProd
            
            lErro = Exporta_Dados_SldProd_Preenche_Obj(sProduto, iAlmoxarifado, dSaldo, objIntegracaoSldProd)
            If lErro <> SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Erro no preenchimento do objeto de exportação do saldo.", objLogBase)
            End If
            
            objIntegracaoSldProd.lSeqRegistro = lSeq
            objIntegracaoSldProd.lNumIntGer = objArqExportacaoAux.lNumIntGeracao
            objIntegracaoSldProd.iTipoInt = TIPO_INTEGRACAO_EXPORTACAO
            
            'Abertura de transação
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 190096
            
            lErro = IntegracaoSldProd_Insere(lComando1, objIntegracaoSldProd)
            If lErro <> SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Não foi possível inserir as informações do saldo disponível na tabela temporária.", objLogBase)
            End If
            
            Call Log_Insere(objLogBase)
            
            If bErroNoRegistro Then
            
                lTotalRegErr = lTotalRegErr + 1
                Call Transacao_Rollback
                
            Else
            
                lTotalRegCertos = lTotalRegCertos + 1
            
                lErro = Comando_Executar(lComando2, "UPDATE EstoqueProduto SET Exportado = ? WHERE Produto = ? AND Almoxarifado = ? ", MARCADO, sProduto, iAlmoxarifado)
                If lErro <> AD_SQL_SUCESSO Then gError 190097
    
                'Confirma a transação
                lErro = Transacao_Commit()
                If lErro <> AD_SQL_SUCESSO Then gError 190098
    
            End If
            
        End If
        
        'Atualiza a barra de progressão
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 190099
    
        lErro = Comando_BuscarProximo(lComandoExt1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190100

    Loop
    
    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Os saldos foram exportados. " & CStr(lTotalRegCertos) & " registros com sucesso e " & CStr(lTotalRegErr) & " registros com erro." & vbNewLine
    
    Exporta_Dados_SldProd = SUCESSO
     
    Exit Function
    
Erro_Exporta_Dados_SldProd:

    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Os saldos foram exportados. " & CStr(lTotalRegCertos) & " registros com sucesso e " & CStr(lTotalRegErr) & " registros com erro." & vbNewLine

    Exporta_Dados_SldProd = gErr
     
    Select Case gErr
    
        Case 190093, 190094, 190099
        
        Case 190095, 190100
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", gErr)

        Case 190096
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 190097
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_ESTOQUEPRODUTO", gErr)
            
        Case 190098
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190101)
     
    End Select

    Call Transacao_Rollback
  
    If gcolLog.Count > 0 Then Call Log_Insere(objLogBase)
  
    Exit Function

End Function

Private Function Exporta_Dados_Cliente_Arq(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal sNomeArq As String) As Long
        
Dim lErro As Long
Dim tIntegracaoCliente As typeIntegracaoCliente
Dim objIntegracaoCliente As ClassIntegracaoCliente
Dim lTransacao As Long
Dim objExportacao As New ClassArqExportacao
Dim sRegistro As String
Dim bArqAberto As Boolean
Dim objLogBase As ClassIntegracaoLog
Dim bErroRegistro As Boolean
Dim lTotalRegistros As Long

On Error GoTo Erro_Exporta_Dados_Cliente_Arq

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 189857
    
    objExportacao.iTipoArq = TIPO_ARQ_CLIENTE
    objExportacao.sNomeArquivo = sNomeArq
    
    lErro = ArqExportacao_Insere(lComando3, objExportacao)
    If lErro <> SUCESSO Then gError 189858

    lErro = IntegracaoCliente_Le_SemArq(lComando1, tIntegracaoCliente, objIntegracaoCliente)
    If lErro <> SUCESSO Then gError 189859
    
    lErro = IntegracaoCliente_Le_Proximo(lComando1, tIntegracaoCliente, objIntegracaoCliente)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189860
    
    Open gsDiretorioExportTemp & sNomeArq For Output As #1
    bArqAberto = True
    
    lTotalRegistros = 0
    
    Do While lErro <> ERRO_LEITURA_SEM_DADOS
    
        lTotalRegistros = lTotalRegistros + 1
    
        bErroRegistro = False
    
        Set objLogBase = New ClassIntegracaoLog
        
        objLogBase.iEtapa = INT_ETAPA_EXP_GERARQ
        objLogBase.iTipoArq = TIPO_ARQ_CLIENTE
        objLogBase.lNumIntDoc = objIntegracaoCliente.lNumIntDoc
        objLogBase.sChave = "Cliente: " & CStr(objIntegracaoCliente.lCodCliente) & " Filial: " & CStr(objIntegracaoCliente.iCodFilial)
                
        lErro = objIntegracaoCliente.Monta_Registro_Cliente(sRegistro)
        If lErro <> SUCESSO Then
            bErroRegistro = True
            Call Log_Adiciona("Não foi possível montar o registro.", objLogBase)
        End If
        
        Call Log_Insere(objLogBase)
        
        If Not bErroRegistro Then
        
            Print #1, sRegistro
        
            objIntegracaoCliente.lNumIntArq = objExportacao.lNumIntArq
            objIntegracaoCliente.dtDataAtualizacao = Date
            objIntegracaoCliente.iComErro = DESMARCADO
            
        Else
            
            objIntegracaoCliente.lNumIntArq = 0
            objIntegracaoCliente.iComErro = MARCADO
            objIntegracaoCliente.dtDataAtualizacao = DATA_NULA
            
        End If
    
        lErro = IntegracaoCliente_Atualiza(lComando2, lComando1, objIntegracaoCliente)
        If lErro <> SUCESSO Then gError 189861
    
        lErro = IntegracaoCliente_Le_Proximo(lComando1, tIntegracaoCliente, objIntegracaoCliente)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189862
    
    Loop
    
    Print #1, "99" & FormataCpoNum(lTotalRegistros, 6)
    
    Close #1
    bArqAberto = False
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 189863
    
    FileCopy gsDiretorioExportTemp & sNomeArq, gsDiretorioExport & sNomeArq
    
    Exporta_Dados_Cliente_Arq = SUCESSO
     
    Exit Function
    
Erro_Exporta_Dados_Cliente_Arq:

    Exporta_Dados_Cliente_Arq = gErr
     
    Select Case gErr

        Case 189857
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 189858 To 189862
            
        Case 189863
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189864)
     
    End Select
    
    'Desfaz Transação
    Call Transacao_Rollback
    
    If bArqAberto Then Close #1
    
    If gcolLog.Count > 0 Then Call Log_Insere(objLogBase)
    
    Exit Function

End Function

Private Function Exporta_Dados_SldProd_Arq(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal sNomeArq As String) As Long
        
Dim lErro As Long
Dim tIntegracaoSldProd As typeIntegracaoSldProd
Dim objIntegracaoSldProd As ClassIntegracaoSldProd
Dim lTransacao As Long
Dim objExportacao As New ClassArqExportacao
Dim sRegistro As String
Dim bArqAberto As Boolean
Dim objLogBase As ClassIntegracaoLog
Dim bErroRegistro As Boolean
Dim lTotalRegistros As Long

On Error GoTo Erro_Exporta_Dados_SldProd_Arq

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190126
    
    objExportacao.iTipoArq = TIPO_ARQ_SLDPROD
    objExportacao.sNomeArquivo = sNomeArq
    
    lErro = ArqExportacao_Insere(lComando3, objExportacao)
    If lErro <> SUCESSO Then gError 190127

    lErro = IntegracaoSldProd_Le_SemArq(lComando1, tIntegracaoSldProd, objIntegracaoSldProd)
    If lErro <> SUCESSO Then gError 190128
    
    lErro = IntegracaoSldProd_Le_Proximo(lComando1, tIntegracaoSldProd, objIntegracaoSldProd)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190129
    
    Open gsDiretorioExportTemp & sNomeArq For Output As #1
    bArqAberto = True
    
    lTotalRegistros = 0
    
    Do While lErro <> ERRO_LEITURA_SEM_DADOS
    
        lTotalRegistros = lTotalRegistros + 1
    
        bErroRegistro = False
    
        Set objLogBase = New ClassIntegracaoLog
        
        objLogBase.iEtapa = INT_ETAPA_EXP_GERARQ
        objLogBase.iTipoArq = TIPO_ARQ_SLDPROD
        objLogBase.lNumIntDoc = objIntegracaoSldProd.lNumIntDoc
        objLogBase.sChave = "Produto: " & CStr(objIntegracaoSldProd.sCodProduto) & " Almoxarifado: " & CStr(objIntegracaoSldProd.iAlmoxarifado)
                
        lErro = objIntegracaoSldProd.Monta_Registro_SldProd(sRegistro)
        If lErro <> SUCESSO Then
            bErroRegistro = True
            Call Log_Adiciona("Não foi possível montar o registro.", objLogBase)
        End If
        
        Call Log_Insere(objLogBase)
        
        If Not bErroRegistro Then
        
            Print #1, sRegistro
        
            objIntegracaoSldProd.lNumIntArq = objExportacao.lNumIntArq
            objIntegracaoSldProd.dtDataAtualizacao = Date
            objIntegracaoSldProd.iComErro = DESMARCADO
            
        Else
            
            objIntegracaoSldProd.lNumIntArq = 0
            objIntegracaoSldProd.iComErro = MARCADO
            objIntegracaoSldProd.dtDataAtualizacao = DATA_NULA
            
        End If
    
        lErro = IntegracaoSldProd_Atualiza(lComando2, lComando1, objIntegracaoSldProd)
        If lErro <> SUCESSO Then gError 190130
    
        lErro = IntegracaoSldProd_Le_Proximo(lComando1, tIntegracaoSldProd, objIntegracaoSldProd)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190131
    
    Loop
    
    Print #1, "99" & FormataCpoNum(lTotalRegistros, 6)
    
    Close #1
    bArqAberto = False
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190132
    
    FileCopy gsDiretorioExportTemp & sNomeArq, gsDiretorioExport & sNomeArq
    
    Exporta_Dados_SldProd_Arq = SUCESSO
     
    Exit Function
    
Erro_Exporta_Dados_SldProd_Arq:

    Exporta_Dados_SldProd_Arq = gErr
     
    Select Case gErr

        Case 190126
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 190127 To 190131
            
        Case 190132
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190133)
     
    End Select
    
    'Desfaz Transação
    Call Transacao_Rollback
    
    If bArqAberto Then Close #1
    
    If gcolLog.Count > 0 Then Call Log_Insere(objLogBase)
    
    Exit Function

End Function

Private Function Exporta_Dados_Cliente_Preenche_Obj(ByVal lCliente As Long, ByVal iFilial As Integer, ByVal objIntegracaoCliente As ClassIntegracaoCliente) As Long
        
Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFilial As New ClassFilialCliente
Dim objEndereco As New ClassEndereco
Dim objEnderecoEnt As New ClassEndereco
Dim objEnderecoCobr As New ClassEndereco

On Error GoTo Erro_Exporta_Dados_Cliente_Preenche_Obj

    Set objFilial = New ClassFilialCliente

    objFilial.lCodCliente = lCliente
    objFilial.iCodFilial = iFilial
    
    'Lê o produto
    lErro = CF("FilialCliente_Le", objFilial)
    If lErro <> SUCESSO And lErro <> 12567 Then gError 189826
        
    Set objCliente = New ClassCliente
    
    objCliente.lCodigo = objFilial.lCodCliente
    
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 189827
    
    Set objEndereco = New ClassEndereco

    objEndereco.lCodigo = objFilial.lEndereco

    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 189828

    'Endereco de Entrega
    Set objEnderecoEnt = New ClassEndereco

    objEnderecoEnt.lCodigo = objFilial.lEnderecoEntrega

    lErro = CF("Endereco_Le", objEnderecoEnt)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 189829

    'Endereco de Cobranca
    Set objEnderecoCobr = New ClassEndereco

    objEnderecoCobr.lCodigo = objFilial.lEnderecoCobranca

    lErro = CF("Endereco_Le", objEnderecoCobr)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 189830

    objIntegracaoCliente.iTipoInt = ROTINA_EXPORTACAO_DADOS
    objIntegracaoCliente.lCodCliente = objFilial.lCodCliente
    objIntegracaoCliente.iCodFilial = objFilial.iCodFilial
    objIntegracaoCliente.sRazaoSocial = objCliente.sRazaoSocial
    objIntegracaoCliente.sNomeReduzido = objCliente.sNomeReduzido
    objIntegracaoCliente.iTipo = objCliente.iTipo
    objIntegracaoCliente.sObservacao = objCliente.sObservacao
    objIntegracaoCliente.dLimiteCredito = objCliente.dLimiteCredito
    objIntegracaoCliente.iCondicaoPagto = objCliente.iCondicaoPagto
    objIntegracaoCliente.iAtivo = objFilial.iAtivo
    objIntegracaoCliente.sFilialNome = objFilial.sNome
    objIntegracaoCliente.sCgc = objFilial.sCgc
    objIntegracaoCliente.sRG = objFilial.sRG
    objIntegracaoCliente.sInscricaoEstadual = objFilial.sInscricaoEstadual
    objIntegracaoCliente.sInscricaoMunicipal = objFilial.sInscricaoMunicipal
    objIntegracaoCliente.sEndereco = objEndereco.sEndereco
    objIntegracaoCliente.sBairro = objEndereco.sBairro
    objIntegracaoCliente.sCidade = objEndereco.sCidade
    objIntegracaoCliente.sSiglaEstado = objEndereco.sSiglaEstado
    objIntegracaoCliente.iCodigoPais = objEndereco.iCodigoPais
    objIntegracaoCliente.sCEP = objEndereco.sCEP
    objIntegracaoCliente.sTelefone1 = objEndereco.sTelefone1
    objIntegracaoCliente.sTelefone2 = objEndereco.sTelefone2
    objIntegracaoCliente.sEmail = objEndereco.sEmail
    objIntegracaoCliente.sFax = objEndereco.sFax
    objIntegracaoCliente.sContato = objEndereco.sContato
    objIntegracaoCliente.sEnderecoEnt = objEnderecoEnt.sEndereco
    objIntegracaoCliente.sBairroEnt = objEnderecoEnt.sBairro
    objIntegracaoCliente.sCidadeEnt = objEnderecoEnt.sCidade
    objIntegracaoCliente.sSiglaEstadoEnt = objEnderecoEnt.sSiglaEstado
    objIntegracaoCliente.iCodigoPaisEnt = objEnderecoEnt.iCodigoPais
    objIntegracaoCliente.sCEPEnt = objEnderecoEnt.sCEP
    objIntegracaoCliente.sTelefone1Ent = objEnderecoEnt.sTelefone1
    objIntegracaoCliente.sTelefone2Ent = objEnderecoEnt.sTelefone2
    objIntegracaoCliente.sEmailEnt = objEnderecoEnt.sEmail
    objIntegracaoCliente.sFaxEnt = objEnderecoEnt.sFax
    objIntegracaoCliente.sContatoEnt = objEnderecoEnt.sContato
    objIntegracaoCliente.sEnderecoCobr = objEnderecoCobr.sEndereco
    objIntegracaoCliente.sBairroCobr = objEnderecoCobr.sBairro
    objIntegracaoCliente.sCidadeCobr = objEnderecoCobr.sCidade
    objIntegracaoCliente.sSiglaEstadoCobr = objEnderecoCobr.sSiglaEstado
    objIntegracaoCliente.iCodigoPaisCobr = objEnderecoCobr.iCodigoPais
    objIntegracaoCliente.sCEPCobr = objEnderecoCobr.sCEP
    objIntegracaoCliente.sTelefone1Cobr = objEnderecoCobr.sTelefone1
    objIntegracaoCliente.sTelefone2Cobr = objEnderecoCobr.sTelefone2
    objIntegracaoCliente.sEmailCobr = objEnderecoCobr.sEmail
    objIntegracaoCliente.sFaxCobr = objEnderecoCobr.sFax
    objIntegracaoCliente.sContatoCobr = objEnderecoCobr.sContato
    objIntegracaoCliente.iComErro = 0
    objIntegracaoCliente.dtDataAtualizacao = DATA_NULA
    objIntegracaoCliente.iVendedor = objFilial.iVendedor
    objIntegracaoCliente.sObservacaoFilial = objFilial.sObservacao
    objIntegracaoCliente.lCodExterno = objFilial.lCodExterno
    
    Exporta_Dados_Cliente_Preenche_Obj = SUCESSO
     
    Exit Function
    
Erro_Exporta_Dados_Cliente_Preenche_Obj:

    Exporta_Dados_Cliente_Preenche_Obj = gErr
     
    Select Case gErr
    
        Case 189826 To 189830

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189831)
     
    End Select
    
    Exit Function

End Function

Private Function Exporta_Dados_SldProd_Preenche_Obj(ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dSaldo As Double, ByVal objIntegracaoSldProd As ClassIntegracaoSldProd) As Long
        
Dim lErro As Long
Dim bAlmValido As Boolean
Dim sProdutoTela As String

On Error GoTo Erro_Exporta_Dados_SldProd_Preenche_Obj

    bAlmValido = True
    
    lErro = Mascara_RetornaProdutoTela(sProduto, sProdutoTela)
    If lErro <> SUCESSO Then gError 190120
        
    objIntegracaoSldProd.sCodProduto = sProdutoTela
    objIntegracaoSldProd.dSaldoDisp = dSaldo
    objIntegracaoSldProd.iAlmoxarifado = iAlmoxarifado

    Exporta_Dados_SldProd_Preenche_Obj = SUCESSO
     
    Exit Function
    
Erro_Exporta_Dados_SldProd_Preenche_Obj:

    Exporta_Dados_SldProd_Preenche_Obj = gErr
     
    Select Case gErr
    
        Case 190120

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190143)
     
    End Select
    
    Exit Function

End Function

Private Function Log_Adiciona(ByVal sCritica As String, ByVal objIntegracaoLog As ClassIntegracaoLog, Optional ByVal bAviso As Boolean = False) As Long

Dim objIntegracaoLogNovo As New ClassIntegracaoLog
Dim iSeq As Integer

    iSeq = 0
    For Each objIntegracaoLogNovo In gcolLog
        If objIntegracaoLogNovo.iTipoArq = objIntegracaoLog.iTipoArq And _
            objIntegracaoLogNovo.lNumIntDoc = objIntegracaoLog.lNumIntDoc And _
            objIntegracaoLogNovo.sChave = objIntegracaoLog.sChave Then
            If iSeq < objIntegracaoLogNovo.iSeq Then
                iSeq = objIntegracaoLogNovo.iSeq
            End If
        End If
    Next
    
    objIntegracaoLogNovo.iEtapa = objIntegracaoLog.iEtapa
    objIntegracaoLogNovo.iSeq = iSeq + 1
    objIntegracaoLogNovo.iTipoArq = objIntegracaoLog.iTipoArq
    objIntegracaoLogNovo.lNumIntDoc = objIntegracaoLog.lNumIntDoc
    objIntegracaoLogNovo.sCritica = sCritica
    objIntegracaoLogNovo.sChave = objIntegracaoLog.sChave
    
    If bAviso Then
        objIntegracaoLogNovo.iAviso = MARCADO
    End If
    
    gcolLog.Add objIntegracaoLogNovo

End Function

Private Function Log_Insere(ByVal objLogBase As ClassIntegracaoLog) As Long
'Grava o Log

Dim lErro As Long
Dim iIndice As Integer
Dim alComandoExt(1 To 4) As Long
Dim lTransacaoExt As Long
Dim objIntegracaoLog As ClassIntegracaoLog
Dim iComErro As Integer
Dim lNumIntDocAnt As Long

On Error GoTo Erro_Log_Insere

    'Abre o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao2)
        If alComandoExt(iIndice) = 0 Then gError 189767
    Next
    
    'Abertura de transação
    lTransacaoExt = Transacao_AbrirExt(glConexao2)
    If lTransacaoExt = 0 Then gError 189768
    
    'Exclui o último log para o registro
    If gcolLog.Count > 0 Then
        For Each objIntegracaoLog In gcolLog
            If lNumIntDocAnt <> objIntegracaoLog.lNumIntDoc Then
                
                lNumIntDocAnt = objIntegracaoLog.lNumIntDoc
            
                lErro = Log_Exclui(alComandoExt(1), alComandoExt(2), objIntegracaoLog)
                If lErro <> SUCESSO Then gError 189769
            End If
        Next
    Else
        lErro = Log_Exclui(alComandoExt(1), alComandoExt(2), objLogBase)
        If lErro <> SUCESSO Then gError 189769
    End If
    
    iComErro = DESMARCADO

    'Para cada mensagem
    For Each objIntegracaoLog In gcolLog
    
        If objIntegracaoLog.iAviso = DESMARCADO Then
            iComErro = MARCADO
        End If
        
        lErro = Log_Insere_Aux(alComandoExt(3), objIntegracaoLog)
        If lErro <> AD_SQL_SUCESSO Then gError 189770
        
    Next
        
    If iComErro = DESMARCADO Then
    
        Set objIntegracaoLog = objLogBase
        
        Select Case objIntegracaoLog.iEtapa
        
            Case INT_ETAPA_EXP_GERARQ
                objIntegracaoLog.sCritica = "Arquivo gerado com Sucesso."
        
            Case INT_ETAPA_EXP_TABINT
                objIntegracaoLog.sCritica = "Registro exportado com Sucesso."
        
            Case INT_ETAPA_IMP_ATUALI
                objIntegracaoLog.sCritica = "Registro atualizado com Sucesso."
        
            Case INT_ETAPA_IMP_TABINT
                objIntegracaoLog.sCritica = "Arquivo importado com Sucesso."
        
        End Select
        
        objIntegracaoLog.iSeq = gcolLog.Count + 1
        objIntegracaoLog.iAviso = MARCADO
    
        lErro = Log_Insere_Aux(alComandoExt(4), objIntegracaoLog)
        If lErro <> AD_SQL_SUCESSO Then gError 189771
    
    End If
    
    Set gcolLog = New Collection
    
    'Confirma a transação
    lErro = Transacao_CommitExt(lTransacaoExt)
    If lErro <> AD_SQL_SUCESSO Then gError 189772
    
    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    Log_Insere = SUCESSO

    Exit Function

Erro_Log_Insere:

    Log_Insere = gErr

    Select Case gErr
 
        Case 189767
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 189768
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 189769, 189770, 189771
        
        Case 189772
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189773)

    End Select
    
    Call Transacao_RollbackExt(lTransacaoExt)
    
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next

    Exit Function

End Function

Private Function Log_Insere_Aux(ByVal lComando As Long, ByVal objIntegracaoLog As ClassIntegracaoLog) As Long
'Grava o Log

Dim lErro As Long

On Error GoTo Erro_Log_Insere_Aux
        
    TelaAcompanhaBatch.Log = objIntegracaoLog.sCritica & vbNewLine

    lErro = Comando_Executar(lComando, "INSERT INTO IntegracaoLog( NumIntDoc, TipoArq, " & _
            "Seq, Critica, Aviso, Data, Hora,Etapa, Chave) VALUES (?,?,?,?,?,?,?,?,?)", _
            objIntegracaoLog.lNumIntDoc, objIntegracaoLog.iTipoArq, _
            objIntegracaoLog.iSeq, objIntegracaoLog.sCritica, objIntegracaoLog.iAviso, Date, CDbl(Time), objIntegracaoLog.iEtapa, objIntegracaoLog.sChave)
    If lErro <> AD_SQL_SUCESSO Then gError 189766
    
    Log_Insere_Aux = SUCESSO

    Exit Function

Erro_Log_Insere_Aux:

    Log_Insere_Aux = gErr

    Select Case gErr

        Case 189766
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INTEGRACAOLOG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189767)

    End Select

    Exit Function

End Function

Private Function Log_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objIntegracaoLog As ClassIntegracaoLog) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_Log_Exclui
    
    'Le a tabelaLogAtualizacao
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM IntegracaoLog WHERE NumIntDoc= ? AND TipoArq= ? AND Chave = ? ", _
                0, lNumIntDoc, objIntegracaoLog.lNumIntDoc, objIntegracaoLog.iTipoArq, objIntegracaoLog.sChave)
    If lErro <> AD_SQL_SUCESSO Then gError 189760

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189761

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 189762
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM IntegracaoLog", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 189763
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189764
        
    Loop

    Log_Exclui = SUCESSO

    Exit Function

Erro_Log_Exclui:

    Log_Exclui = gErr

    Select Case gErr
        
        Case 189760, 189761, 189764
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INTEGRACAOLOG", gErr)

        Case 189762
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_INTEGRACAOLOG", gErr)

        Case 189763
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_INTEGRACAOLOG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189765)

    End Select
    
    Exit Function

End Function

Private Function Abre_Conexoes() As Long

Dim lErro As Long
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String
Dim iLenParamOut As Integer
Dim objEmpresa As New ClassDicEmpresa

On Error GoTo Erro_Abre_Conexoes

    objEmpresa.lCodigo = glEmpresa
    
    lErro = Empresa_Le(objEmpresa)
    If lErro <> SUCESSO Then gError 189756

    sDSN = objEmpresa.sStringConexao
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    glConexao1 = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If glConexao1 = 0 Then gError 189757
    
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    glConexao2 = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If glConexao2 = 0 Then gError 189758
    
    Abre_Conexoes = SUCESSO
    
    Exit Function

Erro_Abre_Conexoes:

    Abre_Conexoes = gErr

    Select Case gErr
    
        Case 189756 To 189758

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189759)

    End Select

    Exit Function

End Function

Private Function Fecha_Conexoes() As Long

On Error GoTo Erro_Fecha_Conexoes

    Call Conexao_FecharExt(glConexao1)
    Call Conexao_FecharExt(glConexao2)
    
    Set gcolLog = Nothing
    
    Fecha_Conexoes = SUCESSO
    
    Exit Function

Erro_Fecha_Conexoes:

    Fecha_Conexoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189755)

    End Select

    Exit Function

End Function

Private Function Rotina_Importar_AtualizaTelaBatch()
'Atualiza tela de acompanhamento do Batch

Dim lErro As Long
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Rotina_Importar_AtualizaTelaBatch

    'MsgBox ("4.3.1")
    
    'Atualiza tela de acompanhamento do Batch
    DoEvents

    'MsgBox ("4.3.2")
    
    If TelaAcompanhaBatch.dValorTotalImp > 0 Then
        TelaAcompanhaBatch.dValorAtualImp = TelaAcompanhaBatch.dValorAtualImp + 1
        TelaAcompanhaBatch.TotArq.Caption = CStr(TelaAcompanhaBatch.dValorAtualImp)
        TelaAcompanhaBatch.ProgressBar1.Value = CInt((TelaAcompanhaBatch.dValorAtualImp / TelaAcompanhaBatch.dValorTotalImp) * 100)
    Else
        TelaAcompanhaBatch.ProgressBar1.Value = 100
    End If

    'MsgBox ("4.3.3")
    
    If TelaAcompanhaBatch.iCancelaBatch = CANCELA_BATCH Then

        vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_IMPORTAR_EXPORTAR")

        If vbMsgBox = vbYes Then gError 189776

        TelaAcompanhaBatch.iCancelaBatch = 0

    End If
    
    Rotina_Importar_AtualizaTelaBatch = SUCESSO
    
    Exit Function

Erro_Rotina_Importar_AtualizaTelaBatch:

    Rotina_Importar_AtualizaTelaBatch = gErr

    Select Case gErr

        Case 189776

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189777)

    End Select
       
    Exit Function

End Function

Private Function Rotina_Atualizar_AtualizaTelaBatch()
'Atualiza tela de acompanhamento do Batch

Dim lErro As Long
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Rotina_Atualizar_AtualizaTelaBatch

    'Atualiza tela de acompanhamento do Batch
    DoEvents

    If TelaAcompanhaBatch.dValorTotalAtu > 0 Then
        TelaAcompanhaBatch.dValorAtualAtu = TelaAcompanhaBatch.dValorAtualAtu + 1
        TelaAcompanhaBatch.TotReg.Caption = CStr(TelaAcompanhaBatch.dValorAtualAtu)
        If TelaAcompanhaBatch.dValorAtualAtu <= TelaAcompanhaBatch.dValorTotalAtu Then
            TelaAcompanhaBatch.ProgressBar2.Value = CInt((TelaAcompanhaBatch.dValorAtualAtu / TelaAcompanhaBatch.dValorTotalAtu) * 100)
        Else
            TelaAcompanhaBatch.ProgressBar2.Value = 100
        End If
    Else
        TelaAcompanhaBatch.ProgressBar2.Value = 100
    End If

    If TelaAcompanhaBatch.iCancelaBatch = CANCELA_BATCH Then

        vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_IMPORTAR_EXPORTAR")

        If vbMsgBox = vbYes Then gError 189774

        TelaAcompanhaBatch.iCancelaBatch = 0

    End If
    
    Rotina_Atualizar_AtualizaTelaBatch = SUCESSO
    
    Exit Function

Erro_Rotina_Atualizar_AtualizaTelaBatch:

    Rotina_Atualizar_AtualizaTelaBatch = gErr

    Select Case gErr

        Case 189774

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189775)

    End Select
       
    Exit Function

End Function

Private Function Exporta_Dados_Cliente_Le(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objArqExportacaoAux As ClassArqExportacaoAux, vlCliente As Variant, viFilialCliente As Variant, vlTotalRegistros As Variant, vlCodExterno As Variant) As Long
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sSQLTotalizador As String
Dim sWhere As String
Dim objFilialCliente As ClassFilialCliente
Dim viMarcado As Variant
Dim vdtDataDe As Variant
Dim vdtDataAte As Variant

On Error GoTo Erro_Exporta_Dados_Cliente_Le
    
    vlCliente = CLng(0)
    viFilialCliente = CInt(0)
    viMarcado = MARCADO
    vdtDataDe = objArqExportacaoAux.dtExpDataDe
    vdtDataAte = objArqExportacaoAux.dtExpDataAte
    
    sSQL = "SELECT CodCliente, CodFilial, CodExterno FROM FiliaisClientes "
    sSQLTotalizador = "SELECT Count(*) FROM FiliaisClientes "
 
    If objArqExportacaoAux.iExportar = EXPORTAR_DADOS_TODOS_NAO_EXPORTADOS Then
        sWhere = sWhere & "WHERE Exportado <> ? "
    End If
 
    If objArqExportacaoAux.iExportar = EXPORTAR_DADOS_POR_PERIODO Then
    
        If objArqExportacaoAux.dtExpDataDe <> DATA_NULA Then
            sWhere = sWhere & "WHERE DataUltAtualizacao >= ? "
        End If
        
        If objArqExportacaoAux.dtExpDataAte <> DATA_NULA Then
            If Len(Trim(sWhere)) > 0 Then
                sWhere = sWhere & "AND DataUltAtualizacao <= ? "
            Else
                sWhere = "WHERE DataUltAtualizacao <= ? "
            End If
        End If
    End If
    
    sSQL = sSQL & sWhere & "ORDER BY DataUltAtualizacao "
    sSQLTotalizador = sSQLTotalizador & sWhere
 
    lErro = Comando_PrepararInt(lComando1, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 189831
 
    lErro = Comando_PrepararInt(lComando2, sSQLTotalizador)
    If (lErro <> AD_SQL_SUCESSO) Then gError 189832
 
    lErro = Comando_BindVarInt(lComando1, vlCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 189833
 
    lErro = Comando_BindVarInt(lComando2, vlTotalRegistros)
    If (lErro <> AD_SQL_SUCESSO) Then gError 189834
 
    lErro = Comando_BindVarInt(lComando1, viFilialCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 189835
    
    lErro = Comando_BindVarInt(lComando1, vlCodExterno)
    If (lErro <> AD_SQL_SUCESSO) Then gError 189835
    
    If objArqExportacaoAux.iExportar = EXPORTAR_DADOS_TODOS_NAO_EXPORTADOS Then
        lErro = Comando_BindVarInt(lComando1, viMarcado)
        If (lErro <> AD_SQL_SUCESSO) Then gError 189836
    
        lErro = Comando_BindVarInt(lComando2, viMarcado)
        If (lErro <> AD_SQL_SUCESSO) Then gError 189837
    End If
    
    If objArqExportacaoAux.iExportar = EXPORTAR_DADOS_POR_PERIODO Then
    
        If objArqExportacaoAux.dtExpDataDe <> DATA_NULA Then
            lErro = Comando_BindVarInt(lComando1, vdtDataDe)
            If (lErro <> AD_SQL_SUCESSO) Then gError 189838
        
            lErro = Comando_BindVarInt(lComando2, vdtDataDe)
            If (lErro <> AD_SQL_SUCESSO) Then gError 189839
        End If
        
        If objArqExportacaoAux.dtExpDataAte <> DATA_NULA Then
            lErro = Comando_BindVarInt(lComando1, vdtDataAte)
            If (lErro <> AD_SQL_SUCESSO) Then gError 189840

            lErro = Comando_BindVarInt(lComando2, vdtDataAte)
            If (lErro <> AD_SQL_SUCESSO) Then gError 189841
        End If
        
    End If
    
    lErro = Comando_ExecutarInt(lComando1)
    If (lErro <> AD_SQL_SUCESSO) Then gError 189842
 
    lErro = Comando_ExecutarInt(lComando2)
    If (lErro <> AD_SQL_SUCESSO) Then gError 189843
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189844
 
    Exporta_Dados_Cliente_Le = SUCESSO
 
    Exit Function
 
Erro_Exporta_Dados_Cliente_Le:
 
    Exporta_Dados_Cliente_Le = gErr
 
    Select Case gErr
 
        Case 189831 To 189844
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189845)
 
    End Select
 
    Exit Function
 
End Function

Private Function Exporta_Dados_SldProd_Le(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objArqExportacaoAux As ClassArqExportacaoAux, vsProduto As Variant, viAlmoxarifado As Variant, vdSaldo As Variant, vlTotalRegistros As Long) As Long
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sSQLTotalizador As String
Dim sWhere As String
Dim viMarcado As Variant
Dim vdtDataDe As Variant
Dim vdtDataAte As Variant

On Error GoTo Erro_Exporta_Dados_SldProd_Le
    
    vsProduto = String(STRING_PRODUTO, 0)
    viAlmoxarifado = CInt(0)
    vdSaldo = CDbl(0)
    viMarcado = MARCADO
    vdtDataDe = objArqExportacaoAux.dtExpDataDe
    vdtDataAte = objArqExportacaoAux.dtExpDataAte
    
    sSQL = "SELECT Produto, Almoxarifado, QuantDispNossa FROM EstoqueProduto "
    sSQLTotalizador = "SELECT Count(*) FROM EstoqueProduto "
 
    If objArqExportacaoAux.iExportar = EXPORTAR_DADOS_TODOS_NAO_EXPORTADOS Then
        sWhere = sWhere & "WHERE Exportado <> ? "
    End If
 
    If objArqExportacaoAux.iExportar = EXPORTAR_DADOS_POR_PERIODO Then
    
        If objArqExportacaoAux.dtExpDataDe <> DATA_NULA Then
            sWhere = sWhere & "WHERE DataUltAtualizacao >= ? "
        End If
        
        If objArqExportacaoAux.dtExpDataAte <> DATA_NULA Then
            If Len(Trim(sWhere)) > 0 Then
                sWhere = sWhere & "AND DataUltAtualizacao <= ? "
            Else
                sWhere = "WHERE DataUltAtualizacao <= ? "
            End If
        End If
    End If
    
    sSQL = sSQL & sWhere & "ORDER BY DataUltAtualizacao "
    sSQLTotalizador = sSQLTotalizador & sWhere
 
    lErro = Comando_PrepararInt(lComando1, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190102
 
    lErro = Comando_PrepararInt(lComando2, sSQLTotalizador)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190103
 
    lErro = Comando_BindVarInt(lComando1, vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190104
 
    lErro = Comando_BindVarInt(lComando2, vlTotalRegistros)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190105
 
    lErro = Comando_BindVarInt(lComando1, viAlmoxarifado)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190106
    
    lErro = Comando_BindVarInt(lComando1, vdSaldo)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190107
    
    If objArqExportacaoAux.iExportar = EXPORTAR_DADOS_TODOS_NAO_EXPORTADOS Then
        lErro = Comando_BindVarInt(lComando1, viMarcado)
        If (lErro <> AD_SQL_SUCESSO) Then gError 190108
    
        lErro = Comando_BindVarInt(lComando2, viMarcado)
        If (lErro <> AD_SQL_SUCESSO) Then gError 190109
    End If
    
    If objArqExportacaoAux.iExportar = EXPORTAR_DADOS_POR_PERIODO Then
    
        If objArqExportacaoAux.dtExpDataDe <> DATA_NULA Then
            lErro = Comando_BindVarInt(lComando1, vdtDataDe)
            If (lErro <> AD_SQL_SUCESSO) Then gError 190110
        
            lErro = Comando_BindVarInt(lComando2, vdtDataDe)
            If (lErro <> AD_SQL_SUCESSO) Then gError 190111
        End If
        
        If objArqExportacaoAux.dtExpDataAte <> DATA_NULA Then
            lErro = Comando_BindVarInt(lComando1, vdtDataAte)
            If (lErro <> AD_SQL_SUCESSO) Then gError 190112

            lErro = Comando_BindVarInt(lComando2, vdtDataAte)
            If (lErro <> AD_SQL_SUCESSO) Then gError 190113
        End If
        
    End If
    
    lErro = Comando_ExecutarInt(lComando1)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190114
 
    lErro = Comando_ExecutarInt(lComando2)
    If (lErro <> AD_SQL_SUCESSO) Then gError 190115
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190116
 
    Exporta_Dados_SldProd_Le = SUCESSO
 
    Exit Function
 
Erro_Exporta_Dados_SldProd_Le:
 
    Exporta_Dados_SldProd_Le = gErr
 
    Select Case gErr
 
        Case 190102 To 190116
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190117)
 
    End Select
 
    Exit Function
 
End Function

Private Function Exportacao_Inicia(ByVal sNomeArq As String, ByVal lTotalRegistros As Long)

Dim lErro As Long

On Error GoTo Erro_Exportacao_Inicia

    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Início da exportação de " & sNomeArq & vbNewLine

    TelaAcompanhaBatch.NomeArqImp = sNomeArq
    
    TelaAcompanhaBatch.dValorAtualImp = 0
    TelaAcompanhaBatch.dValorTotalImp = lTotalRegistros
    
    DoEvents
    
    If lTotalRegistros = 0 Then
        
        'Atualiza a barra de progressão
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 189846
    
    End If
    
    Exportacao_Inicia = SUCESSO
    
    Exit Function

Erro_Exportacao_Inicia:

    Exportacao_Inicia = gErr

    Select Case gErr
    
        Case 189846

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189847)

    End Select
       
    Exit Function

End Function

Private Function GeraArq_Inicia(ByVal lTotalArquivos As Long)

Dim lErro As Long

On Error GoTo Erro_GeraArq_Inicia

    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Início da geração dos arquivos" & vbNewLine

    TelaAcompanhaBatch.dValorAtualAtu = 0
    TelaAcompanhaBatch.dValorTotalAtu = lTotalArquivos
    
    If lTotalArquivos = 0 Then
        
        'Atualiza a barra de progressão
        lErro = Rotina_Atualizar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 189850
    
    End If
    
    GeraArq_Inicia = SUCESSO
    
    Exit Function

Erro_GeraArq_Inicia:

    GeraArq_Inicia = gErr

    Select Case gErr
    
        Case 189850

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189851)

    End Select
       
    Exit Function

End Function

Private Function ImportaArq_Inicia(ByVal lTotalArquivos As Long) As Long

Dim lErro As Long

On Error GoTo Erro_ImportaArq_Inicia

    'MsgBox ("4.3")
    
    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Início da Importação dos arquivos" & vbNewLine
    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Transferindo arquivos ... " & vbNewLine

    TelaAcompanhaBatch.dValorAtualImp = 0
    TelaAcompanhaBatch.dValorTotalImp = lTotalArquivos
    
    'MsgBox ("4.4")
    
    If lTotalArquivos = 0 Then
        
        'Atualiza a barra de progressão
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 189910
    
    End If
    
    'MsgBox ("4.5")
    
    ImportaArq_Inicia = SUCESSO
    
    Exit Function

Erro_ImportaArq_Inicia:

    ImportaArq_Inicia = gErr

    Select Case gErr
    
        Case 189910

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189911)

    End Select
       
    Exit Function

End Function

Private Function Exportacao_Obtem_NomeArq(ByVal objTipoArq As ClassTipoArqIntegracao, sNomeArq As String)
    sNomeArq = objTipoArq.sSiglaArq & Format(Now, "YYYYMMDDHHMMSS") & ".txt"
End Function

Public Function Importa_Dados_Cliente_Preenche_Cliente(ByVal objIntegracaoCliente As ClassIntegracaoCliente, ByVal objCliente As ClassCliente, ByVal colEndereco As Collection)

Dim lErro As Long
Dim objEndereco As ClassEndereco
Dim lCodigo As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sUF As String
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Importa_Dados_Cliente_Preenche_Cliente

    objFilialEmpresa.iCodFilial = giFilialEmpresa 'FILIAL_MATRIZ 'iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
    
    sUF = objFilialEmpresa.objEndereco.sSiglaEstado

    'Se é um cliente novo
    If objIntegracaoCliente.lCodCliente = 0 Then
    
        'Gera código automático do próximo cliente
        lErro = CF("Cliente_Automatico", lCodigo)
        If lErro <> SUCESSO Then gError 189934
        
        objCliente.lCodigo = lCodigo
        objIntegracaoCliente.iCodFilial = FILIAL_MATRIZ
        objIntegracaoCliente.sFilialNome = "Matriz"
    
    End If

    objCliente.dLimiteCredito = objIntegracaoCliente.dLimiteCredito
    objCliente.iAtivo = objIntegracaoCliente.iAtivo
    
    objCliente.dtDataUltVisita = DATA_NULA
    
    If objIntegracaoCliente.iCondicaoPagto <> 0 Then
        objCliente.iCondicaoPagto = objIntegracaoCliente.iCondicaoPagto
    End If
    
'    If objIntegracaoCliente.iTipo <> 0 Then
'        objCliente.iTipo = objIntegracaoCliente.iTipo
'    End If
    
    objFilialCliente.sCgc = objIntegracaoCliente.sCgc
    
    lErro = CF("Integracao_Cliente_Obtem_Tipo_Customizado", objCliente, objFilialCliente, colEndereco.Item(1), objIntegracaoCliente.iTipo, objIntegracaoCliente.sTipoCliente)
    If lErro <> SUCESSO Then gError 192485
    
    lErro = CF("Integracao_Cliente_Obtem_PadraoCobr_Cust", objCliente, objFilialCliente, colEndereco.Item(1))
    If lErro <> SUCESSO Then gError 192485
        
    objCliente.sNomeReduzido = objIntegracaoCliente.sNomeReduzido
    objCliente.sRazaoSocial = objIntegracaoCliente.sRazaoSocial
    objCliente.sObservacao = objIntegracaoCliente.sObservacao
    objCliente.iTipoFrete = TIPO_CIF
    objCliente.iTabelaPreco = objIntegracaoCliente.iTabelaPreco
    
    'Se a filial é a matriz ou é uma nova
    If objIntegracaoCliente.iCodFilial = FILIAL_MATRIZ Or objIntegracaoCliente.iCodFilial = 0 Then
    
        objCliente.sCgc = objIntegracaoCliente.sCgc
        objCliente.sRG = objIntegracaoCliente.sRG
        objCliente.sInscricaoEstadual = objIntegracaoCliente.sInscricaoEstadual
        objCliente.sInscricaoMunicipal = objIntegracaoCliente.sInscricaoMunicipal
        
        If objIntegracaoCliente.iVendedor <> 0 Then
            objCliente.iVendedor = objIntegracaoCliente.iVendedor
        End If
    
        Set objEndereco = colEndereco.Item(1)
    
        objEndereco.iCodigoPais = objIntegracaoCliente.iCodigoPais
        If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = PAIS_BRASIL
        objEndereco.sBairro = objIntegracaoCliente.sBairro
        objEndereco.sCEP = objIntegracaoCliente.sCEP
        objEndereco.sCidade = objIntegracaoCliente.sCidade
        objEndereco.sContato = objIntegracaoCliente.sContato
        objEndereco.sEmail = objIntegracaoCliente.sEmail
        objEndereco.sEndereco = objIntegracaoCliente.sEndereco
        objEndereco.sFax = objIntegracaoCliente.sFax
        objEndereco.sSiglaEstado = objIntegracaoCliente.sSiglaEstado
        If Len(Trim(objEndereco.sSiglaEstado)) = 0 Then objEndereco.sSiglaEstado = sUF
        sUF = objEndereco.sSiglaEstado
        objEndereco.sTelefone1 = objIntegracaoCliente.sTelefone1
        objEndereco.sTelefone2 = objIntegracaoCliente.sTelefone2
        
        objEndereco.iFaxDDD = objIntegracaoCliente.iFaxDDD
        objEndereco.iTelDDD1 = objIntegracaoCliente.iTelDDD1
        objEndereco.iTelDDD2 = objIntegracaoCliente.iTelDDD2
        objEndereco.lNumero = objIntegracaoCliente.lNumero
        objEndereco.sComplemento = objIntegracaoCliente.sComplemento
        objEndereco.sEmail2 = objIntegracaoCliente.sEmail2
        objEndereco.sFaxNumero = objIntegracaoCliente.sFaxNumero
        objEndereco.sTelNumero1 = objIntegracaoCliente.sTelNumero1
        objEndereco.sTelNumero2 = objIntegracaoCliente.sTelNumero2
        objEndereco.sLogradouro = objIntegracaoCliente.sLogradouro
        objEndereco.sTipoLogradouro = objIntegracaoCliente.sTipoLogradouro
        objEndereco.sReferencia = objIntegracaoCliente.sReferencia
    
        Set objEndereco = colEndereco.Item(2)
    
        objEndereco.iCodigoPais = objIntegracaoCliente.iCodigoPaisEnt
        If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = PAIS_BRASIL
        objEndereco.sBairro = objIntegracaoCliente.sBairroEnt
        objEndereco.sCEP = objIntegracaoCliente.sCEPEnt
        objEndereco.sCidade = objIntegracaoCliente.sCidadeEnt
        objEndereco.sContato = objIntegracaoCliente.sContatoEnt
        objEndereco.sEmail = objIntegracaoCliente.sEmailEnt
        objEndereco.sEndereco = objIntegracaoCliente.sEnderecoEnt
        objEndereco.sFax = objIntegracaoCliente.sFaxEnt
        objEndereco.sSiglaEstado = objIntegracaoCliente.sSiglaEstadoEnt
        If Len(Trim(objEndereco.sSiglaEstado)) = 0 Then objEndereco.sSiglaEstado = sUF
        objEndereco.sTelefone1 = objIntegracaoCliente.sTelefone1Ent
        objEndereco.sTelefone2 = objIntegracaoCliente.sTelefone2Ent
            
        objEndereco.iFaxDDD = objIntegracaoCliente.iFaxDDDEnt
        objEndereco.iTelDDD1 = objIntegracaoCliente.iTelDDD1Ent
        objEndereco.iTelDDD2 = objIntegracaoCliente.iTelDDD2Ent
        objEndereco.lNumero = objIntegracaoCliente.lNumeroEnt
        objEndereco.sComplemento = objIntegracaoCliente.sComplementoEnt
        objEndereco.sEmail2 = objIntegracaoCliente.sEmail2Ent
        objEndereco.sFaxNumero = objIntegracaoCliente.sFaxNumeroEnt
        objEndereco.sTelNumero1 = objIntegracaoCliente.sTelNumero1Ent
        objEndereco.sTelNumero2 = objIntegracaoCliente.sTelNumero2Ent
        objEndereco.sLogradouro = objIntegracaoCliente.sLogradouroEnt
        objEndereco.sTipoLogradouro = objIntegracaoCliente.sTipoLogradouroEnt
        objEndereco.sReferencia = objIntegracaoCliente.sReferenciaEnt
    
        Set objEndereco = colEndereco.Item(3)
    
        objEndereco.iCodigoPais = objIntegracaoCliente.iCodigoPaisCobr
        If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = PAIS_BRASIL
        objEndereco.sBairro = objIntegracaoCliente.sBairroCobr
        objEndereco.sCEP = objIntegracaoCliente.sCEPCobr
        objEndereco.sCidade = objIntegracaoCliente.sCidadeCobr
        objEndereco.sContato = objIntegracaoCliente.sContatoCobr
        objEndereco.sEmail = objIntegracaoCliente.sEmailCobr
        objEndereco.sEndereco = objIntegracaoCliente.sEnderecoCobr
        objEndereco.sFax = objIntegracaoCliente.sFaxCobr
        objEndereco.sSiglaEstado = objIntegracaoCliente.sSiglaEstadoCobr
        If Len(Trim(objEndereco.sSiglaEstado)) = 0 Then objEndereco.sSiglaEstado = sUF
        objEndereco.sTelefone1 = objIntegracaoCliente.sTelefone1Cobr
        objEndereco.sTelefone2 = objIntegracaoCliente.sTelefone2Cobr
    
    End If
      
    Importa_Dados_Cliente_Preenche_Cliente = SUCESSO
    
    Exit Function

Erro_Importa_Dados_Cliente_Preenche_Cliente:

    Importa_Dados_Cliente_Preenche_Cliente = gErr

    Select Case gErr
    
        Case 189934, 192485

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189935)

    End Select
       
    Exit Function

End Function

Public Function Importa_Dados_Cliente_Preenche_Filial(ByVal objIntegracaoCliente As ClassIntegracaoCliente, ByVal objFilial As ClassFilialCliente, ByVal colEndereco As Collection)

Dim lErro As Long
Dim objEndereco As ClassEndereco
Dim sIE As String
Dim sUF As String
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Importa_Dados_Cliente_Preenche_Filial

    objFilialEmpresa.iCodFilial = giFilialEmpresa 'FILIAL_MATRIZ 'iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
    
    sUF = objFilialEmpresa.objEndereco.sSiglaEstado

    objFilial.sNome = objIntegracaoCliente.sFilialNome
    If objIntegracaoCliente.iCodFilial = 0 Then
        objFilial.iCodFilial = FILIAL_MATRIZ
    Else
        objFilial.iCodFilial = objIntegracaoCliente.iCodFilial
    End If
    If Len(Trim(objFilial.sNome)) = 0 Then objFilial.sNome = "Matriz"
    
    objFilial.sCgc = objIntegracaoCliente.sCgc
    objFilial.sRG = objIntegracaoCliente.sRG
    objFilial.sInscricaoEstadual = objIntegracaoCliente.sInscricaoEstadual
    Call Formata_String_Numero(objFilial.sInscricaoEstadual, sIE)
    If Len(Trim(sIE)) = 0 Then
        objFilial.iIEIsento = MARCADO
        objFilial.iIENaoContrib = MARCADO
    Else
        objFilial.iIEIsento = DESMARCADO
        objFilial.iIENaoContrib = DESMARCADO
    End If
    objFilial.sInscricaoMunicipal = objIntegracaoCliente.sInscricaoMunicipal
    objFilial.iTipoFrete = TIPO_CIF
    
    If objIntegracaoCliente.iVendedor <> 0 Then
        objFilial.iVendedor = objIntegracaoCliente.iVendedor
    End If
    objFilial.iRegimeTributario = REGIME_TRIBUTARIO_NORMAL
        
    objFilial.sObservacao = objIntegracaoCliente.sObservacaoFilial
    objFilial.iAtivo = objIntegracaoCliente.iAtivo

    Set objEndereco = colEndereco.Item(1)

    objEndereco.iCodigoPais = objIntegracaoCliente.iCodigoPais
    If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = PAIS_BRASIL
    objEndereco.sBairro = objIntegracaoCliente.sBairro
    objEndereco.sCEP = objIntegracaoCliente.sCEP
    objEndereco.sCidade = objIntegracaoCliente.sCidade
    objEndereco.sContato = objIntegracaoCliente.sContato
    objEndereco.sEmail = objIntegracaoCliente.sEmail
    objEndereco.sEndereco = objIntegracaoCliente.sEndereco
    objEndereco.sFax = objIntegracaoCliente.sFax
    objEndereco.sSiglaEstado = objIntegracaoCliente.sSiglaEstado
    If Len(Trim(objEndereco.sSiglaEstado)) = 0 Then objEndereco.sSiglaEstado = sUF
    sUF = objEndereco.sSiglaEstado
    objEndereco.sTelefone1 = objIntegracaoCliente.sTelefone1
    objEndereco.sTelefone2 = objIntegracaoCliente.sTelefone2
    
    objEndereco.iFaxDDD = objIntegracaoCliente.iFaxDDD
    objEndereco.iTelDDD1 = objIntegracaoCliente.iTelDDD1
    objEndereco.iTelDDD2 = objIntegracaoCliente.iTelDDD2
    objEndereco.lNumero = objIntegracaoCliente.lNumero
    objEndereco.sComplemento = objIntegracaoCliente.sComplemento
    objEndereco.sEmail2 = objIntegracaoCliente.sEmail2
    objEndereco.sFaxNumero = objIntegracaoCliente.sFaxNumero
    objEndereco.sTelNumero1 = objIntegracaoCliente.sTelNumero1
    objEndereco.sTelNumero2 = objIntegracaoCliente.sTelNumero2
    objEndereco.sLogradouro = objIntegracaoCliente.sLogradouro
    objEndereco.sTipoLogradouro = objIntegracaoCliente.sTipoLogradouro
    objEndereco.sReferencia = objIntegracaoCliente.sReferencia

    Set objEndereco = colEndereco.Item(2)

    objEndereco.iCodigoPais = objIntegracaoCliente.iCodigoPaisEnt
    If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = PAIS_BRASIL
    objEndereco.sBairro = objIntegracaoCliente.sBairroEnt
    objEndereco.sCEP = objIntegracaoCliente.sCEPEnt
    objEndereco.sCidade = objIntegracaoCliente.sCidadeEnt
    objEndereco.sContato = objIntegracaoCliente.sContatoEnt
    objEndereco.sEmail = objIntegracaoCliente.sEmailEnt
    objEndereco.sEndereco = objIntegracaoCliente.sEnderecoEnt
    objEndereco.sFax = objIntegracaoCliente.sFaxEnt
    objEndereco.sSiglaEstado = objIntegracaoCliente.sSiglaEstadoEnt
    If Len(Trim(objEndereco.sSiglaEstado)) = 0 Then objEndereco.sSiglaEstado = sUF
    objEndereco.sTelefone1 = objIntegracaoCliente.sTelefone1Ent
    objEndereco.sTelefone2 = objIntegracaoCliente.sTelefone2Ent
    
    objEndereco.iFaxDDD = objIntegracaoCliente.iFaxDDDEnt
    objEndereco.iTelDDD1 = objIntegracaoCliente.iTelDDD1Ent
    objEndereco.iTelDDD2 = objIntegracaoCliente.iTelDDD2Ent
    objEndereco.lNumero = objIntegracaoCliente.lNumeroEnt
    objEndereco.sComplemento = objIntegracaoCliente.sComplementoEnt
    objEndereco.sEmail2 = objIntegracaoCliente.sEmail2Ent
    objEndereco.sFaxNumero = objIntegracaoCliente.sFaxNumeroEnt
    objEndereco.sTelNumero1 = objIntegracaoCliente.sTelNumero1Ent
    objEndereco.sTelNumero2 = objIntegracaoCliente.sTelNumero2Ent
    objEndereco.sLogradouro = objIntegracaoCliente.sLogradouroEnt
    objEndereco.sReferencia = objIntegracaoCliente.sReferenciaEnt

    Set objEndereco = colEndereco.Item(3)

    objEndereco.iCodigoPais = objIntegracaoCliente.iCodigoPaisCobr
    If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = PAIS_BRASIL
    objEndereco.sBairro = objIntegracaoCliente.sBairroCobr
    objEndereco.sCEP = objIntegracaoCliente.sCEPCobr
    objEndereco.sCidade = objIntegracaoCliente.sCidadeCobr
    objEndereco.sContato = objIntegracaoCliente.sContatoCobr
    objEndereco.sEmail = objIntegracaoCliente.sEmailCobr
    objEndereco.sEndereco = objIntegracaoCliente.sEnderecoCobr
    objEndereco.sFax = objIntegracaoCliente.sFaxCobr
    objEndereco.sSiglaEstado = objIntegracaoCliente.sSiglaEstadoCobr
    If Len(Trim(objEndereco.sSiglaEstado)) = 0 Then objEndereco.sSiglaEstado = sUF
    objEndereco.sTelefone1 = objIntegracaoCliente.sTelefone1Cobr
    objEndereco.sTelefone2 = objIntegracaoCliente.sTelefone2Cobr
    
    Importa_Dados_Cliente_Preenche_Filial = SUCESSO
    
    Exit Function

Erro_Importa_Dados_Cliente_Preenche_Filial:

    Importa_Dados_Cliente_Preenche_Filial = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189936)

    End Select
       
    Exit Function

End Function

Private Function Importa_Dados_PV_Preenche_PV(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long, ByVal lComando8 As Long, ByVal lComando9 As Long, ByVal lComando10 As Long, ByVal colIntegracaoPV As Collection, ByVal objPVTela As ClassPedidoDeVenda, ByVal objLogBase As ClassIntegracaoLog, bErroNoRegistro As Boolean, bCancelamento As Boolean, ByVal objTelaPedido As Object, ByVal objTelaPedidoOCX As Object, ByVal lComando11 As Long, ByVal lComando12 As Long)

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim objItemPV2 As ClassItemPedido
Dim lCodigo As Long
Dim objIntegracaoPV As ClassIntegracaoPV
Dim lCliente As Long
Dim iFilial As Integer
Dim iCondPagto As Integer
Dim iVendedor As Integer
Dim sProduto As String
Dim iTransp As Integer
Dim bAchouCliente As Boolean
Dim bAchouCondPagto As Boolean
Dim bAchouVendedor As Boolean
Dim bAchouProduto As Boolean
Dim bAchouTransp As Boolean
Dim bAchouAlm As Boolean
Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim objMensagem As New ClassMensagem
Dim objTipoCliente As New ClassTipoCliente
Dim objTransportadora As New ClassTransportadora
Dim objProduto As ClassProduto
Dim iSeq As Integer
Dim iProdutoPreenchido As Integer
Dim objEndereco As New ClassEndereco
Dim objOV As New ClassOrcamentoVenda
Dim objPV As New ClassPedidoDeVenda
Dim iAlmoxarifado As Integer
Dim bNovoPedido As Boolean
Dim sProdutoRegistro As String
Dim objReservaItem As ClassReservaItem
Dim objComissao As ClassComissaoPedVendas
Dim objComissaoAux As Object
Dim sNaturezaOPAnt As String
Dim iIndice As Integer, iCodTabelaPrecoItem As Integer
Dim dValorPV As Double, dPrecoUnitario As Double, iCodTabelaPreco As Integer
Dim colItensRomaneioGrade As New Collection
Dim objKitVenda As New ClassKitVenda, colComissaoAux As Collection

On Error GoTo Erro_Importa_Dados_PV_Preenche_PV

    If colIntegracaoPV.Count = 0 Then Exit Function
    
    Set objIntegracaoPV = colIntegracaoPV.Item(1)
    
    'É novo
    If objPVTela.lCodigo = 0 Then
        bNovoPedido = True
    Else
        bNovoPedido = False
    End If

    'É novo
    If bNovoPedido Then
    
        If gobjFAT.iImpDadosInterno = MARCADO Then
            lCodigo = objIntegracaoPV.lCodPedidoExterno
        Else

            'Obtém o código
            lErro = CF("PedidoDeVenda_Automatico", lCodigo)
            If lErro <> SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Ocorreu um erro na obtenção do código do pedido de venda.", objLogBase)
            Else
                Call Log_Adiciona("O código do pedido no corporator será " & CStr(lCodigo) & ".", objLogBase, True)
            End If
            
        End If
        
        objPV.lCodigo = lCodigo
        objPV.iFilialEmpresa = giFilialEmpresa
        
        objPV.dtDataEntrega = DATA_NULA
        
        'Preenche o cliente
        'Se não tiver o código do corporator tenta pelo código externo e depois pelo CGC
        If objIntegracaoPV.lCodClienteCorporator = 0 Or objIntegracaoPV.iCodFilialCorporator = 0 Then

            Call Log_Adiciona("O Cliente não possui o código do corporator.", objLogBase, True)

            If objIntegracaoPV.lCodClienteExterno <> 0 Then

                lErro = Comando_Executar(lComando1, "SELECT CodCliente, CodFilial FROM FiliaisClientes WHERE CodExterno = ? ", lCliente, iFilial, objIntegracaoPV.lCodClienteExterno)
                If lErro <> AD_SQL_SUCESSO Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Ocorreu um erro na leitura da filial cliente pelo código externo.", objLogBase)
                End If

                lErro = Comando_BuscarProximo(lComando1)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Ocorreu um erro na leitura da filial cliente pelo código externo.", objLogBase)
                End If

                If lErro = AD_SQL_SUCESSO Then
                    objPV.lCliente = lCliente
                    objPV.iFilial = iFilial
                    bAchouCliente = True
                    Call Log_Adiciona("O Cliente foi localizado pelo código externo. " & "Cliente: " & CStr(lCliente) & " Filial: " & CStr(iFilial), objLogBase, True)
                Else
                    Call Log_Adiciona("O código externo do cliente não foi localizado no corporator.", objLogBase, True)
                End If
                
                'Se tem o código externo e veio da app de pesquisa de preço pega da tabela de IntegracaoCliente o código do mesmo
                If objIntegracaoPV.iCanalVenda = 999 Then
                
                    lErro = Comando_Executar(lComando12, "SELECT CodCliente, CodFilial FROM IntegracaoCliente WHERE CodExterno = ? ORDER BY NumIntDoc DESC", lCliente, iFilial, objIntegracaoPV.lCodClienteExterno)
                    If lErro <> AD_SQL_SUCESSO Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Ocorreu um erro na leitura da filial cliente pelo código externo.", objLogBase)
                    End If
    
                    lErro = Comando_BuscarProximo(lComando12)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Ocorreu um erro na leitura da filial cliente pelo código externo.", objLogBase)
                    End If
    
                    If lErro = AD_SQL_SUCESSO Then
                        objPV.lCliente = lCliente
                        objPV.iFilial = iFilial
                        bAchouCliente = True
                        Call Log_Adiciona("O Cliente foi localizado pelo código externo. " & "Cliente: " & CStr(lCliente) & " Filial: " & CStr(iFilial), objLogBase, True)
                    Else
                        Call Log_Adiciona("O código externo do cliente não foi localizado no corporator.", objLogBase, True)
                    End If
                
                End If
                

            End If

            If Len(Trim(objIntegracaoPV.sCGCCliente)) <> 0 And Not bAchouCliente Then

                lErro = Comando_Executar(lComando2, "SELECT CodCliente, CodFilial FROM FiliaisClientes WHERE CGC = ? ", lCliente, iFilial, objIntegracaoPV.sCGCCliente)
                If lErro <> AD_SQL_SUCESSO Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Ocorreu um erro na leitura da filial empresa pelo CGC.", objLogBase)
                End If
                
                lErro = Comando_BuscarProximo(lComando2)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Ocorreu um erro na leitura da filial empresa pelo CGC.", objLogBase)
                End If
                
                If lErro = AD_SQL_SUCESSO Then
                    objPV.lCliente = lCliente
                    objPV.iFilial = iFilial
                    bAchouCliente = True
                    Call Log_Adiciona("O Cliente foi localizado pelo CGC. " & "Cliente: " & CStr(lCliente) & " Filial: " & CStr(iFilial), objLogBase, True)
                Else
                    Call Log_Adiciona("O CGC do cliente não foi localizado no corporator.", objLogBase, True)
                End If

            End If

        Else
        
            objPV.lCliente = objIntegracaoPV.lCodClienteCorporator
            objPV.iFilial = objIntegracaoPV.iCodFilialCorporator

        End If
        
        objFilialCliente.lCodCliente = objPV.lCliente
        objFilialCliente.iCodFilial = objPV.iFilial
        
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then
            bErroNoRegistro = True
            Call Log_Adiciona("Não foi possível ler a filial do cliente. Cliente: " & CStr(objFilialCliente.lCodCliente) & " Filial: " & CStr(objFilialCliente.iCodFilial) & ".", objLogBase)
        End If
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Filial cliente não cadastrada. Cliente: " & CStr(objFilialCliente.lCodCliente) & " Filial: " & CStr(objFilialCliente.iCodFilial) & ".", objLogBase)
        End If
        
        objCliente.lCodigo = objPV.lCliente
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then
            bErroNoRegistro = True
            Call Log_Adiciona("Não foi possível ler o cliente. Cliente: " & CStr(objCliente.lCodigo) & ".", objLogBase)
        End If
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Cliente não cadastrado. Cliente: " & CStr(objCliente.lCodigo) & ".", objLogBase)
        End If
        
        objTipoCliente.iCodigo = objCliente.iTipo
        
        lErro = CF("TipoCliente_Le", objTipoCliente)
        If lErro <> SUCESSO And lErro <> 19062 Then
            bErroNoRegistro = True
            Call Log_Adiciona("Não foi possível ler o tipo de cliente. Tipo " & CStr(objTipoCliente.iCodigo) & ".", objLogBase)
        End If
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Tipo de cliente não cadastrado. Tipo " & CStr(objTipoCliente.iCodigo) & ".", objLogBase)
        End If
        
        objEndereco.lCodigo = objFilialCliente.lEndereco
    
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO And lErro <> 12309 Then
            bErroNoRegistro = True
            Call Log_Adiciona("Não foi possível ler o endereço. Endereço: " & CStr(objEndereco.lCodigo) & ".", objLogBase)
        End If
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Endereço não cadastrado. Endereço: " & CStr(objEndereco.lCodigo) & ".", objLogBase)
        End If
        
        objPV.dtDataEmissao = objIntegracaoPV.dtDataEmissao
        objPV.dtDataReferencia = objIntegracaoPV.dtDataEmissao
        objPV.dtDataRefFluxo = objIntegracaoPV.dtDataEmissao
        
        If objIntegracaoPV.iCodCondPagto = 0 Then
        
            Call Log_Adiciona("A condição de pagamento não possui o código do corporator.", objLogBase, True)

            If objIntegracaoPV.iCodCondPagtoExterno <> 0 Then

                lErro = Comando_Executar(lComando3, "SELECT Codigo FROM CondicoesPagto WHERE CodExterno = ? ", iCondPagto, objIntegracaoPV.iCodCondPagtoExterno)
                If lErro <> AD_SQL_SUCESSO Then gError 189944

                lErro = Comando_BuscarProximo(lComando3)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189945

                If lErro = AD_SQL_SUCESSO Then
                    objPV.iCondicaoPagto = iCondPagto
                    bAchouCondPagto = True
                    Call Log_Adiciona("A condição de pagamento foi localizada pelo código externo. " & "Condição de pagamento: " & CStr(iCondPagto), objLogBase, True)
                Else
                    Call Log_Adiciona("A condição de pagamento não foi localizada pelo código externo.", objLogBase, True)

                End If

            End If
            
            If Len(Trim(objIntegracaoPV.sNomeCondPagtoExterno)) <> 0 And Not bAchouCondPagto Then

                lErro = Comando_Executar(lComando4, "SELECT Codigo FROM CondicoesPagto WHERE DescReduzida = ? ", iCondPagto, objIntegracaoPV.sNomeCondPagtoExterno)
                If lErro <> AD_SQL_SUCESSO Then gError 189944

                lErro = Comando_BuscarProximo(lComando4)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189945

                If lErro = AD_SQL_SUCESSO Then
                    objPV.iCondicaoPagto = iCondPagto
                    bAchouCondPagto = True
                    Call Log_Adiciona("A condição de pagamento foi localizada pela descrição. " & "Condição de pagamento: " & CStr(iCondPagto), objLogBase, True)
                Else
                    Call Log_Adiciona("A condição de pagamento não foi localizada pela descrição.", objLogBase, True)

                End If

            End If
        Else
            objPV.iCondicaoPagto = objIntegracaoPV.iCodCondPagto
        End If
        
        If objIntegracaoPV.iCodTabelaPreco = 0 And objPV.iCondicaoPagto <> 0 And InStr(UCase(gsNomeEmpresa), "FEMAR") <> 0 Then
        
            lErro = Comando_Executar(lComando11, "SELECT TabelaPreco FROM CondicoesPagto WHERE Codigo = ?", iCodTabelaPreco, objPV.iCondicaoPagto)
            If lErro <> AD_SQL_SUCESSO Then gError 201176

            lErro = Comando_BuscarProximo(lComando11)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201177
            If lErro = AD_SQL_SUCESSO Then objPV.iTabelaPreco = iCodTabelaPreco
            
        Else
            objPV.iTabelaPreco = objIntegracaoPV.iCodTabelaPreco
        End If
        
        dValorPV = 0
        For Each objIntegracaoPV In colIntegracaoPV
            
            If objIntegracaoPV.iCodTabelaPrecoItem = 0 Then
                iCodTabelaPrecoItem = objPV.iTabelaPreco
            Else
                iCodTabelaPrecoItem = objIntegracaoPV.iCodTabelaPrecoItem
            End If
            
            dPrecoUnitario = objIntegracaoPV.dPrecoUnitario
            lErro = CF("PV_CalculaPrecoUnitario", giFilialEmpresa, objIntegracaoPV.sCodProduto, iCodTabelaPrecoItem, objPV.dtDataEmissao, dPrecoUnitario)
            If lErro <> SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Erro no calculo do preco unitario. Pedido " & CStr(objIntegracaoPV.lCodClienteExterno) & ", Item " & CStr(objIntegracaoPV.iItem) & ".", objLogBase)
            End If
            objIntegracaoPV.dPrecoUnitario = dPrecoUnitario
            
            dValorPV = dValorPV + objIntegracaoPV.dPrecoUnitario * objIntegracaoPV.dQuantidadePedida - objIntegracaoPV.dValorDescontoItem
        Next
        
        Set objIntegracaoPV = colIntegracaoPV.Item(1)
        
        objPV.dValorTotal = dValorPV
        
        If Len(Trim(objIntegracaoPV.sNaturezaOp)) > 0 Then
            objPV.iNaturezaOpManual = VAR_PREENCH_MANUAL
            objPV.sNaturezaOp = objIntegracaoPV.sNaturezaOp
        Else
        
            objPV.iNaturezaOpManual = VAR_PREENCH_AUTOMATICO
        
            lErro = CF("Integracao_PV_Obtem_NaturezaOP_Customizado", objPV, objFilialCliente, objEndereco)
            If lErro <> SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Não foi possível obter a natureza de operação.", objLogBase)
            End If
            
        End If
            
        objPV.iFilialEmpresaFaturamento = objIntegracaoPV.iFilialEmpresaFat
        objPV.dValorDesconto = objIntegracaoPV.dValorDescontoPedido
        objPV.iFreteRespons = objIntegracaoPV.iFrete + 1
        objPV.dValorFrete = objIntegracaoPV.dValorFretePedido
        objPV.dValorSeguro = objIntegracaoPV.dValorSeguroPedido
        objPV.dValorOutrasDespesas = objIntegracaoPV.dValorOutrasDespesasPedido
        objPV.iFilialEntrega = objIntegracaoPV.iFilialEmpresaEnt
        
        If objPV.iFilialEmpresaFaturamento = 0 Then objPV.iFilialEmpresaFaturamento = giFilialEmpresa
        If objPV.iFilialEntrega = 0 Then objPV.iFilialEntrega = objPV.iFilial  'giFilialEmpresa
        
        If objPV.dValorFrete = 0 Then
            lErro = CF("Integracao_PV_Obtem_Frete_Customizado", objPV, objTipoCliente)
            If lErro <> SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Não foi possível obter o frete.", objLogBase)
            End If
            
        End If
        
        bAchouVendedor = False
        If objIntegracaoPV.iCodVendedor = 0 Then
                
            If objIntegracaoPV.iCodVendedorExterno <> 0 Then

                lErro = Comando_Executar(lComando7, "SELECT Codigo FROM Vendedores WHERE CodExterno = ? ", iVendedor, objIntegracaoPV.iCodVendedorExterno)
                If lErro <> AD_SQL_SUCESSO Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Erro na leitura do vendedor com código externo " & CStr(objIntegracaoPV.iCodVendedorExterno) & ".", objLogBase)
                End If

                lErro = Comando_BuscarProximo(lComando7)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Erro na leitura do vendedor com código externo " & CStr(objIntegracaoPV.iCodVendedorExterno) & ".", objLogBase)
                End If

                If lErro = AD_SQL_SUCESSO Then
                    bAchouVendedor = True
                    Call Log_Adiciona("O Vendedor foi localizado pelo código externo. " & "Vendedor: " & CStr(iVendedor) & ".", objLogBase, True)
                Else
                    Call Log_Adiciona("O Vendedor não foi localizado pelo código externo.", objLogBase, True)
                End If

            End If
            
            If Len(Trim(objIntegracaoPV.sNomeVendedorExterno)) <> 0 And Not bAchouVendedor Then

                lErro = Comando_Executar(lComando8, "SELECT Codigo FROM Vendedores WHERE NomeReduzido = ? ", iVendedor, objIntegracaoPV.sNomeVendedorExterno)
                If lErro <> AD_SQL_SUCESSO Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Erro na leitura do vendedor com nome reduzido " & CStr(objIntegracaoPV.sNomeVendedorExterno) & ".", objLogBase)
                End If

                lErro = Comando_BuscarProximo(lComando8)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Erro na leitura do vendedor com nome reduzido " & CStr(objIntegracaoPV.sNomeVendedorExterno) & ".", objLogBase)
                End If

                If lErro = AD_SQL_SUCESSO Then
                    bAchouVendedor = True
                    Call Log_Adiciona("O Vendedor foi localizado pelo nome reduzido. " & "Vendedor: " & CStr(iVendedor) & ".", objLogBase, True)
                Else
                    Call Log_Adiciona("O Vendedor não foi localizado pelo nome reduzido.", objLogBase, True)
                End If

            End If
            
        Else
            iVendedor = objIntegracaoPV.iCodVendedor
        End If
        
        If objIntegracaoPV.iTrazerTranspAuto = MARCADO Then
        
            If objFilialCliente.iCodTransportadora <> 0 Then
    
                objTransportadora.iCodigo = objFilialCliente.iCodTransportadora
                
                lErro = CF("Transportadora_Le", objTransportadora)
                If lErro <> SUCESSO And lErro <> 19250 Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Erro na leitura da transportadora código " & CStr(objTransportadora.iCodigo) & ".", objLogBase)
                End If
                If lErro = SUCESSO Then
                    objPV.iCodTransportadora = objTransportadora.iCodigo
                Else
                    bErroNoRegistro = True
                    Call Log_Adiciona("Transportadora " & CStr(objTransportadora.iCodigo) & " não cadastrada.", objLogBase)
                End If
    
            ElseIf objCliente.iTipo <> 0 Then
    
                'Testa se TipoCliente tem Transportadora
                If objTipoCliente.iCodTransportadora <> 0 Then
                
                    objTransportadora.iCodigo = objTipoCliente.iCodTransportadora
                    
                    lErro = CF("Transportadora_Le", objTransportadora)
                    If lErro <> SUCESSO And lErro <> 19250 Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Erro na leitura da transportadora " & CStr(objTransportadora.iCodigo) & ".", objLogBase)
                    End If
                    If lErro = SUCESSO Then
                        objPV.iCodTransportadora = objTransportadora.iCodigo
                    Else
                        bErroNoRegistro = True
                        Call Log_Adiciona("Transportadora " & CStr(objTransportadora.iCodigo) & " não cadastrada.", objLogBase)
                    End If
                End If
    
            End If

        Else
        
            If objIntegracaoPV.iCodTransportadora = 0 Then
            
                Call Log_Adiciona("A transportadora não possui o código do corporator.", objLogBase, True)
    
                If objIntegracaoPV.iCodTransportadoraExterno <> 0 Then
    
                    lErro = Comando_Executar(lComando5, "SELECT Codigo FROM Transportadora WHERE CodExterno = ? ", iTransp, objIntegracaoPV.iCodTransportadoraExterno)
                    If lErro <> AD_SQL_SUCESSO Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Erro na leitura da transportadora com código externo " & CStr(objIntegracaoPV.iCodTransportadoraExterno) & ".", objLogBase)
                    End If
    
                    lErro = Comando_BuscarProximo(lComando5)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Erro na leitura da transportadora com código externo " & CStr(objIntegracaoPV.iCodTransportadoraExterno) & ".", objLogBase)
                    End If
    
                    If lErro = AD_SQL_SUCESSO Then
                        objPV.iCodTransportadora = iTransp
                        bAchouTransp = True
                        Call Log_Adiciona("A transportadora foi localizada pelo código externo. " & "Transportadora: " & CStr(iTransp), objLogBase, True)
                    Else
                        Call Log_Adiciona("A transportadora não foi localizada pelo código externo.", objLogBase, True)
                    End If
    
                End If
                
                If Len(Trim(objIntegracaoPV.sNomeTransportadoraExterno)) <> 0 And Not bAchouTransp Then
    
                    lErro = Comando_Executar(lComando6, "SELECT Codigo FROM Transportadora WHERE NomeReduzido = ? ", iTransp, objIntegracaoPV.sNomeTransportadoraExterno)
                    If lErro <> AD_SQL_SUCESSO Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Erro na leitura da transportadora com nome externo " & CStr(objIntegracaoPV.sNomeTransportadoraExterno) & ".", objLogBase)
                    End If
    
                    lErro = Comando_BuscarProximo(lComando6)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Erro na leitura da transportadora com nome externo " & CStr(objIntegracaoPV.sNomeTransportadoraExterno) & ".", objLogBase)
                    End If
    
                    If lErro = AD_SQL_SUCESSO Then
                        objPV.iCodTransportadora = iTransp
                        bAchouTransp = True
                        Call Log_Adiciona("A transportadora foi localizada pela descrição. " & "Transportadora: " & CStr(iTransp), objLogBase, True)
                    Else
                        Call Log_Adiciona("A transportadora não foi localizada pela descrição.", objLogBase, True)
                    End If
    
                End If
            
                '????
                If Not bAchouTransp And InStr(UCase(gsNomeEmpresa), "MULTIPLUS") <> 0 Then
                    objPV.iCodTransportadora = 1 'NOSSO CARRO
                End If
            
            Else
                objPV.iCodTransportadora = objIntegracaoPV.iCodTransportadora
            End If
        
        End If
        
        If objIntegracaoPV.iTrazerMensagemAuto = MARCADO Then

            If objFilialCliente.iCodMensagem > 0 Then
                objMensagem.iCodigo = objFilialCliente.iCodMensagem
            Else
                objMensagem.iCodigo = objCliente.iCodMensagem
            End If
            
            If objMensagem.iCodigo <> 0 Then
            
                lErro = CF("Mensagem_Le", objMensagem)
                If lErro <> SUCESSO And lErro <> 19234 Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Erro na leitura da mensagem de código " & CStr(objMensagem.iCodigo) & ".", objLogBase)
                End If

                If lErro <> SUCESSO Then
                    bErroNoRegistro = True
                    Call Log_Adiciona("Mensagem de código " & CStr(objMensagem.iCodigo) & " não cadastrada.", objLogBase)
                End If
            
                objPV.sMensagemNota = objMensagem.sDescricao
                
            End If
                    
        Else
            objPV.sMensagemNota = objIntegracaoPV.sMensagemPedido
        End If
       
        If objIntegracaoPV.iTrazerCanalVendaAuto = MARCADO Then
        
            objPV.iCanalVenda = 0
            
            lErro = CF("Integracao_PV_Obtem_CanalVenda_Customizado", objPV, objTipoCliente)
            If lErro <> AD_SQL_SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Não foi possível obter o canal de venda.", objLogBase)
            End If
            
        Else
            objPV.iCanalVenda = objIntegracaoPV.iCanalVenda
        End If
        
        objPV.sPedidoCliente = CStr(objIntegracaoPV.lCodPedidoExterno)
        
        iSeq = 0
        For Each objIntegracaoPV In colIntegracaoPV
            
            objLogBase.lNumIntDoc = objIntegracaoPV.lNumIntDoc
            
            iSeq = objIntegracaoPV.iItem
            Set objItemPV = New ClassItemPedido
            Set objProduto = New ClassProduto
            
            objItemPV.iItem = iSeq
            
            If gobjFAT.iImpDadosInterno = MARCADO Then
                sProduto = objIntegracaoPV.sCodProduto
            Else
                If Len(Trim(objIntegracaoPV.sCodProduto)) > 0 Then
                    sProduto = objIntegracaoPV.sCodProduto
                Else
                    sProdutoRegistro = objIntegracaoPV.sCodProdutoExterno
                    lErro = CF("Produto_Formata", sProdutoRegistro, sProduto, iProdutoPreenchido)
                    If lErro <> SUCESSO Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Não foi possível formatar o produto. Produto: " & sProdutoRegistro & ".", objLogBase)
                    End If
                End If
            End If
                       
            objProduto.sCodigo = sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then
                bErroNoRegistro = True
                Call Log_Adiciona("Não foi possível ler o produto. Produto: " & sProduto & ".", objLogBase)
            End If
            If lErro = 28030 Then
                bErroNoRegistro = True
                Call Log_Adiciona("Produto não cadastrado. Produto: " & sProduto & ".", objLogBase)
            End If
            
            If objIntegracaoPV.sUM = "" Then objIntegracaoPV.sUM = objProduto.sSiglaUMVenda
                       
            objItemPV.sProduto = objProduto.sCodigo
            objIntegracaoPV.sProdutoFormatado = objProduto.sCodigo
            
            If objIntegracaoPV.iTrazerDescricaoAuto = MARCADO Then
                objItemPV.sDescricao = objProduto.sDescricao
                objItemPV.sProdutoDescricao = objProduto.sDescricao
            Else
                objItemPV.sDescricao = objIntegracaoPV.sDescricaoItem
                objItemPV.sProdutoDescricao = objIntegracaoPV.sDescricaoItem
            End If
            
            If objIntegracaoPV.dQuantidadeCancelada - objIntegracaoPV.dQuantidadePedida > -QTDE_ESTOQUE_DELTA Then
                objIntegracaoPV.dQuantidadePedida = objIntegracaoPV.dQuantidadeCancelada
                objIntegracaoPV.dValorDescontoItem = 0
                objItemPV.dQuantCancelada = 0
            Else
                objItemPV.dQuantCancelada = objIntegracaoPV.dQuantidadeCancelada
            End If
            
            objItemPV.dQuantidade = objIntegracaoPV.dQuantidadePedida
            'objItemPV.dQuantCancelada = objIntegracaoPV.dQuantidadeCancelada
            objItemPV.sUnidadeMed = objIntegracaoPV.sUM
            
            '????
            If InStr(UCase(gsNomeEmpresa), "MULTIPLUS") <> 0 Then
                objPV.lVolumeQuant = objPV.lVolumeQuant + objItemPV.dQuantidade
            End If
            
            objItemPV.dPrecoUnitario = objIntegracaoPV.dPrecoUnitario
            objItemPV.dValorDesconto = objIntegracaoPV.dValorDescontoItem
            objItemPV.dtDataEntrega = objIntegracaoPV.dtDataEntrega
            'objItemPV.dPrecoTotal = objIntegracaoPV.dPrecoUnitario * (objItemPV.dQuantidade - objIntegracaoPV.dQuantidadeCancelada)
            objItemPV.dPrecoTotal = objIntegracaoPV.dPrecoUnitario * (objItemPV.dQuantidade - objItemPV.dQuantCancelada) - objItemPV.dValorDesconto

            If objItemPV.dPrecoTotal < DELTA_VALORMONETARIO Then objItemPV.dPrecoTotal = 0

            lErro = CF("Integracao_PV_Obtem_UM_Customizado", objItemPV, objProduto)
            If lErro <> SUCESSO Then
                bErroNoRegistro = True
                Call Log_Adiciona("Não foi possível obter a unidade de media.", objLogBase)
            End If
            
            Set colItensRomaneioGrade = New Collection
            
            If objProduto.iKitVendaComp = MARCADO Then
            
                Set objKitVenda = New ClassKitVenda
            
                lErro = CF("ProdutoKitVenda_Le_Romaneio", objProduto, colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 180031
                
                objKitVenda.sProduto = objProduto.sCodigo
                objKitVenda.sUM = objItemPV.sUnidadeMed
                objKitVenda.dQuantidade = objItemPV.dQuantidade
                
                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 180032
            
            End If
            
            With objItemPV
                Set objItemPV2 = objPV.IncluirItem(.sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .ColReserva, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, objProduto.sSiglaUMEstoque, .iClasseUM, .dQuantFaturada, 0, .sDescricao, .iStatus, objProduto.iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, 0, Nothing)
            End With
            
            Set objItemPV2.colItensRomaneioGrade = colItensRomaneioGrade
            
            objItemPV2.objTributacao.iOrigemMercadoria = objProduto.iOrigemMercadoria
       
        Next
        
        Set objIntegracaoPV = colIntegracaoPV.Item(1)
        
    End If
    
    'Verifica se todos os itens estão cancelados
    bCancelamento = True
    For Each objIntegracaoPV In colIntegracaoPV
        If objIntegracaoPV.dQuantidadePedida - objIntegracaoPV.dQuantidadeCancelada > QTDE_ESTOQUE_DELTA Then
            bCancelamento = False
            Exit For
        End If
    Next
    
    Set objIntegracaoPV = colIntegracaoPV.Item(1)
    
    If bCancelamento Then
        objPV.dValorDesconto = 0
    End If
    
    If bNovoPedido Then
    
        If objPVTela.colParcelas.Count > 0 Then Set objPV.colParcelas = objPVTela.colParcelas
    
        lErro = objTelaPedido.Traz_Pedido_Tela(objPV, objOV)
    Else
        lErro = objTelaPedido.Trata_Parametros(objPVTela)
    End If
    If lErro <> SUCESSO Then
        bErroNoRegistro = True
        Call Log_Adiciona("Não conseguiu trazer o pedido para tela.", objLogBase)
    End If
    
    'É novo
    If bNovoPedido Then
    
        lErro = objTelaPedido.Trata_FilialCliente
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Não conseguiu trazer o pedido para tela.", objLogBase)
        End If

        For iSeq = 1 To colIntegracaoPV.Count
            Call CF2(objTelaPedido, "PV_PrecoTotal_Calcula", iSeq)
        Next
        
        lErro = objTelaPedido.ValorTotal_Calcula()
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Não conseguiu calcular o valor total do pedido.", objLogBase)
        End If
        
        objTelaPedido.gobjTribTab.NatOpInterna_Change
        sNaturezaOPAnt = objPV.sNaturezaOp
        
        Call objTelaPedido.gobjTribTab.TributacaoRecalcular_Click
        
        objTelaPedido.gobjTribTab.sNatOpInterna = sNaturezaOPAnt
        Call objTelaPedido.gobjTribTab.NatOpInterna_Validate(bSGECancelDummy)
        
        objTelaPedido.cobrancaautomatica.Value = 1
    
        Call objTelaPedido.Cobranca_Automatica
        
        If iVendedor <> 0 Then
            objFilialCliente.iVendedor = iVendedor
        End If
        
        lErro = objTelaPedido.objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Não conseguiu calcular a comissão.", objLogBase)
        End If
        
        Call objTelaPedido.objTabComissoes.Comissoes_Calcula
        
        If objIntegracaoPV.iTrazerPesoAuto = MARCADO Then
            'Calcula o peso dos itens
            Call objTelaPedido.Pesos_Recalcula
        
        End If
        
        '????
        If InStr(UCase(gsNomeEmpresa), "MULTIPLUS") <> 0 And objPV.lVolumeQuant > 0 Then
            objTelaPedido.VolumeQuant.Text = CStr(objPV.lVolumeQuant)
        End If
        
        'Faz a reserva automática para todos os itens
        If objIntegracaoPV.iTrazerReservaAuto = MARCADO Then
            Call objTelaPedido.BotaoRefazAlocacao_Click
        End If
                
    End If
    
    Set objPVTela.colParcelas = New ColParcelaPedidoVenda
        
    lErro = objTelaPedido.Move_Pedido_Memoria(objPVTela)
    If lErro <> SUCESSO Then
        bErroNoRegistro = True
        Call Log_Adiciona("Não conseguiu mover a tela para memória.", objLogBase)
    End If
    
    If bNovoPedido Then
    
        objPVTela.sPedidoRepresentante = objIntegracaoPV.sPedRepr
    
        If objIntegracaoPV.iTrazerPesoAuto = DESMARCADO Then
            'Pega os pesos do registro de importação
            objPVTela.dPesoBruto = objIntegracaoPV.dPesoBruto
            objPVTela.dPesoLiq = objIntegracaoPV.dPesoLiquido
        
        End If
        objPVTela.sMensagemNota = objPV.sMensagemNota
        
        Set colComissaoAux = New Collection
        For Each objComissaoAux In objPVTela.colComissoes
            Set objComissao = New ClassComissaoPedVendas
            objComissao.iCodVendedor = objComissaoAux.iCodVendedor
            If iVendedor <> 0 And objComissao.iCodVendedor <> iVendedor Then
                objComissao.iCodVendedor = iVendedor
                objPVTela.iComissaoAutomatica = DESMARCADO
            End If
            objComissao.dPercentual = objComissaoAux.dPercentual
            objComissao.dPercentualEmissao = objComissaoAux.dPercentualEmissao
            objComissao.dValor = objComissaoAux.dValor
            objComissao.dValorBase = objComissaoAux.dValorBase
            objComissao.dValorEmissao = objComissaoAux.dValorEmissao
            objComissao.iIndireta = objComissaoAux.iIndireta
            colComissaoAux.Add objComissao
        Next
        Set objPVTela.colComissoes = colComissaoAux
        
        For iIndice = objPVTela.colComissoes.Count To 1 Step -1
            Set objComissao = objPVTela.colComissoes.Item(iIndice)
            
            If objComissao.iCodVendedor <> 0 Then
                If objComissao.dValor < 0.01 Then objComissao.dValor = 0.01
            Else
                objPVTela.colComissoes.Remove iIndice
            End If
        Next
        
        lErro = objTelaPedido.objTabComissoes.Carrega_Tab_Comissoes(objPVTela)
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Não conseguiu trazer a comissão para tela.", objLogBase)
        End If
        
        iSeq = 0
        For Each objIntegracaoPV In colIntegracaoPV
            iSeq = iSeq + 1
            Set objItemPV = objPVTela.colItensPedido.Item(iSeq)
            
            objItemPV.iTabelaPreco = objIntegracaoPV.iCodTabelaPrecoItem
            
            bAchouAlm = False
            If objIntegracaoPV.iCodAlmoxarifado = 0 Then
                    
                If objIntegracaoPV.iCodAlmoxarifadoExterno <> 0 Then
    
                    lErro = Comando_Executar(lComando9, "SELECT Codigo FROM Almoxarifado WHERE CodExterno = ? ", iAlmoxarifado, objIntegracaoPV.iCodAlmoxarifadoExterno)
                    If lErro <> AD_SQL_SUCESSO Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Erro na leitura do almoxarifado com código externo " & CStr(objIntegracaoPV.iCodAlmoxarifadoExterno) & ".", objLogBase)
                    End If
    
                    lErro = Comando_BuscarProximo(lComando9)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Erro na leitura do almoxarifado com código externo " & CStr(objIntegracaoPV.iCodAlmoxarifadoExterno) & ".", objLogBase)
                    End If
    
                    If lErro = AD_SQL_SUCESSO Then
                        bAchouAlm = True
                        Call Log_Adiciona("O almoxarifado foi localizado pelo código externo. " & "Almoxarifado: " & CStr(iAlmoxarifado) & ".", objLogBase, True)
                    Else
                        Call Log_Adiciona("O almoxarifado não foi localizado pelo código externo.", objLogBase, True)
                    End If
    
                End If
                
                If Len(Trim(objIntegracaoPV.sNomeAlmoxarifadoExterno)) <> 0 And Not bAchouAlm Then
    
                    lErro = Comando_Executar(lComando10, "SELECT Codigo FROM Almoxarifado WHERE NomeReduzido = ? ", iAlmoxarifado, objIntegracaoPV.sNomeAlmoxarifadoExterno)
                    If lErro <> AD_SQL_SUCESSO Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Erro na leitura do almoxarifado com nome reduzido " & CStr(objIntegracaoPV.sNomeAlmoxarifadoExterno) & ".", objLogBase)
                    End If
    
                    lErro = Comando_BuscarProximo(lComando10)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                        bErroNoRegistro = True
                        Call Log_Adiciona("Erro na leitura do almoxarifado com nome reduzido " & CStr(objIntegracaoPV.sNomeAlmoxarifadoExterno) & ".", objLogBase)
                    End If
    
                    If lErro = AD_SQL_SUCESSO Then
                        bAchouAlm = True
                        Call Log_Adiciona("O Almoxarifado foi localizado pelo nome reduzido. " & "Almoxarifado: " & CStr(iAlmoxarifado) & ".", objLogBase, True)
                    Else
                        Call Log_Adiciona("O Almoxarifado não foi localizado pelo nome reduzido.", objLogBase, True)
                    End If
    
                End If
                
            Else
                iAlmoxarifado = objIntegracaoPV.iCodAlmoxarifado
            End If

            If iAlmoxarifado <> 0 Then
                For Each objReservaItem In objItemPV.ColReserva
                    objReservaItem.iAlmoxarifado = iAlmoxarifado
                Next
            End If
            
        Next
        
        Set objIntegracaoPV = colIntegracaoPV.Item(1)
                
        lErro = objTelaPedido.Gravar_Registro(True)
        If lErro <> SUCESSO Then
            bErroNoRegistro = True
            Call Log_Adiciona("Ocorreram erros na validação das informações. Verificar o arquivo AdLog.txt para mais detalhes.", objLogBase)
        End If
        
    End If
    
    'If Not bNovoPedido Then
    If Not bCancelamento Then
    
        'VERIFICAR AS QUANTIDADE CANCELADAS PARA VER SE O PEDIDO FOI CANCELADO
        'SEGUNDO O SASCHA É O ÚNICO TIPO DE ALTERAÇÃO POSSÍVEL
        iSeq = 0
        For Each objIntegracaoPV In colIntegracaoPV
            iSeq = iSeq + 1
            Set objItemPV = objPVTela.colItensPedido.Item(iSeq)
            objItemPV.dQuantCancelada = objIntegracaoPV.dQuantidadeCancelada
        Next
    
    End If
    'End If
    
'    'Verifica se todos os itens estão cancelados
'    bCancelamento = True
'    For Each objItemPV In objPV.colItensPedido
'        If objItemPV.dQuantidade - objItemPV.dQuantCancelada > QTDE_ESTOQUE_DELTA Then
'            bCancelamento = False
'            Exit For
'        End If
'    Next
'
'    If bCancelamento Then
'        objPV.dValorDesconto = 0
'    End If
    
    Importa_Dados_PV_Preenche_PV = SUCESSO
    
    Exit Function

Erro_Importa_Dados_PV_Preenche_PV:

    Importa_Dados_PV_Preenche_PV = gErr

    Select Case gErr
    
        Case 190058

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190059)

    End Select
       
    Exit Function

End Function

Private Function Importacao_Inicia(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objTipoArq As ClassTipoArqIntegracao) As Long
'Calcula o total de registros que vai ser atualizado e acerta a tela de acompanhamento

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntArq As Long
Dim lTotalRegAux As Long
Dim lTotalRegistros As Long

On Error GoTo Erro_Importacao_Inicia
   
    TelaAcompanhaBatch.Log = TelaAcompanhaBatch.Log & "Início da importação de " & objTipoArq.sDescricao & vbNewLine
    
    'Le os arquivos que vão ser atualizados
    lErro = Comando_Executar(lComando1, "SELECT NumIntArq FROM ArqImportacao WHERE TipoArq = ? AND DataAtualizacao = ? ORDER BY NumIntArq", lNumIntArq, objTipoArq.iCodigo, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 189965
    
    'Busca o primeiro
    lErro = Comando_BuscarProximo(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189966
    
    'Enquando existirem arquivos não atualizados
    Do While lErro = AD_SQL_SUCESSO
    
        'Le o total de registros não atualizados desse arquivo
        lErro = Comando_Executar(lComando2, "SELECT Count(*) FROM " & objTipoArq.sTabela & " WHERE NumIntArq = ? AND DataAtualizacao = ? ", lTotalRegAux, lNumIntArq, DATA_NULA)
        If lErro <> AD_SQL_SUCESSO Then gError 189967
        
        'Busca
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189968
        
        'Totaliza os registros não atualizados
        lTotalRegistros = lTotalRegistros + lTotalRegAux

        'Busca o próximo registro
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189969

    Loop
    
    TelaAcompanhaBatch.NomeArqAtu = objTipoArq.sDescricao
    TelaAcompanhaBatch.dValorTotalAtu = lTotalRegistros
    TelaAcompanhaBatch.dValorAtualAtu = 0
    
    If lTotalRegistros = 0 Then
        
        'Atualiza a barra de progressão
        lErro = Rotina_Atualizar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 189970
    
    End If

    Importacao_Inicia = SUCESSO

    Exit Function

Erro_Importacao_Inicia:

    Importacao_Inicia = gErr

    Select Case gErr
    
        Case 189965 To 189969
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQIMPORTACAO", gErr)
            
        Case 189970

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189971)

    End Select

    Exit Function

End Function

Public Function Importa_Dados_Customizado(ByVal objImportacaoAux As ClassArqImportacaoAux) As Long

Dim lErro As Long

On Error GoTo Erro_Importa_Dados_Customizado

    If gobjFAT.sDirImportPVSOC <> "" Then
        lErro = Importa_Dados_Cust_SOC(gobjFAT.sDirImportPVSOC)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
            
    Importa_Dados_Customizado = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_Customizado:

    Importa_Dados_Customizado = gErr
     
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208905)
     
    End Select
    
    Exit Function
    
End Function

Public Function Importa_Dados_Cust_SOC(ByVal sDiretorio As String) As Long
        
Dim lErro As Long
Dim iIndice As Integer, lTransacao As Long
Dim alComando(1 To 7) As Long
Dim lCliente As Long, iFilial As Integer, sCodCli As String, sCodPV As String
Dim objFolder As Folder
Dim objFSO As New FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim objIntegracaoPVAux As ClassIntegracaoPV
Dim objIntegracaoPV As ClassIntegracaoPV
Dim sTipoReg As String, sRegistro As String, lSeq As Long, iItem As Integer
Dim objArqImportacao As ClassArqImportacao, dValorItens As Double, dValorFats As Double
Dim objProduto As ClassProduto, bComErro As Boolean, bArqAberto As Boolean
Dim dValorIR As Double, dValorIRParc As Double, dValorIRFaltaAplic As Double, iNumParcelas As Integer
Dim dValorParc As Double

On Error GoTo Erro_Importa_Dados_Cust_SOC

    bComErro = False
    bArqAberto = False
    
    'Pega todos os aquivos da pasta da empresa
    Set objFolder = objFSO.GetFolder(sDiretorio)
    
    'Para cada arquivo da empresa
    For Each objFile In objFolder.Files

        'Abertura de transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 208906

        'abertura de comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 208907
        Next
        
        Set objArqImportacao = New ClassArqImportacao
        
        objArqImportacao.iTipoArq = TIPO_ARQ_PV
        objArqImportacao.sNomeArquivo = Format(Now, "yyyymmddhhmmss") & "_" & objFile.Name
    
        lErro = ArqImportacao_Insere(alComando(3), objArqImportacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        bArqAberto = False
        Set objTS = objFile.OpenAsTextStream(ForReading)
        bArqAberto = True
        
        lSeq = 0
        Do While Not objTS.AtEndOfLine
        
            sRegistro = objTS.ReadLine
            
            sTipoReg = left(sRegistro, 1)
        
            Set objIntegracaoPV = New ClassIntegracaoPV

            If sTipoReg = "P" Then
            
                Set objIntegracaoPVAux = New ClassIntegracaoPV
                
                sCodCli = SOC_Campo(sRegistro, 3)
                sCodPV = SOC_Campo(sRegistro, 6)
                
                'Busca o cliente pelo código de integração
                lErro = Comando_Executar(alComando(4), "SELECT Cliente, Filial FROM FilialClienteCategorias WHERE Categoria = ? AND Item = ?", lCliente, iFilial, "Código Integração", sCodCli)
                If lErro <> AD_SQL_SUCESSO Then gError 209515
                
                'Busca o primeiro
                lErro = Comando_BuscarProximo(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209516
                
                If lErro <> AD_SQL_SUCESSO Then
                    bComErro = True
                    Call Rotina_Erro(vbOKOnly, "ERRO_SOC_COD_CLI_NAO_LOCALIZADO", 209517, sCodCli, sCodPV)
                    'gError 209517
                End If
            
                objIntegracaoPVAux.lCodClienteCorporator = lCliente 'StrParaLong(SOC_Campo(sRegistro, 3))
                objIntegracaoPVAux.iCodFilialCorporator = iFilial
                objIntegracaoPVAux.dtDataEmissao = StrParaDate(SOC_Campo(sRegistro, 5))
                objIntegracaoPVAux.lCodPedidoExterno = StrParaLong(Mid(SOC_Campo(sRegistro, 6), 4))
            
                iItem = 0
                dValorItens = 0
                dValorFats = 0
                dValorIR = 0
                iNumParcelas = 0
            
            ElseIf sTipoReg = "I" Then
                
                lSeq = lSeq + 1
                iItem = iItem + 1
                
                objIntegracaoPV.lNumIntArq = objArqImportacao.lNumIntArq
                objIntegracaoPV.lSeqRegistro = lSeq
                objIntegracaoPV.iTipoInt = TIPO_INTEGRACAO_IMPORTACAO
                
                objIntegracaoPV.lCodClienteCorporator = objIntegracaoPVAux.lCodClienteCorporator
                objIntegracaoPV.iCodFilialCorporator = objIntegracaoPVAux.iCodFilialCorporator 'FILIAL_MATRIZ
                objIntegracaoPV.dtDataEmissao = objIntegracaoPVAux.dtDataEmissao
                objIntegracaoPV.lCodPedidoExterno = objIntegracaoPVAux.lCodPedidoExterno
                
                objIntegracaoPV.sCodProdutoExterno = SOC_Campo(sRegistro, 2)
                objIntegracaoPV.sCodProduto = SOC_Campo(sRegistro, 2)
                objIntegracaoPV.sCodProduto = objIntegracaoPV.sCodProduto & String(10 - Len(objIntegracaoPV.sCodProduto), " ")
                objIntegracaoPV.sNaturezaOp = "" 'SOC_Campo(sRegistro, 3)
                objIntegracaoPV.dQuantidadePedida = StrParaDbl(SOC_Campo(sRegistro, 4))
                objIntegracaoPV.dPrecoUnitario = StrParaDbl(SOC_Campo(sRegistro, 5))
                
                Set objProduto = New ClassProduto
                objProduto.sCodigo = objIntegracaoPV.sCodProduto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                If lErro <> SUCESSO Then
                    bComErro = True
                    Call Rotina_Erro(vbOKOnly, "ERRO_SOC_COD_PROD_NAO_LOCALIZADO", 209519, Trim(objProduto.sCodigo), sCodPV)
                    'gError 209519
                End If
                
                dValorItens = dValorItens + Arredonda_Moeda(objIntegracaoPV.dQuantidadePedida * objIntegracaoPV.dPrecoUnitario)
                
                objIntegracaoPV.dtDataEntrega = DATA_NULA
                
                objIntegracaoPV.iItem = iItem
                objIntegracaoPV.iTrazerDescricaoAuto = MARCADO
                
                If Not bComErro Then
                    lErro = IntegracaoPV_Insere(alComando(1), objIntegracaoPV)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                End If
                
            ElseIf sTipoReg = "F" Then
            
                iNumParcelas = iNumParcelas + 1
            
                If Not bComErro Then
                    lErro = Comando_Executar(alComando(2), "INSERT INTO IntegracaoPVParc(NumIntArq, Pedido, FilialEmpresa, Parcela, Vencimento, Valor) VALUES (?,?,?,?,?,?) ", _
                    objArqImportacao.lNumIntArq, objIntegracaoPVAux.lCodPedidoExterno, giFilialEmpresa, StrParaInt(SOC_Campo(sRegistro, 1)), StrParaDate(SOC_Campo(sRegistro, 2)), StrParaDbl(SOC_Campo(sRegistro, 3)))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208908
                End If
                
                dValorFats = dValorFats + StrParaDbl(SOC_Campo(sRegistro, 3))
                
            ElseIf sTipoReg = "C" Then
                
                If Abs(dValorItens - dValorFats) > 1 Then
                    bComErro = True
                    Call Rotina_Erro(vbOKOnly, "ERRO_SOC_COD_VALOR_DIF", 209518, Format(dValorItens, "STANDARD"), Format(dValorFats, "STANDARD"), sCodPV)
                    'gError 209518
                End If
                
                dValorIR = Arredonda_Moeda(dValorItens * 0.015)
                If dValorIR < 10 Then dValorIR = 0
                dValorIRFaltaAplic = dValorIR
                
                If dValorIR > 0 Then
                
                    lErro = Comando_ExecutarPos(alComando(6), "SELECT Valor FROM IntegracaoPVParc WHERE NumIntArq =? AND Pedido = ? AND FilialEmpresa = ? ORDER BY Parcela", 0, dValorParc, objArqImportacao.lNumIntArq, objIntegracaoPVAux.lCodPedidoExterno, giFilialEmpresa)
                    If lErro <> AD_SQL_SUCESSO Then gError 209515
                    
                    'Busca o primeiro
                    lErro = Comando_BuscarProximo(alComando(6))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209516
                                        
                    iIndice = 0
                    Do While lErro <> AD_SQL_SEM_DADOS
                    
                        iIndice = iIndice + 1
                    
                        If iIndice = iNumParcelas Then
                            dValorIRParc = dValorIRFaltaAplic
                        Else
                            dValorIRParc = Arredonda_Moeda(dValorIR / iNumParcelas)
                        End If
                        dValorIRFaltaAplic = dValorIRFaltaAplic - dValorIRParc
                    
                        lErro = Comando_ExecutarPos(alComando(7), "UPDATE IntegracaoPVParc SET Valor = Valor - ? ", alComando(6), dValorIRParc)
                        If lErro <> AD_SQL_SUCESSO Then gError 209515
                    
                        lErro = Comando_BuscarProximo(alComando(6))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209516
                    
                    Loop
                
                
                End If
                
                
            End If
            
        Loop
        
        Call objTS.Close
        bArqAberto = False
        
        If bComErro Then gError ERRO_SEM_MENSAGEM
        
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
                
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 208909
        
        'Copia para pasta de processados e apaga o anterior
        Call FileCopy(sDiretorio & objFile.Name, sDiretorio & "Proc\" & Format(Now, "yyyymmddhhmmss") & "_" & objFile.Name)
         
        Call Kill(sDiretorio & objFile.Name)
        
    Next
        
    Importa_Dados_Cust_SOC = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_Cust_SOC:

    Importa_Dados_Cust_SOC = gErr
     
    Select Case gErr
    
        Case 208906
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                
        Case 208907
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208908
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INTEGRACAOPV", gErr)
        
        Case 208909
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 209515, 209516
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS", gErr)
        
        Case 209517
            Call Rotina_Erro(vbOKOnly, "ERRO_SOC_COD_CLI_NAO_LOCALIZADO", gErr, sCodCli, sCodPV)

        Case 209518
            Call Rotina_Erro(vbOKOnly, "ERRO_SOC_COD_VALOR_DIF", gErr, Format(dValorItens, "STANDARD"), Format(dValorFats, "STANDARD"), sCodPV)

        Case 209519
            Call Rotina_Erro(vbOKOnly, "ERRO_SOC_COD_PROD_NAO_LOCALIZADO", gErr, Trim(objProduto.sCodigo), sCodPV)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208910)
     
    End Select
    
    If bArqAberto Then objTS.Close
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function SOC_Campo(ByVal sRegistro As String, ByVal iCampo As Integer) As String

Dim iPos1 As Integer
Dim iPos2 As Integer
Dim iIndice As Integer
Dim sAux As String

    iPos2 = 0
    iPos1 = InStr(iPos2 + 1, sRegistro, ";")
    For iIndice = 1 To iCampo - 1
        iPos2 = iPos1
        iPos1 = InStr(iPos2 + 1, sRegistro, ";")
        If iPos1 = 0 Then iPos1 = Len(sRegistro) + 1
    Next
    sAux = Mid(sRegistro, iPos2 + 1, iPos1 - iPos2 - 1)
    If iCampo = 1 Then sAux = Mid(sAux, 2)
    SOC_Campo = sAux

End Function

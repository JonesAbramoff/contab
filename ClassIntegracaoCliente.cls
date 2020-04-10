VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIntegracaoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntGer As Long
Private mvarlNumIntArq As Long
Private mvarlSeqRegistro As Long
Private mvariTipoInt As Integer
Private mvarlCodCliente As Long
Private mvariCodFilial As Integer
Private mvarsRazaoSocial As String
Private mvarsNomeReduzido As String
Private mvariTipo As Integer
Private mvarsObservacao As String
Private mvardLimiteCredito As Double
Private mvariCondicaoPagto As Integer
Private mvariAtivo As Integer
Private mvarsFilialNome As String
Private mvarsCGC As String
Private mvarsRG As String
Private mvarsInscricaoEstadual As String
Private mvarsInscricaoMunicipal As String
Private mvarsEndereco As String
Private mvarsBairro As String
Private mvarsCidade As String
Private mvarsSiglaEstado As String
Private mvariCodigoPais As Integer
Private mvarsCEP As String
Private mvarsTelefone1 As String
Private mvarsTelefone2 As String
Private mvarsEmail As String
Private mvarsFax As String
Private mvarsContato As String
Private mvarsEnderecoEnt As String
Private mvarsBairroEnt As String
Private mvarsCidadeEnt As String
Private mvarsSiglaEstadoEnt As String
Private mvariCodigoPaisEnt As Integer
Private mvarsCEPEnt As String
Private mvarsTelefone1Ent As String
Private mvarsTelefone2Ent As String
Private mvarsEmailEnt As String
Private mvarsFaxEnt As String
Private mvarsContatoEnt As String
Private mvarsEnderecoCobr As String
Private mvarsBairroCobr As String
Private mvarsCidadeCobr As String
Private mvarsSiglaEstadoCobr As String
Private mvariCodigoPaisCobr As Integer
Private mvarsCEPCobr As String
Private mvarsTelefone1Cobr As String
Private mvarsTelefone2Cobr As String
Private mvarsEmailCobr As String
Private mvarsFaxCobr As String
Private mvarsContatoCobr As String
Private mvariComErro As Integer
Private mvardtDataAtualizacao As Date
Private mvariVendedor As Integer
Private mvarsObservacaoFilial As String
Private mvarlCodExterno As Long
Private mvarsTipoCliente As String
Private mvarsLogradouro As String
Private mvarsComplemento As String
Private mvarsTipoLogradouro As String
Private mvarsEmail2 As String
Private mvarlNumero As Long
Private mvariTelDDD1 As Integer
Private mvariTelDDD2 As Integer
Private mvariFaxDDD As Integer
Private mvarsTelNumero1 As String
Private mvarsTelNumero2 As String
Private mvarsFaxNumero As String
Private mvarsLogradouroEnt As String
Private mvarsComplementoEnt As String
Private mvarsTipoLogradouroEnt As String
Private mvarsEmail2Ent As String
Private mvarlNumeroEnt As Long
Private mvariTelDDD1Ent As Integer
Private mvariTelDDD2Ent As Integer
Private mvariFaxDDDEnt As Integer
Private mvarsTelNumero1Ent As String
Private mvarsTelNumero2Ent As String
Private mvarsFaxNumeroEnt As String
Private mvariTabelaPreco As Integer
Private mvarsReferenciaEnt As String
Private mvarsReferencia As String

Public Property Let sReferencia(ByVal vData As String)
    mvarsReferencia = vData
End Property

Public Property Get sReferencia() As String
    sReferencia = mvarsReferencia
End Property

Public Property Let sReferenciaEnt(ByVal vData As String)
    mvarsReferenciaEnt = vData
End Property

Public Property Get sReferenciaEnt() As String
    sReferenciaEnt = mvarsReferenciaEnt
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let sLogradouro(ByVal vData As String)
    mvarsLogradouro = vData
End Property

Public Property Get sLogradouro() As String
    sLogradouro = mvarsLogradouro
End Property

Public Property Let sComplemento(ByVal vData As String)
    mvarsComplemento = vData
End Property

Public Property Get sComplemento() As String
    sComplemento = mvarsComplemento
End Property

Public Property Let sTipoLogradouro(ByVal vData As String)
    mvarsTipoLogradouro = vData
End Property

Public Property Get sTipoLogradouro() As String
    sTipoLogradouro = mvarsTipoLogradouro
End Property

Public Property Let sEmail2(ByVal vData As String)
    mvarsEmail2 = vData
End Property

Public Property Get sEmail2() As String
    sEmail2 = mvarsEmail2
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let iTelDDD1(ByVal vData As Integer)
    mvariTelDDD1 = vData
End Property

Public Property Get iTelDDD1() As Integer
    iTelDDD1 = mvariTelDDD1
End Property

Public Property Let iTelDDD2(ByVal vData As Integer)
    mvariTelDDD2 = vData
End Property

Public Property Get iTelDDD2() As Integer
    iTelDDD2 = mvariTelDDD2
End Property

Public Property Let iFaxDDD(ByVal vData As Integer)
    mvariFaxDDD = vData
End Property

Public Property Get iFaxDDD() As Integer
    iFaxDDD = mvariFaxDDD
End Property

Public Property Let sTelNumero1(ByVal vData As String)
    mvarsTelNumero1 = vData
End Property

Public Property Get sTelNumero1() As String
    sTelNumero1 = mvarsTelNumero1
End Property

Public Property Let sTelNumero2(ByVal vData As String)
    mvarsTelNumero2 = vData
End Property

Public Property Get sTelNumero2() As String
    sTelNumero2 = mvarsTelNumero2
End Property

Public Property Let sFaxNumero(ByVal vData As String)
    mvarsFaxNumero = vData
End Property

Public Property Get sFaxNumero() As String
    sFaxNumero = mvarsFaxNumero
End Property

Public Property Let sLogradouroEnt(ByVal vData As String)
    mvarsLogradouroEnt = vData
End Property

Public Property Get sLogradouroEnt() As String
    sLogradouroEnt = mvarsLogradouroEnt
End Property

Public Property Let sComplementoEnt(ByVal vData As String)
    mvarsComplementoEnt = vData
End Property

Public Property Get sComplementoEnt() As String
    sComplementoEnt = mvarsComplementoEnt
End Property

Public Property Let sTipoLogradouroEnt(ByVal vData As String)
    mvarsTipoLogradouroEnt = vData
End Property

Public Property Get sTipoLogradouroEnt() As String
    sTipoLogradouroEnt = mvarsTipoLogradouroEnt
End Property

Public Property Let sEmail2Ent(ByVal vData As String)
    mvarsEmail2Ent = vData
End Property

Public Property Get sEmail2Ent() As String
    sEmail2Ent = mvarsEmail2Ent
End Property

Public Property Let lNumeroEnt(ByVal vData As Long)
    mvarlNumeroEnt = vData
End Property

Public Property Get lNumeroEnt() As Long
    lNumeroEnt = mvarlNumeroEnt
End Property

Public Property Let iTelDDD1Ent(ByVal vData As Integer)
    mvariTelDDD1Ent = vData
End Property

Public Property Get iTelDDD1Ent() As Integer
    iTelDDD1Ent = mvariTelDDD1Ent
End Property

Public Property Let iTelDDD2Ent(ByVal vData As Integer)
    mvariTelDDD2Ent = vData
End Property

Public Property Get iTelDDD2Ent() As Integer
    iTelDDD2Ent = mvariTelDDD2Ent
End Property

Public Property Let iFaxDDDEnt(ByVal vData As Integer)
    mvariFaxDDDEnt = vData
End Property

Public Property Get iFaxDDDEnt() As Integer
    iFaxDDDEnt = mvariFaxDDDEnt
End Property

Public Property Let sTelNumero1Ent(ByVal vData As String)
    mvarsTelNumero1Ent = vData
End Property

Public Property Get sTelNumero1Ent() As String
    sTelNumero1Ent = mvarsTelNumero1Ent
End Property

Public Property Let sTelNumero2Ent(ByVal vData As String)
    mvarsTelNumero2Ent = vData
End Property

Public Property Get sTelNumero2Ent() As String
    sTelNumero2Ent = mvarsTelNumero2Ent
End Property

Public Property Let sFaxNumeroEnt(ByVal vData As String)
    mvarsFaxNumeroEnt = vData
End Property

Public Property Get sFaxNumeroEnt() As String
    sFaxNumeroEnt = mvarsFaxNumeroEnt
End Property

Public Property Let sTipoCliente(ByVal vData As String)
    mvarsTipoCliente = vData
End Property

Public Property Get sTipoCliente() As String
    sTipoCliente = mvarsTipoCliente
End Property

Public Property Let lCodExterno(ByVal vData As Long)
    mvarlCodExterno = vData
End Property

Public Property Get lCodExterno() As Long
    lCodExterno = mvarlCodExterno
End Property

Public Property Let sObservacaoFilial(ByVal vData As String)
    mvarsObservacaoFilial = vData
End Property

Public Property Get sObservacaoFilial() As String
    sObservacaoFilial = mvarsObservacaoFilial
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntGer(ByVal vData As Long)
    mvarlNumIntGer = vData
End Property

Public Property Get lNumIntGer() As Long
    lNumIntGer = mvarlNumIntGer
End Property

Public Property Let lNumIntArq(ByVal vData As Long)
    mvarlNumIntArq = vData
End Property

Public Property Get lNumIntArq() As Long
    lNumIntArq = mvarlNumIntArq
End Property

Public Property Let lSeqRegistro(ByVal vData As Long)
    mvarlSeqRegistro = vData
End Property

Public Property Get lSeqRegistro() As Long
    lSeqRegistro = mvarlSeqRegistro
End Property

Public Property Let iTipoInt(ByVal vData As Integer)
    mvariTipoInt = vData
End Property

Public Property Get iTipoInt() As Integer
    iTipoInt = mvariTipoInt
End Property

Public Property Let lCodCliente(ByVal vData As Long)
    mvarlCodCliente = vData
End Property

Public Property Get lCodCliente() As Long
    lCodCliente = mvarlCodCliente
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

Public Property Let sRazaoSocial(ByVal vData As String)
    mvarsRazaoSocial = vData
End Property

Public Property Get sRazaoSocial() As String
    sRazaoSocial = mvarsRazaoSocial
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dLimiteCredito(ByVal vData As Double)
    mvardLimiteCredito = vData
End Property

Public Property Get dLimiteCredito() As Double
    dLimiteCredito = mvardLimiteCredito
End Property

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let sFilialNome(ByVal vData As String)
    mvarsFilialNome = vData
End Property

Public Property Get sFilialNome() As String
    sFilialNome = mvarsFilialNome
End Property

Public Property Let sCgc(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let sRG(ByVal vData As String)
    mvarsRG = vData
End Property

Public Property Get sRG() As String
    sRG = mvarsRG
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let sInscricaoMunicipal(ByVal vData As String)
    mvarsInscricaoMunicipal = vData
End Property

Public Property Get sInscricaoMunicipal() As String
    sInscricaoMunicipal = mvarsInscricaoMunicipal
End Property

Public Property Let sEndereco(ByVal vData As String)
    mvarsEndereco = vData
End Property

Public Property Get sEndereco() As String
    sEndereco = mvarsEndereco
End Property

Public Property Let sBairro(ByVal vData As String)
    mvarsBairro = vData
End Property

Public Property Get sBairro() As String
    sBairro = mvarsBairro
End Property

Public Property Let sCidade(ByVal vData As String)
    mvarsCidade = vData
End Property

Public Property Get sCidade() As String
    sCidade = mvarsCidade
End Property

Public Property Let sSiglaEstado(ByVal vData As String)
    mvarsSiglaEstado = vData
End Property

Public Property Get sSiglaEstado() As String
    sSiglaEstado = mvarsSiglaEstado
End Property

Public Property Let iCodigoPais(ByVal vData As Integer)
    mvariCodigoPais = vData
End Property

Public Property Get iCodigoPais() As Integer
    iCodigoPais = mvariCodigoPais
End Property

Public Property Let sCEP(ByVal vData As String)
    mvarsCEP = vData
End Property

Public Property Get sCEP() As String
    sCEP = mvarsCEP
End Property

Public Property Let sTelefone1(ByVal vData As String)
    mvarsTelefone1 = vData
End Property

Public Property Get sTelefone1() As String
    sTelefone1 = mvarsTelefone1
End Property

Public Property Let sTelefone2(ByVal vData As String)
    mvarsTelefone2 = vData
End Property

Public Property Get sTelefone2() As String
    sTelefone2 = mvarsTelefone2
End Property

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Let sFax(ByVal vData As String)
    mvarsFax = vData
End Property

Public Property Get sFax() As String
    sFax = mvarsFax
End Property

Public Property Let sContato(ByVal vData As String)
    mvarsContato = vData
End Property

Public Property Get sContato() As String
    sContato = mvarsContato
End Property

Public Property Let sEnderecoEnt(ByVal vData As String)
    mvarsEnderecoEnt = vData
End Property

Public Property Get sEnderecoEnt() As String
    sEnderecoEnt = mvarsEnderecoEnt
End Property

Public Property Let sBairroEnt(ByVal vData As String)
    mvarsBairroEnt = vData
End Property

Public Property Get sBairroEnt() As String
    sBairroEnt = mvarsBairroEnt
End Property

Public Property Let sCidadeEnt(ByVal vData As String)
    mvarsCidadeEnt = vData
End Property

Public Property Get sCidadeEnt() As String
    sCidadeEnt = mvarsCidadeEnt
End Property

Public Property Let sSiglaEstadoEnt(ByVal vData As String)
    mvarsSiglaEstadoEnt = vData
End Property

Public Property Get sSiglaEstadoEnt() As String
    sSiglaEstadoEnt = mvarsSiglaEstadoEnt
End Property

Public Property Let iCodigoPaisEnt(ByVal vData As Integer)
    mvariCodigoPaisEnt = vData
End Property

Public Property Get iCodigoPaisEnt() As Integer
    iCodigoPaisEnt = mvariCodigoPaisEnt
End Property

Public Property Let sCEPEnt(ByVal vData As String)
    mvarsCEPEnt = vData
End Property

Public Property Get sCEPEnt() As String
    sCEPEnt = mvarsCEPEnt
End Property

Public Property Let sTelefone1Ent(ByVal vData As String)
    mvarsTelefone1Ent = vData
End Property

Public Property Get sTelefone1Ent() As String
    sTelefone1Ent = mvarsTelefone1Ent
End Property

Public Property Let sTelefone2Ent(ByVal vData As String)
    mvarsTelefone2Ent = vData
End Property

Public Property Get sTelefone2Ent() As String
    sTelefone2Ent = mvarsTelefone2Ent
End Property

Public Property Let sEmailEnt(ByVal vData As String)
    mvarsEmailEnt = vData
End Property

Public Property Get sEmailEnt() As String
    sEmailEnt = mvarsEmailEnt
End Property

Public Property Let sFaxEnt(ByVal vData As String)
    mvarsFaxEnt = vData
End Property

Public Property Get sFaxEnt() As String
    sFaxEnt = mvarsFaxEnt
End Property

Public Property Let sContatoEnt(ByVal vData As String)
    mvarsContatoEnt = vData
End Property

Public Property Get sContatoEnt() As String
    sContatoEnt = mvarsContatoEnt
End Property

Public Property Let sEnderecoCobr(ByVal vData As String)
    mvarsEnderecoCobr = vData
End Property

Public Property Get sEnderecoCobr() As String
    sEnderecoCobr = mvarsEnderecoCobr
End Property

Public Property Let sBairroCobr(ByVal vData As String)
    mvarsBairroCobr = vData
End Property

Public Property Get sBairroCobr() As String
    sBairroCobr = mvarsBairroCobr
End Property

Public Property Let sCidadeCobr(ByVal vData As String)
    mvarsCidadeCobr = vData
End Property

Public Property Get sCidadeCobr() As String
    sCidadeCobr = mvarsCidadeCobr
End Property

Public Property Let sSiglaEstadoCobr(ByVal vData As String)
    mvarsSiglaEstadoCobr = vData
End Property

Public Property Get sSiglaEstadoCobr() As String
    sSiglaEstadoCobr = mvarsSiglaEstadoCobr
End Property

Public Property Let iCodigoPaisCobr(ByVal vData As Integer)
    mvariCodigoPaisCobr = vData
End Property

Public Property Get iCodigoPaisCobr() As Integer
    iCodigoPaisCobr = mvariCodigoPaisCobr
End Property

Public Property Let sCEPCobr(ByVal vData As String)
    mvarsCEPCobr = vData
End Property

Public Property Get sCEPCobr() As String
    sCEPCobr = mvarsCEPCobr
End Property

Public Property Let sTelefone1Cobr(ByVal vData As String)
    mvarsTelefone1Cobr = vData
End Property

Public Property Get sTelefone1Cobr() As String
    sTelefone1Cobr = mvarsTelefone1Cobr
End Property

Public Property Let sTelefone2Cobr(ByVal vData As String)
    mvarsTelefone2Cobr = vData
End Property

Public Property Get sTelefone2Cobr() As String
    sTelefone2Cobr = mvarsTelefone2Cobr
End Property

Public Property Let sEmailCobr(ByVal vData As String)
    mvarsEmailCobr = vData
End Property

Public Property Get sEmailCobr() As String
    sEmailCobr = mvarsEmailCobr
End Property

Public Property Let sFaxCobr(ByVal vData As String)
    mvarsFaxCobr = vData
End Property

Public Property Get sFaxCobr() As String
    sFaxCobr = mvarsFaxCobr
End Property

Public Property Let sContatoCobr(ByVal vData As String)
    mvarsContatoCobr = vData
End Property

Public Property Get sContatoCobr() As String
    sContatoCobr = mvarsContatoCobr
End Property

Public Property Let iComErro(ByVal vData As Integer)
    mvariComErro = vData
End Property

Public Property Get iComErro() As Integer
    iComErro = mvariComErro
End Property

Public Property Let dtDataAtualizacao(ByVal vData As Date)
    mvardtDataAtualizacao = vData
End Property

Public Property Get dtDataAtualizacao() As Date
    dtDataAtualizacao = mvardtDataAtualizacao
End Property

Public Function Monta_Registro_Cliente(sRegistro As String)

Dim lErro As Long

On Error GoTo Erro_Monta_Registro_Cliente

    sRegistro = "01"

    sRegistro = sRegistro & FormataCpoNum(lCodCliente, 6) & FormataCpoNum(iCodFilial, 4) & FormataCpoTexto(sRazaoSocial, STRING_CLIENTE_RAZAO_SOCIAL) & FormataCpoTexto(sNomeReduzido, STRING_CLIENTE_NOME_REDUZIDO)
    sRegistro = sRegistro & FormataCpoNum(iTipo, 3) & FormataCpoTexto(sObservacao, STRING_CLIENTE_OBSERVACAO) & FormataCpoValor(dLimiteCredito, 12) & FormataCpoNum(iCondicaoPagto, 3) & FormataCpoNum(iAtivo, 1)
    sRegistro = sRegistro & FormataCpoTexto(sFilialNome, STRING_FILIAL_CLIENTE_NOME) & FormataCpoTexto(sCgc, STRING_CGC) & FormataCpoTexto(sRG, STRING_RG) & FormataCpoTexto(sInscricaoEstadual, STRING_INSCR_EST) & FormataCpoTexto(sInscricaoMunicipal, STRING_INSCR_MUN)
    
    sRegistro = sRegistro & FormataCpoTexto(sEndereco, STRING_ENDERECO) & FormataCpoTexto(sBairro, STRING_BAIRRO) & FormataCpoTexto(sCidade, STRING_CIDADE) & FormataCpoTexto(sSiglaEstado, STRING_ESTADO) & FormataCpoNum(iCodigoPais, 2) & FormataCpoTexto(sCEP, STRING_CEP) & FormataCpoTexto(sTelefone1, STRING_TELEFONE) & FormataCpoTexto(sTelefone2, STRING_TELEFONE) & FormataCpoTexto(sEmail, STRING_EMAIL) & FormataCpoTexto(sFax, STRING_FAX) & FormataCpoTexto(sContato, STRING_CONTATO)
    sRegistro = sRegistro & FormataCpoTexto(sEnderecoEnt, STRING_ENDERECO) & FormataCpoTexto(sBairroEnt, STRING_BAIRRO) & FormataCpoTexto(sCidadeEnt, STRING_CIDADE) & FormataCpoTexto(sSiglaEstadoEnt, STRING_ESTADO) & FormataCpoNum(iCodigoPaisEnt, 2) & FormataCpoTexto(sCEPEnt, STRING_CEP) & FormataCpoTexto(sTelefone1Ent, STRING_TELEFONE) & FormataCpoTexto(sTelefone2Ent, STRING_TELEFONE) & FormataCpoTexto(sEmailEnt, STRING_EMAIL) & FormataCpoTexto(sFaxEnt, STRING_FAX) & FormataCpoTexto(sContatoEnt, STRING_CONTATO)
    sRegistro = sRegistro & FormataCpoTexto(sEnderecoCobr, STRING_ENDERECO) & FormataCpoTexto(sBairroCobr, STRING_BAIRRO) & FormataCpoTexto(sCidadeCobr, STRING_CIDADE) & FormataCpoTexto(sSiglaEstadoCobr, STRING_ESTADO) & FormataCpoNum(iCodigoPaisCobr, 2) & FormataCpoTexto(sCEPCobr, STRING_CEP) & FormataCpoTexto(sTelefone1Cobr, STRING_TELEFONE) & FormataCpoTexto(sTelefone2Cobr, STRING_TELEFONE) & FormataCpoTexto(sEmailCobr, STRING_EMAIL) & FormataCpoTexto(sFaxCobr, STRING_FAX) & FormataCpoTexto(sContatoCobr, STRING_CONTATO)
    
    sRegistro = sRegistro & FormataCpoTexto(sObservacaoFilial, STRING_CLIENTE_OBSERVACAO) & FormataCpoNum(iVendedor, 4) & FormataCpoNum(lCodExterno, 6)

    Monta_Registro_Cliente = SUCESSO

    Exit Function

Erro_Monta_Registro_Cliente:

    Monta_Registro_Cliente = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189710)

    End Select

    Exit Function
    
End Function

Public Function Desmonta_Registro_Cliente(sRegistro As String)

Dim lErro As Long
Dim iPOS As Integer
On Error GoTo Erro_Desmonta_Registro_Cliente
    
    '00
    iPOS = 3
    
    lCodCliente = StrParaLong(Mid(sRegistro, iPOS, 6))
    iPOS = iPOS + 6
    
    iCodFilial = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    sRazaoSocial = Trim(Mid(sRegistro, iPOS, STRING_RAZAO_SOCIAL))
    iPOS = iPOS + STRING_CLIENTE_RAZAO_SOCIAL
    
    sNomeReduzido = Trim(Mid(sRegistro, iPOS, STRING_CLIENTE_NOME_REDUZIDO))
    iPOS = iPOS + STRING_CLIENTE_NOME_REDUZIDO
    
    iTipo = StrParaInt(Mid(sRegistro, iPOS, 3))
    iPOS = iPOS + 3
    
    sObservacao = Trim(Mid(sRegistro, iPOS, STRING_CLIENTE_OBSERVACAO))
    iPOS = iPOS + STRING_CLIENTE_OBSERVACAO
    
    dLimiteCredito = StrParaDbl(Mid(sRegistro, iPOS, 12)) / 100
    iPOS = iPOS + 12
    
    iCondicaoPagto = StrParaInt(Mid(sRegistro, iPOS, 3))
    iPOS = iPOS + 3
    
    iAtivo = StrParaInt(Mid(sRegistro, iPOS, 1))
    iPOS = iPOS + 1
    
    sFilialNome = Trim(Mid(sRegistro, iPOS, STRING_FILIAL_CLIENTE_NOME))
    iPOS = iPOS + STRING_FILIAL_CLIENTE_NOME
    
    sCgc = Trim(Mid(sRegistro, iPOS, STRING_CGC))
    iPOS = iPOS + STRING_CGC
    
    sRG = Trim(Mid(sRegistro, iPOS, STRING_RG))
    iPOS = iPOS + STRING_RG
    
    sInscricaoEstadual = Trim(Mid(sRegistro, iPOS, STRING_INSCR_EST))
    iPOS = iPOS + STRING_INSCR_EST
    
    sInscricaoMunicipal = Trim(Mid(sRegistro, iPOS, STRING_INSCR_MUN))
    iPOS = iPOS + STRING_INSCR_MUN
    
    sEndereco = Trim(Mid(sRegistro, iPOS, STRING_ENDERECO))
    iPOS = iPOS + STRING_ENDERECO
    
    sBairro = Trim(Mid(sRegistro, iPOS, STRING_BAIRRO))
    iPOS = iPOS + STRING_BAIRRO
    
    sCidade = Trim(Mid(sRegistro, iPOS, STRING_CIDADE))
    iPOS = iPOS + STRING_CIDADE
    
    sSiglaEstado = Trim(Mid(sRegistro, iPOS, STRING_ESTADO))
    iPOS = iPOS + STRING_ESTADO
    
    iCodigoPais = StrParaInt(Mid(sRegistro, iPOS, 2))
    iPOS = iPOS + 2
    
    sCEP = Trim(Mid(sRegistro, iPOS, STRING_CEP))
    iPOS = iPOS + STRING_CEP
    
    sTelefone1 = Trim(Mid(sRegistro, iPOS, STRING_TELEFONE))
    iPOS = iPOS + STRING_TELEFONE
    
    sTelefone2 = Trim(Mid(sRegistro, iPOS, STRING_TELEFONE))
    iPOS = iPOS + STRING_TELEFONE
    
    sEmail = Trim(Mid(sRegistro, iPOS, STRING_EMAIL))
    iPOS = iPOS + STRING_EMAIL
    
    sFax = Trim(Mid(sRegistro, iPOS, STRING_FAX))
    iPOS = iPOS + STRING_FAX
    
    sContato = Trim(Mid(sRegistro, iPOS, STRING_CONTATO))
    iPOS = iPOS + STRING_CONTATO
    
    sEnderecoEnt = Trim(Mid(sRegistro, iPOS, STRING_ENDERECO))
    iPOS = iPOS + STRING_ENDERECO
    
    sBairroEnt = Trim(Mid(sRegistro, iPOS, STRING_BAIRRO))
    iPOS = iPOS + STRING_BAIRRO
    
    sCidadeEnt = Trim(Mid(sRegistro, iPOS, STRING_CIDADE))
    iPOS = iPOS + STRING_CIDADE
    
    sSiglaEstadoEnt = Trim(Mid(sRegistro, iPOS, STRING_ESTADO))
    iPOS = iPOS + STRING_ESTADO
    
    iCodigoPaisEnt = StrParaInt(Mid(sRegistro, iPOS, 2))
    iPOS = iPOS + 2
    
    sCEPEnt = Trim(Mid(sRegistro, iPOS, STRING_CEP))
    iPOS = iPOS + STRING_CEP
    
    sTelefone1Ent = Trim(Mid(sRegistro, iPOS, STRING_TELEFONE))
    iPOS = iPOS + STRING_TELEFONE
    
    sTelefone2Ent = Trim(Mid(sRegistro, iPOS, STRING_TELEFONE))
    iPOS = iPOS + STRING_TELEFONE
    
    sEmailEnt = Trim(Mid(sRegistro, iPOS, STRING_EMAIL))
    iPOS = iPOS + STRING_EMAIL
    
    sFaxEnt = Trim(Mid(sRegistro, iPOS, STRING_FAX))
    iPOS = iPOS + STRING_FAX
    
    sContatoEnt = Trim(Mid(sRegistro, iPOS, STRING_CONTATO))
    iPOS = iPOS + STRING_CONTATO
    
    sEnderecoCobr = Trim(Mid(sRegistro, iPOS, STRING_ENDERECO))
    iPOS = iPOS + STRING_ENDERECO
    
    sBairroCobr = Trim(Mid(sRegistro, iPOS, STRING_BAIRRO))
    iPOS = iPOS + STRING_BAIRRO
    
    sCidadeCobr = Trim(Mid(sRegistro, iPOS, STRING_CIDADE))
    iPOS = iPOS + STRING_CIDADE
    
    sSiglaEstadoCobr = Trim(Mid(sRegistro, iPOS, STRING_ESTADO))
    iPOS = iPOS + STRING_ESTADO
    
    iCodigoPaisCobr = StrParaInt(Mid(sRegistro, iPOS, 2))
    iPOS = iPOS + 2
    
    sCEPCobr = Trim(Mid(sRegistro, iPOS, STRING_CEP))
    iPOS = iPOS + STRING_CEP
    
    sTelefone1Cobr = Trim(Mid(sRegistro, iPOS, STRING_TELEFONE))
    iPOS = iPOS + STRING_TELEFONE
    
    sTelefone2Cobr = Trim(Mid(sRegistro, iPOS, STRING_TELEFONE))
    iPOS = iPOS + STRING_TELEFONE
    
    sEmailCobr = Trim(Mid(sRegistro, iPOS, STRING_EMAIL))
    iPOS = iPOS + STRING_EMAIL
    
    sFaxCobr = Trim(Mid(sRegistro, iPOS, STRING_FAX))
    iPOS = iPOS + STRING_FAX
    
    sContatoCobr = Trim(Mid(sRegistro, iPOS, STRING_CONTATO))
    iPOS = iPOS + STRING_CONTATO
    
    sObservacaoFilial = Trim(Mid(sRegistro, iPOS, STRING_CLIENTE_OBSERVACAO))
    iPOS = iPOS + STRING_CLIENTE_OBSERVACAO
    
    iVendedor = StrParaLong(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    lCodExterno = StrParaLong(Mid(sRegistro, iPOS, 6))
    iPOS = iPOS + 6

    sTipoCliente = Trim(Mid(sRegistro, iPOS, 50))
    iPOS = iPOS + 50

    Desmonta_Registro_Cliente = SUCESSO

    Exit Function

Erro_Desmonta_Registro_Cliente:

    Desmonta_Registro_Cliente = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189710)

    End Select

    Exit Function
    
End Function

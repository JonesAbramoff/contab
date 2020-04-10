VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTribNFAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'utilizada em telas de entrada de NFs

Implements ClassTributoDoc

Private mvarobjTributacaoDoc As ClassTributacaoDocAux

Private mvarcolItens As New Collection
Private mvarcolItensSRV As New Collection
Private mvarcolTiposTribICMS As New Collection
Private mvarcolTiposTribIPI As New Collection
Private mvarcolTiposTribPISCOFINS As New Collection
Private mvarcolTiposTribICMSSimples As New Collection
Private mvarcolTiposTribISS As New Collection

Private mvarobjNFiscal As ClassNFiscal

Private mvariFilialEmpresa As Integer
'Private mvarlCodigo As Long
Private mvarlCliente As Long
Private mvariFilialCli As Integer
Private mvariFilialEntrega As Integer 'do cliente
Private mvarlFornecedor As Long
Private mvariFilialForn As Integer
Private mvarsNaturezaOp As String
Private mvariNaturezaOpManual As Integer
Private mvardValorFrete As Double
Private mvardValorDesconto As Double
Private mvardValorSeguro As Double
Private mvardValorOutrasDespesas As Double
Private mvariTipoNFiscal As Integer
Private mvarsSiglaTipoNFiscal As String
Private mvardtDataEmissao As Date
Private mvardtDataEntrada As Date
Private mvarsSerie As String

Private mvariTipoTributacao As Integer
Private mvariTipoTributacaoManual As Integer
Private mvarsNatOpExterna As String
Private mvariNatOpExternaManual As Integer
Private mvardICMSBase As Double
Private mvariICMSBaseManual As Integer
Private mvardICMSValor As Double
Private mvariICMSValorManual As Integer
Private mvardICMSSubstBase As Double
Private mvariICMSSubstBaseManual As Integer
Private mvardICMSSubstValor As Double
Private mvariICMSSubstValorManual As Integer
Private mvardICMSCredito As Double
Private mvariICMSCreditoManual As Integer
Private mvardIPIBase As Double
Private mvariIPIBaseManual As Integer
Private mvardIPIValor As Double
Private mvariIPIValorManual As Integer
Private mvardIPICredito As Double
Private mvariIPICreditoManual As Integer
Private mvardIRRFBase As Double
Private mvariIRRFBaseManual As Integer
Private mvardIRRFAliquota As Double
Private mvariIRRFAliquotaManual As Integer
Private mvardIRRFValor As Double
Private mvariIRRFValorManual As Integer
Private mvariISSIncluso As Integer
Private mvardISSBase As Double
Private mvariISSBaseManual As Integer
Private mvardISSAliquota As Double
Private mvariISSAliquotaManual As Integer
Private mvardISSValor As Double
Private mvariISSValorManual As Integer
Private mvariISSInclusoManual As Integer
Private mvardPISCredito As Double
Private mvariPISCreditoManual As Integer
Private mvardCOFINSCredito As Double
Private mvariCOFINSCreditoManual As Integer

Private mvardINSSBase As Double
Private mvariINSSBaseManual As Integer
Private mvardINSSDeducoes As Double
Private mvariINSSDeducoesManual As Integer
Private mvardINSSValor As Double
Private mvariINSSValorManual As Integer
Private mvariINSSRetido As Integer
Private mvariINSSRetidoManual As Integer

Private mvardPISRetido As Double
Private mvariPISRetidoManual As Integer
Private mvardISSRetido As Double
Private mvariISSRetidoManual As Integer
Private mvardCOFINSRetido As Double
Private mvariCOFINSRetidoManual As Integer
Private mvardCSLLRetido As Double
Private mvariCSLLRetidoManual As Integer

Private mvarobjTribFrete As New ClassTribItemNFAux
Private mvarobjTribSeguro As New ClassTribItemNFAux
Private mvarobjTribDesconto As New ClassTribItemNFAux
Private mvarobjTribOutras As New ClassTribItemNFAux
Private mvarobjTributacao As New ClassTributacaoDoc

Private mvardPISValor As Double
Private mvardCOFINSValor As Double

Private mvardTotTrib As Double
Private mvariTotTribManual As Integer
Private mvariTotTribTipo As Integer
Private mvariTotTribTipoManual As Integer

'nfe 3.10
Public dICMSValorIsento As Double
Public dISSValorDeducao As Double
Public dISSValorOutrasRet As Double
Public dISSValorDescIncond As Double
Public dISSValorDescCond As Double
Public iDestino As Integer
Public iDestinoManual As Integer
Public iFinalidadeNFe As Integer
Public iFinalidadeNFeManual As Integer
Public iIndConsumidorFinal As Integer
Public iIndConsumidorFinalManual As Integer
Public iIndPresenca As Integer
Public iIndPresencaManual As Integer
Public iRegTribEspecial As Integer
Public iRegTribEspecialManual As Integer
Public dtDataPrestServico As Date
Public iDataPrestServicoManual As Integer
Public dIIValor As Double
Public sISSMunicFaturamento As String

'IBPTax 0.0.9
Public dTotTribFed As Double
Public dTotTribEst As Double
Public dTotTribMunic As Double
Public iTotTribFedManual As Integer
Public iTotTribEstManual As Integer
Public iTotTribMunicManual As Integer

'NT 2015.003 (ICMS Interestadual)
Public dICMSInterestVlrFCPUFDest As Double
Public dICMSInterestVlrUFDest As Double
Public dICMSInterestVlrUFRemet As Double

'nfe 4.00
Public dICMSVlrFCP As Double
Public dICMSVlrFCPST As Double
Public dICMSVlrFCPSTRet As Double
Public dIPIVlrDevolvido As Double

Public Property Get dPISSTBase() As Double
Dim objTribDocItem As ClassTribItemNFAux
Dim dAux As Double
    
    dAux = 0

    For Each objTribDocItem In colItens
        dAux = Arredonda_Moeda(dAux + objTribDocItem.dPISSTBase)
    Next
    
    dPISSTBase = dAux

End Property

Public Property Get dPISSTValor() As Double
Dim objTribDocItem As ClassTribItemNFAux
Dim dAux As Double
    
    dAux = 0

    For Each objTribDocItem In colItens
        dAux = Arredonda_Moeda(dAux + objTribDocItem.dPISSTValor)
    Next
    
    dPISSTValor = dAux

End Property

Public Property Get dCOFINSSTBase() As Double
Dim objTribDocItem As ClassTribItemNFAux
Dim dAux As Double
    
    dAux = 0

    For Each objTribDocItem In colItens
        dAux = Arredonda_Moeda(dAux + objTribDocItem.dCOFINSSTBase)
    Next
    
    dCOFINSSTBase = dAux

End Property

Public Property Get dCOFINSSTValor() As Double
Dim objTribDocItem As ClassTribItemNFAux
Dim dAux As Double
    
    dAux = 0

    For Each objTribDocItem In colItens
        dAux = Arredonda_Moeda(dAux + objTribDocItem.dCOFINSSTValor)
    Next
    
    dCOFINSSTValor = dAux

End Property

Public Property Set colTiposTribICMSSimples(ByVal vData As Object)
    Set mvarcolTiposTribICMSSimples = vData
End Property

Public Property Get colTiposTribICMSSimples() As Collection
    Set colTiposTribICMSSimples = mvarcolTiposTribICMSSimples
End Property

Public Property Set colTiposTribISS(ByVal vData As Object)
    Set mvarcolTiposTribISS = vData
End Property

Public Property Get colTiposTribISS() As Collection
    Set colTiposTribISS = mvarcolTiposTribISS
End Property

Public Property Let dPISValor(ByVal vData As Double)
    mvardPISValor = vData
End Property

Public Property Get dPISValor() As Double
    dPISValor = mvardPISValor
End Property

Public Property Let dCOFINSValor(ByVal vData As Double)
    mvardCOFINSValor = vData
End Property

Public Property Get dCOFINSValor() As Double
    dCOFINSValor = mvardCOFINSValor
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let dCOFINSCredito(ByVal vData As Double)
    mvardCOFINSCredito = vData
End Property

Public Property Get dCOFINSCredito() As Double
    dCOFINSCredito = mvardCOFINSCredito
End Property

Public Property Let dPISCredito(ByVal vData As Double)
    mvardPISCredito = vData
End Property

Public Property Get dPISCredito() As Double
    dPISCredito = mvardPISCredito
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtDataEntrada(ByVal vData As Date)
    mvardtDataEntrada = vData
End Property

Public Property Get dtDataEntrada() As Date
    dtDataEntrada = mvardtDataEntrada
End Property

Public Property Let sSiglaTipoNFiscal(ByVal vData As String)
    mvarsSiglaTipoNFiscal = vData
End Property

Public Property Get sSiglaTipoNFiscal() As String
    sSiglaTipoNFiscal = mvarsSiglaTipoNFiscal
End Property

Public Property Let iFilialCli(ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Integer
    iFilialCli = mvariFilialCli
End Property

Public Property Let dICMSCredito(ByVal vData As Double)
    mvardICMSCredito = vData
End Property

Public Property Get dICMSCredito() As Double
    dICMSCredito = mvardICMSCredito
End Property

Public Property Let dIPICredito(ByVal vData As Double)
    mvardIPICredito = vData
End Property

Public Property Get dIPICredito() As Double
    dIPICredito = mvardIPICredito
End Property

Public Property Let sNatOpExterna(ByVal vData As String)
    mvarsNatOpExterna = vData
End Property

Public Property Get sNatOpExterna() As String
    sNatOpExterna = mvarsNatOpExterna
End Property

Public Property Let iTipoNFiscal(ByVal vData As Integer)
    mvariTipoNFiscal = vData
End Property

Public Property Get iTipoNFiscal() As Integer
    iTipoNFiscal = mvariTipoNFiscal
End Property

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let dValorSeguro(ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro() As Double
    dValorSeguro = mvardValorSeguro
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorFrete(ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete() As Double
    dValorFrete = mvardValorFrete
End Property

Public Property Let dValorOutrasDespesas(ByVal vData As Double)
    mvardValorOutrasDespesas = vData
End Property

Public Property Get dValorOutrasDespesas() As Double
    dValorOutrasDespesas = mvardValorOutrasDespesas
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let iFilialEntrega(ByVal vData As Integer)
    mvariFilialEntrega = vData
End Property

Public Property Get iFilialEntrega() As Integer
    iFilialEntrega = mvariFilialEntrega
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

''Public Property Let lCodigo(ByVal vData As Long)
''    mvarlCodigo = vData
''End Property
''
''Public Property Get lCodigo() As Long
''    lCodigo = mvarlCodigo
''End Property
''
Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Get colItensSRV() As Collection
    Set colItensSRV = mvarcolItensSRV
End Property

Private Function ClassTributoDoc_Ativar() As Long
Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItem As ClassTribItemNFAux
On Error GoTo Erro_ClassTributoDoc_Ativar

    lErro = gobjTributacao.Tributacao_Doc_Iniciar(Me)
    If lErro <> SUCESSO Then Error 27812
    
    Set mvarobjTributacaoDoc = New ClassTributacaoDocAux
    
    'ativar a NF em si
    If mvarobjTributacaoDoc.Ativar(Me) <> SUCESSO Then
        Set mvarobjTributacaoDoc = Nothing
        Error 27813
    End If
    
    'ativar os seus itens
    For Each objItem In colItens
        Set objTributoDocItem = objItem
        lErro = objTributoDocItem.Ativar(Me)
        If lErro <> SUCESSO Then Error 27814
    Next
    
    'ativar os seus itens
    For Each objItem In colItensSRV
        Set objTributoDocItem = objItem
        lErro = objTributoDocItem.Ativar(Me)
        If lErro <> SUCESSO Then Error 27814
    Next
        
    ClassTributoDoc_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_Ativar:

    ClassTributoDoc_Ativar = Err

    Select Case Err

        Case 27812, 27813, 27814
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153708)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_Alterado() As Long
    ClassTributoDoc_Alterado = mvarobjTributacaoDoc.Alterado
End Function

Private Function ClassTributoDoc_Desativar() As Long
Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItem As ClassTribItemNFAux
On Error GoTo Erro_ClassTributoDoc_Desativar
    
    'desativar os seus itens
    For Each objItem In colItens
        Set objTributoDocItem = objItem
        lErro = objTributoDocItem.Desativar()
        If lErro <> SUCESSO Then Error 27815
    Next
    
    'desativar os seus itens
    For Each objItem In colItensSRV
        Set objTributoDocItem = objItem
        lErro = objTributoDocItem.Desativar()
        If lErro <> SUCESSO Then Error 27815
    Next
    
    'desativar a NF em si
    If Not (mvarobjTributacaoDoc Is Nothing) Then
        lErro = mvarobjTributacaoDoc.Desativar()
        If lErro <> SUCESSO Then Error 27816
        Set mvarobjTributacaoDoc = Nothing
    End If
     
    lErro = gobjTributacao.Tributacao_Doc_Terminar(Me)
    If lErro <> SUCESSO Then Error 27817
    
    Set mvarobjNFiscal = Nothing
    
    ClassTributoDoc_Desativar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_Desativar:

    ClassTributoDoc_Desativar = Err

    Select Case Err

        Case 27815, 27816, 27817
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153709)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_GuardarVarTela(sIdentificador As String, vValor As Variant) As Long
Dim lErro As Long, objVariavelCalculo As New ClassVariavelCalculo
On Error GoTo Erro_ClassTributoDoc_GuardarVarTela

    objVariavelCalculo.sIdentificador = sIdentificador
    objVariavelCalculo.vValor = vValor
    objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
    
    ClassTributoDoc_GuardarVarTela = ClassTributoDoc_GuardarVar(objVariavelCalculo)

    Exit Function

Erro_ClassTributoDoc_GuardarVarTela:

    ClassTributoDoc_GuardarVarTela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153710)

    End Select

    Exit Function

End Function

Private Property Get ClassTributoDoc_objTributacaoDoc() As GlobaisTRB.ClassTributacaoDocAux
    Set ClassTributoDoc_objTributacaoDoc = mvarobjTributacaoDoc
End Property

Private Function ClassTributoDoc_ObterColDocCompl(colDocCompl As Collection) As Long
'Dim objTribCompl As ClassTribComplNFAux
'
'    Set colDocCompl = New Collection
'
'    Set objTribCompl = New ClassTribComplNFAux
'    Call objTribCompl.Iniciar(Me, mvarobjTribFrete)
'    Call colDocCompl.Add(objTribCompl)
'
'    Set objTribCompl = New ClassTribComplNFAux
'    Call objTribCompl.Iniciar(Me, mvarobjTribSeguro)
'    Call colDocCompl.Add(objTribCompl)
'
'    Set objTribCompl = New ClassTribComplNFAux
'    Call objTribCompl.Iniciar(Me, mvarobjTribDesconto)
'    Call colDocCompl.Add(objTribCompl)
'
'    Set objTribCompl = New ClassTribComplNFAux
'    Call objTribCompl.Iniciar(Me, mvarobjTribOutras)
'    Call colDocCompl.Add(objTribCompl)
    
    ClassTributoDoc_ObterColDocCompl = SUCESSO
    
End Function

Private Function ClassTributoDoc_ObterColDocItens(colDocItens As Collection) As Long
'retorna colecao de itens que implementam interface como TaxDocItemExemplo
Dim objItem As ClassTribItemNFAux

   Set colDocItens = New Collection
   
   For Each objItem In mvarcolItens
        Call colDocItens.Add(objItem)
   Next
   
   For Each objItem In mvarcolItensSRV
        Call colDocItens.Add(objItem)
   Next
   
   ClassTributoDoc_ObterColDocItens = SUCESSO
   
End Function

Private Function ClassTributoDoc_ObterDocAux(objTributacaoDoc As ClassTributacaoDocAux) As Long
'retorna objeto auxiliar vinculado ao Doc
    Set objTributacaoDoc = mvarobjTributacaoDoc
    ClassTributoDoc_ObterDocAux = SUCESSO
End Function

Private Function ClassTributoDoc_ObterNatOpPadrao(sNatOp As String) As Long
Dim lErro As Long
On Error GoTo Erro_ClassTributoDoc_ObterNatOpPadrao

    lErro = mvarobjTributacaoDoc.ObterNaturezaOpPadrao(sNatOp)
    If lErro <> SUCESSO Then Error 27818
    
    ClassTributoDoc_ObterNatOpPadrao = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterNatOpPadrao:

    ClassTributoDoc_ObterNatOpPadrao = Err

    Select Case Err

        Case 27818
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153711)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_ObterTribPadrao(iTipoTrib As Integer) As Long
Dim lErro As Long
On Error GoTo Erro_ClassTributoDoc_ObterTribPadrao

    lErro = mvarobjTributacaoDoc.ObterTribPadrao(iTipoTrib)
    If lErro <> SUCESSO Then Error 27819
    
    ClassTributoDoc_ObterTribPadrao = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterTribPadrao:

    ClassTributoDoc_ObterTribPadrao = Err

    Select Case Err

        Case 27819
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153712)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'retorna o valor solicitado se este estiver "disponivel" em seu "contexto"
'objVariavelCalculo.sIdentificador tem que vir preenchida
Dim lErro As Long
Dim vValorAux As Variant

On Error GoTo Erro_ClassTributoDoc_ObterVar
    
    If objVariavelCalculo Is Nothing Then
        Set objVariavelCalculo = New ClassVariavelCalculo
    End If
    
    objVariavelCalculo.sIdentificador = sIdentificador
    
    Select Case sIdentificador
        
        Case "NFISCAL_OBJ"
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            Set objVariavelCalculo.vValor = mvarobjNFiscal
            
        Case "SERIE"
            objVariavelCalculo.vValor = mvarsSerie
            If mvarsSerie <> "" Then
                objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            Else
                objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
            End If
        
        Case "DATA_ENTRADA"
            objVariavelCalculo.vValor = mvardtDataEntrada
            If mvardtDataEntrada <> DATA_NULA Then
                objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            Else
                objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
            End If
        
        Case "DATA_EMISSAO"
            objVariavelCalculo.vValor = mvardtDataEmissao
            If mvardtDataEmissao <> DATA_NULA Then
                objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            Else
                objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
            End If
        
'        Case "NAT_OPERACAO"
'            objVariavelCalculo.vValor = mvarsNaturezaOp
'            objVariavelCalculo.iStatus = mvariNaturezaOpManual
        
        Case "NAT_OPERACAO_EXTERNA"
            objVariavelCalculo.vValor = mvarsNatOpExterna
            objVariavelCalculo.iStatus = mvariNatOpExternaManual
        
'        Case "TIPO_TRIB"
'            objVariavelCalculo.vValor = mvariTipoTributacao
'            objVariavelCalculo.iStatus = mvariTipoTributacaoManual
        
        Case "TIPO_DOC_SIGLA"
            objVariavelCalculo.vValor = mvarsSiglaTipoNFiscal
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "TIPO_DOC_CODIGO"
            objVariavelCalculo.vValor = mvariTipoNFiscal
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "CLIENTE_CODIGO"
            objVariavelCalculo.vValor = mvarlCliente
            If mvarlCliente <> 0 Then
                objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            Else
                objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
            End If
            
        Case "FILIAL_CLIENTE_CODIGO"
            objVariavelCalculo.vValor = mvariFilialCli
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "FORN_CODIGO"
            objVariavelCalculo.vValor = mvarlFornecedor
            If mvarlFornecedor <> 0 Then
                objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            Else
                objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
            End If
            
        Case "FILIAL_FORN_CODIGO"
            objVariavelCalculo.vValor = mvariFilialForn
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "DESCONTO_VALOR"
            objVariavelCalculo.vValor = mvardValorDesconto
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "FRETE_VALOR"
            objVariavelCalculo.vValor = mvardValorFrete
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "SEGURO_VALOR"
            objVariavelCalculo.vValor = mvardValorSeguro
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "DESP_ACESS_VALOR"
            objVariavelCalculo.vValor = mvardValorOutrasDespesas
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

'        'para ICMS
'        Case "ICMS_BASE"
'            objVariavelCalculo.vValor = mvardICMSBase
'            objVariavelCalculo.iStatus = mvariICMSBaseManual
'
'        Case "ICMS_VALOR"
'            objVariavelCalculo.vValor = mvardICMSValor
'            objVariavelCalculo.iStatus = mvariICMSValorManual
'
'        Case "ICMS_SUBST_BASE"
'            objVariavelCalculo.vValor = mvardICMSSubstBase
'            objVariavelCalculo.iStatus = mvariICMSSubstBaseManual
'
'        Case "ICMS_SUBST_VALOR"
'            objVariavelCalculo.vValor = mvardICMSSubstValor
'            objVariavelCalculo.iStatus = mvariICMSSubstValorManual
'
'        Case "ICMS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvardICMSCredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        'para IPI
'        Case "IPI_BASE"
'            objVariavelCalculo.vValor = mvardIPIBase
'            objVariavelCalculo.iStatus = mvariIPIBaseManual
'
'        Case "IPI_VALOR"
'            objVariavelCalculo.vValor = mvardIPIValor
'            objVariavelCalculo.iStatus = mvariIPIValorManual
'
'        Case "IPI_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvardIPICredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        'para ISS
'        Case "ISS_INCLUSO"
'            objVariavelCalculo.vValor = mvariISSIncluso
'            objVariavelCalculo.iStatus = mvariISSInclusoManual
'
'        Case "ISS_BASE"
'            objVariavelCalculo.vValor = mvardISSBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "ISS_ALIQUOTA"
'            objVariavelCalculo.vValor = mvardISSAliquota
'            objVariavelCalculo.iStatus = mvariISSAliquotaManual
'
'        Case "ISS_VALOR"
'            objVariavelCalculo.vValor = mvardISSValor
'            objVariavelCalculo.iStatus = mvariISSValorManual
'
'        'para IR
'        Case "IR_BASE"
'            objVariavelCalculo.vValor = mvardIRRFBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "IR_ALIQUOTA"
'            objVariavelCalculo.vValor = mvardIRRFAliquota
'            objVariavelCalculo.iStatus = mvariIRRFAliquotaManual
'
'        Case "IR_VALOR"
'            objVariavelCalculo.vValor = mvardIRRFValor
'            objVariavelCalculo.iStatus = mvariIRRFValorManual
'
'        'para INSS
'        Case "INSS_BASE"
'            objVariavelCalculo.vValor = mvardINSSBase
'            objVariavelCalculo.iStatus = mvariINSSBaseManual
'
'        Case "INSS_VALOR_DEDUCOES"
'            objVariavelCalculo.vValor = mvardINSSDeducoes
'            objVariavelCalculo.iStatus = mvariINSSDeducoesManual
'
'        Case "INSS_VALOR"
'            objVariavelCalculo.vValor = mvardINSSValor
'            objVariavelCalculo.iStatus = mvariINSSValorManual
'
'        Case "INSS_RETIDO"
'            objVariavelCalculo.vValor = mvariINSSRetido
'            objVariavelCalculo.iStatus = mvariINSSRetidoManual
'
'        Case "PIS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvardPISCredito
'            objVariavelCalculo.iStatus = mvariPISCreditoManual
'
'        Case "COFINS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvardCOFINSCredito
'            objVariavelCalculo.iStatus = mvariCOFINSCreditoManual
'
'        Case "PIS_RETIDO"
'            objVariavelCalculo.vValor = mvardPISRetido
'            objVariavelCalculo.iStatus = mvariPISRetidoManual
'
'        Case "ISS_RETIDO"
'            objVariavelCalculo.vValor = mvardISSRetido
'            objVariavelCalculo.iStatus = mvariISSRetidoManual
'
'        Case "COFINS_RETIDO"
'            objVariavelCalculo.vValor = mvardCOFINSRetido
'            objVariavelCalculo.iStatus = mvariCOFINSRetidoManual
'
'        Case "CSLL_RETIDO"
'            objVariavelCalculo.vValor = mvardCSLLRetido
'            objVariavelCalculo.iStatus = mvariCSLLRetidoManual
'
'        Case "BASE_ICMS_IMPORTACAO_SIMPLES"
'            objVariavelCalculo.vValor = Arredonda_Moeda(mvardIPIBase + mvardIPIValor + mvardValorOutrasDespesas, 2)
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "PIS_TIPO_OBJ"
            Set vValorAux = objVariavelCalculo.vValor
            Call PISCOFINSTipo_Obter(vValorAux)
            Set objVariavelCalculo.vValor = vValorAux
            
        Case "COFINS_TIPO_OBJ"
            Set vValorAux = objVariavelCalculo.vValor
            Call PISCOFINSTipo_Obter(vValorAux)
            Set objVariavelCalculo.vValor = vValorAux
            
        Case Else
'        Case "FILIAL_CLIENTE_OBJ"
'        Case "FILIAL_CLIENTE_INSC_EST"
'        Case "CLIENTE_ENDERECO_OBJ"
'        Case "FILIAL_FORN_OBJ"
'        Case "FILIAL_FORN_INSC_EST"
'        Case "FORN_ENDERECO_OBJ"
'        Case "FORN_UF" '(EX p/exterior)
'        Case "CLIENTE_UF" '(EX p/exterior)
            
            'delegar p/tributacaoDoc
            lErro = objTributacao.ObterVar_Padrao(sIdentificador, objVariavelCalculo, Me)
            If lErro <> SUCESSO Then
                lErro = mvarobjTributacaoDoc.ObterVar(sIdentificador, objVariavelCalculo)
                If lErro <> SUCESSO Then Error 27820
            End If
            
    End Select
    
    ClassTributoDoc_ObterVar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterVar:

    ClassTributoDoc_ObterVar = Err

    Select Case Err

        Case 27820
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153713)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_GuardarVar(objVariavelCalculo As ClassVariavelCalculo) As Long
'armazena info de variavel guardando se esta foi alterada ou nao desde ...
Dim lErro As Long
On Error GoTo Erro_ClassTributoDoc_GuardarVar
        
    Select Case objVariavelCalculo.sIdentificador
        
'        Case "NAT_OPERACAO"
'            mvarsNaturezaOp = objVariavelCalculo.vValor
'            mvariNaturezaOpManual = objVariavelCalculo.iStatus
            
        Case "NAT_OPERACAO_EXTERNA"
            mvarsNatOpExterna = objVariavelCalculo.vValor
            mvariNatOpExternaManual = objVariavelCalculo.iStatus
            
        Case "TIPO_TRIB"
            mvariTipoTributacao = objVariavelCalculo.vValor
            mvariTipoTributacaoManual = objVariavelCalculo.iStatus
        
        Case "TIPO_DOC_SIGLA"
        Case "TIPO_DOC_CODIGO"
        
        Case "CLIENTE_CODIGO"
        Case "FORN_CODIGO"
        
        Case "FILIAL_CLIENTE_OBJ"
        Case "FILIAL_FORN_OBJ"
        
        Case "FILIAL_CLIENTE_CODIGO"
        Case "FILIAL_FORN_CODIGO"
        
        Case "FILIAL_CLIENTE_INSC_EST"
        Case "FILIAL_FORN_INSC_EST"

        Case "CLIENTE_ENDERECO_OBJ"
        Case "FORN_ENDERECO_OBJ"
        
        Case "FORN_UF" '(EX p/exterior)
        Case "CLIENTE_UF" '(EX p/exterior)
        
'        Case "FRETE_VALOR"
'
'        Case "SEGURO_VALOR"
'
'        Case "DESP_ACESS_VALOR"
'
'        'para ICMS
'        Case "ICMS_BASE"
'            mvardICMSBase = objVariavelCalculo.vValor
'            mvariICMSBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR"
'            mvardICMSValor = objVariavelCalculo.vValor
'            mvariICMSValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_SUBST_BASE"
'            mvardICMSSubstBase = objVariavelCalculo.vValor
'            mvariICMSSubstBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_SUBST_VALOR"
'            mvardICMSSubstValor = objVariavelCalculo.vValor
'            mvariICMSSubstValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR_CREDITO"
'            mvardICMSCredito = objVariavelCalculo.vValor
'
'        'para IPI
'        Case "IPI_BASE"
'            mvardIPIBase = objVariavelCalculo.vValor
'            mvariIPIBaseManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR"
'            mvardIPIValor = objVariavelCalculo.vValor
'            mvariIPIValorManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR_CREDITO"
'            mvardIPICredito = objVariavelCalculo.vValor
'
'        'para ISS
'        Case "ISS_INCLUSO"
'            mvariISSIncluso = objVariavelCalculo.vValor
'            mvariISSInclusoManual = objVariavelCalculo.iStatus
'
'        Case "ISS_BASE"
'            mvardISSBase = objVariavelCalculo.vValor
'
'        Case "ISS_ALIQUOTA"
'            mvardISSAliquota = objVariavelCalculo.vValor
'            mvariISSAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "ISS_VALOR"
'            mvardISSValor = objVariavelCalculo.vValor
'            mvariISSValorManual = objVariavelCalculo.iStatus
'
'        'para IR
'        Case "IR_BASE"
'            mvardIRRFBase = objVariavelCalculo.vValor
'
'        Case "IR_ALIQUOTA"
'            mvardIRRFAliquota = objVariavelCalculo.vValor
'            mvariIRRFAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "IR_VALOR"
'            mvardIRRFValor = objVariavelCalculo.vValor
'            mvariIRRFValorManual = objVariavelCalculo.iStatus
'
'        'para INSS
'        Case "INSS_BASE"
'            mvardINSSBase = objVariavelCalculo.vValor
'            mvariINSSBaseManual = objVariavelCalculo.iStatus
'
'        Case "INSS_VALOR_DEDUCOES"
'            mvardINSSDeducoes = objVariavelCalculo.vValor
'            mvariINSSDeducoesManual = objVariavelCalculo.iStatus
'
'        Case "INSS_VALOR"
'            mvardINSSValor = objVariavelCalculo.vValor
'            mvariINSSValorManual = objVariavelCalculo.iStatus
'
'        Case "INSS_RETIDO"
'            mvariINSSRetido = objVariavelCalculo.vValor
'            mvariINSSRetidoManual = objVariavelCalculo.iStatus
'
'        Case "PIS_VALOR_CREDITO"
'            mvardPISCredito = objVariavelCalculo.vValor
'            mvariPISCreditoManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_VALOR_CREDITO"
'            mvardCOFINSCredito = objVariavelCalculo.vValor
'            mvariCOFINSCreditoManual = objVariavelCalculo.iStatus
'
'        Case "PIS_RETIDO"
'            mvardPISRetido = objVariavelCalculo.vValor
'            mvariPISRetidoManual = objVariavelCalculo.iStatus
'
'        Case "ISS_RETIDO"
'            mvardISSRetido = objVariavelCalculo.vValor
'            mvariISSRetidoManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_RETIDO"
'            mvardCOFINSRetido = objVariavelCalculo.vValor
'            mvariCOFINSRetidoManual = objVariavelCalculo.iStatus
'
'        Case "CSLL_RETIDO"
'            mvardCSLLRetido = objVariavelCalculo.vValor
'            mvariCSLLRetidoManual = objVariavelCalculo.iStatus
        
        Case Else
        
            lErro = objTributacao.GuardarVar_Padrao(objVariavelCalculo, Me)
            If lErro <> SUCESSO Then
                'delegar p/tributacaoDoc
                lErro = mvarobjTributacaoDoc.GuardarVar(objVariavelCalculo)
                If lErro <> SUCESSO Then Error 27821
            End If
            
    End Select

    ClassTributoDoc_GuardarVar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_GuardarVar:

    ClassTributoDoc_GuardarVar = Err

    Select Case Err

        Case 27821
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153714)

    End Select

    Exit Function

End Function

Public Function IncluirItem(objItem As ClassTribItemNFAux) As Long
'associa um novo item a NF
Dim lErro As Long
On Error GoTo Erro_IncluirItem

    Call mvarcolItens.Add(objItem)

    IncluirItem = SUCESSO

    Exit Function

Erro_IncluirItem:

    IncluirItem = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153715)

    End Select

    Exit Function

End Function

Public Function IncluirItemSRV(objItem As ClassTribItemNFAux) As Long
'associa um novo item a NF
Dim lErro As Long
On Error GoTo Erro_IncluirItemSRV

    Call mvarcolItensSRV.Add(objItem)

    IncluirItemSRV = SUCESSO

    Exit Function

Erro_IncluirItemSRV:

    IncluirItemSRV = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153715)

    End Select

    Exit Function

End Function

Public Function RemoverItem(vntIndexKey As Variant) As Long
'exclui um item da NF identificando-o como na colecao de itens
Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItem As ClassTribItemNFAux
On Error GoTo Erro_RemoverItem

    Set objItem = colItens.Item(vntIndexKey)
    
    Set objTributoDocItem = objItem
    lErro = objTributoDocItem.Desativar()
    If lErro <> SUCESSO Then Error 27822
        
    Call colItens.Remove(vntIndexKey)
    
    RemoverItem = SUCESSO

    Exit Function

Erro_RemoverItem:

    RemoverItem = Err

    Select Case Err

        Case 27822
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153716)

    End Select

    Exit Function

End Function

Public Function RemoverItemSRV(vntIndexKey As Variant) As Long
'exclui um item da NF identificando-o como na colecao de itens
Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItem As ClassTribItemNFAux
On Error GoTo Erro_RemoverItemSRV

    Set objItem = colItensSRV.Item(vntIndexKey)
    
    Set objTributoDocItem = objItem
    lErro = objTributoDocItem.Desativar()
    If lErro <> SUCESSO Then Error 27822
        
    Call colItensSRV.Remove(vntIndexKey)
    
    RemoverItemSRV = SUCESSO

    Exit Function

Erro_RemoverItemSRV:

    RemoverItemSRV = Err

    Select Case Err

        Case 27822
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153716)

    End Select

    Exit Function

End Function

Public Property Get objTribFrete() As ClassTribItemNFAux
    Set objTribFrete = mvarobjTribFrete
End Property

Public Property Set objTribFrete(ByVal vData As Object)
    Set mvarobjTribFrete = vData
End Property

Public Property Get objTribSeguro() As ClassTribItemNFAux
    Set objTribSeguro = mvarobjTribSeguro
End Property

Public Property Set objTribSeguro(ByVal vData As Object)
    Set mvarobjTribSeguro = vData
End Property

Public Property Get objTribDesconto() As ClassTribItemNFAux
    Set objTribDesconto = mvarobjTribDesconto
End Property

Public Property Set objTribDesconto(ByVal vData As Object)
    Set mvarobjTribDesconto = vData
End Property

Public Property Get objTribOutras() As ClassTribItemNFAux
    Set objTribOutras = mvarobjTribOutras
End Property

Public Property Set objTribOutras(ByVal vData As Object)
    Set mvarobjTribOutras = vData
End Property

'Public Property Let iFilialEmpresa(ByVal vData As Integer)
'    mvariFilialEmpresa = vData
'End Property
'
'Public Property Get iFilialEmpresa() As Integer
'    iFilialEmpresa = mvariFilialEmpresa
'End Property
'
Public Property Let iIRRFValorManual(ByVal vData As Integer)
    mvariIRRFValorManual = vData
End Property

Public Property Get iIRRFValorManual() As Integer
    iIRRFValorManual = mvariIRRFValorManual
End Property

Public Property Let dIRRFValor(ByVal vData As Double)
    mvardIRRFValor = vData
End Property

Public Property Get dIRRFValor() As Double
    dIRRFValor = mvardIRRFValor
End Property

Public Property Let iISSIncluso(ByVal vData As Integer)
    mvariISSIncluso = vData
End Property

Public Property Get iISSIncluso() As Integer
    iISSIncluso = mvariISSIncluso
End Property

Public Property Let iISSValorManual(ByVal vData As Integer)
    mvariISSValorManual = vData
End Property

Public Property Get iISSValorManual() As Integer
    iISSValorManual = mvariISSValorManual
End Property

Public Property Let iPISCreditoManual(ByVal vData As Integer)
    mvariPISCreditoManual = vData
End Property

Public Property Get iPISCreditoManual() As Integer
    iPISCreditoManual = mvariPISCreditoManual
End Property

Public Property Let iCOFINSCreditoManual(ByVal vData As Integer)
    mvariCOFINSCreditoManual = vData
End Property

Public Property Get iCOFINSCreditoManual() As Integer
    iCOFINSCreditoManual = mvariCOFINSCreditoManual
End Property

Public Property Let iISSInclusoManual(ByVal vData As Integer)
    mvariISSInclusoManual = vData
End Property

Public Property Get iISSInclusoManual() As Integer
    iISSInclusoManual = mvariISSInclusoManual
End Property

Public Property Let dISSValor(ByVal vData As Double)
    mvardISSValor = vData
End Property

Public Property Get dISSValor() As Double
    dISSValor = mvardISSValor
End Property

Public Property Let iIPIValorManual(ByVal vData As Integer)
    mvariIPIValorManual = vData
End Property

Public Property Get iIPIValorManual() As Integer
    iIPIValorManual = mvariIPIValorManual
End Property

Public Property Let dIPIValor(ByVal vData As Double)
    mvardIPIValor = vData
End Property

Public Property Get dIPIValor() As Double
    dIPIValor = mvardIPIValor
End Property

Public Property Let iICMSSubstValorManual(ByVal vData As Integer)
    mvariICMSSubstValorManual = vData
End Property

Public Property Get iICMSSubstValorManual() As Integer
    iICMSSubstValorManual = mvariICMSSubstValorManual
End Property

Public Property Let dICMSSubstValor(ByVal vData As Double)
    mvardICMSSubstValor = vData
End Property

Public Property Get dICMSSubstValor() As Double
    dICMSSubstValor = mvardICMSSubstValor
End Property

Public Property Let iICMSSubstBaseManual(ByVal vData As Integer)
    mvariICMSSubstBaseManual = vData
End Property

Public Property Get iICMSSubstBaseManual() As Integer
    iICMSSubstBaseManual = mvariICMSSubstBaseManual
End Property

Public Property Let dICMSSubstBase(ByVal vData As Double)
    mvardICMSSubstBase = vData
End Property

Public Property Get dICMSSubstBase() As Double
    dICMSSubstBase = mvardICMSSubstBase
End Property

Public Property Let iICMSValorManual(ByVal vData As Integer)
    mvariICMSValorManual = vData
End Property

Public Property Get iICMSValorManual() As Integer
    iICMSValorManual = mvariICMSValorManual
End Property

Public Property Let dICMSValor(ByVal vData As Double)
    mvardICMSValor = vData
End Property

Public Property Get dICMSValor() As Double
    dICMSValor = mvardICMSValor
End Property

Public Property Let iICMSBaseManual(ByVal vData As Integer)
    mvariICMSBaseManual = vData
End Property

Public Property Get iICMSBaseManual() As Integer
    iICMSBaseManual = mvariICMSBaseManual
End Property

Public Property Let dICMSBase(ByVal vData As Double)
    mvardICMSBase = vData
End Property

Public Property Get dICMSBase() As Double
    dICMSBase = mvardICMSBase
End Property

Public Property Let iTipoTributacaoManual(ByVal vData As Integer)
    mvariTipoTributacaoManual = vData
End Property

Public Property Get iTipoTributacaoManual() As Integer
    iTipoTributacaoManual = mvariTipoTributacaoManual
End Property

Public Property Let iNaturezaOpManual(ByVal vData As Integer)
    mvariNaturezaOpManual = vData
End Property

Public Property Get iNaturezaOpManual() As Integer
    iNaturezaOpManual = mvariNaturezaOpManual
End Property

Public Property Let iNatOpExternaManual(ByVal vData As Integer)
    mvariNatOpExternaManual = vData
End Property

Public Property Get iNatOpExternaManual() As Integer
    iNatOpExternaManual = mvariNatOpExternaManual
End Property

Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

'Public Property Let iTaxacaoAutomatica(ByVal vData As Integer)
'    mvariTaxacaoAutomatica = vData
'End Property
'
'Public Property Get iTaxacaoAutomatica() As Integer
'    iTaxacaoAutomatica = mvariTaxacaoAutomatica
'End Property
'
Public Property Let iIPIBaseManual(ByVal vData As Integer)
    mvariIPIBaseManual = vData
End Property

Public Property Get iIPIBaseManual() As Integer
    iIPIBaseManual = mvariIPIBaseManual
End Property

Public Property Let dIPIBase(ByVal vData As Double)
    mvardIPIBase = vData
End Property

Public Property Get dIPIBase() As Double
    dIPIBase = mvardIPIBase
End Property

Public Property Let iISSAliquotaManual(ByVal vData As Integer)
    mvariISSAliquotaManual = vData
End Property

Public Property Get iISSAliquotaManual() As Integer
    iISSAliquotaManual = mvariISSAliquotaManual
End Property

Public Property Let dISSAliquota(ByVal vData As Double)
    mvardISSAliquota = vData
End Property

Public Property Get dISSAliquota() As Double
    dISSAliquota = mvardISSAliquota
End Property

Public Property Let iIRRFAliquotaManual(ByVal vData As Integer)
    mvariIRRFAliquotaManual = vData
End Property

Public Property Get iIRRFAliquotaManual() As Integer
    iIRRFAliquotaManual = mvariIRRFAliquotaManual
End Property

Public Property Let dIRRFAliquota(ByVal vData As Double)
    mvardIRRFAliquota = vData
End Property

Public Property Get dIRRFAliquota() As Double
    dIRRFAliquota = mvardIRRFAliquota
End Property

Public Property Let dISSBase(ByVal vData As Double)
    mvardISSBase = vData
End Property

Public Property Get dISSBase() As Double
    dISSBase = mvardISSBase
End Property

Public Property Let dIRRFBase(ByVal vData As Double)
    mvardIRRFBase = vData
End Property

Public Property Get dIRRFBase() As Double
    dIRRFBase = mvardIRRFBase
End Property

Public Property Let dINSSBase(ByVal vData As Double)
    mvardINSSBase = vData
End Property

Public Property Get dINSSBase() As Double
    dINSSBase = mvardINSSBase
End Property

Public Property Let dINSSDeducoes(ByVal vData As Double)
    mvardINSSDeducoes = vData
End Property

Public Property Get dINSSDeducoes() As Double
    dINSSDeducoes = mvardINSSDeducoes
End Property

Public Property Let dINSSValor(ByVal vData As Double)
    mvardINSSValor = vData
End Property

Public Property Get dINSSValor() As Double
    dINSSValor = mvardINSSValor
End Property

Public Property Let iINSSRetido(ByVal vData As Integer)
    mvariINSSRetido = vData
End Property

Public Property Get iINSSRetido() As Integer
    iINSSRetido = mvariINSSRetido
End Property

Public Property Let iINSSBaseManual(ByVal vData As Integer)
    mvariINSSBaseManual = vData
End Property

Public Property Get iINSSBaseManual() As Integer
    iINSSBaseManual = mvariINSSBaseManual
End Property

Public Property Let iINSSDeducoesManual(ByVal vData As Integer)
    mvariINSSDeducoesManual = vData
End Property

Public Property Get iINSSDeducoesManual() As Integer
    iINSSDeducoesManual = mvariINSSDeducoesManual
End Property

Public Property Let iINSSValorManual(ByVal vData As Integer)
    mvariINSSValorManual = vData
End Property

Public Property Get iINSSValorManual() As Integer
    iINSSValorManual = mvariINSSValorManual
End Property

Public Property Let iINSSRetidoManual(ByVal vData As Integer)
    mvariINSSRetidoManual = vData
End Property

Public Property Get iINSSRetidoManual() As Integer
    iINSSRetidoManual = mvariINSSRetidoManual
End Property

Public Property Let iTotTribManual(ByVal vData As Integer)
    mvariTotTribManual = vData
End Property

Public Property Get iTotTribManual() As Integer
    iTotTribManual = mvariTotTribManual
End Property

Public Property Let dTotTrib(ByVal vData As Double)
    mvardTotTrib = vData
End Property

Public Property Get dTotTrib() As Double
    dTotTrib = mvardTotTrib
End Property

Public Property Let iTotTribTipo(ByVal vData As Integer)
    mvariTotTribTipo = vData
End Property

Public Property Get iTotTribTipo() As Integer
    iTotTribTipo = mvariTotTribTipo
End Property

Public Property Let iTotTribTipoManual(ByVal vData As Integer)
    mvariTotTribTipoManual = vData
End Property

Public Property Get iTotTribTipoManual() As Integer
    iTotTribTipoManual = mvariTotTribTipoManual
End Property

Private Sub Class_Initialize()
    
    mvarobjTribFrete.iItem = ITEM_TIPO_FRETE
    mvarobjTribSeguro.iItem = ITEM_TIPO_SEGURO
    mvarobjTribDesconto.iItem = ITEM_TIPO_DESCONTO
    mvarobjTribOutras.iItem = ITEM_TIPO_OUTRAS_DESP
    
    dtDataPrestServico = DATA_NULA

End Sub

Public Property Let dPISRetido(ByVal vData As Double)
    mvardPISRetido = vData
End Property

Public Property Get dPISRetido() As Double
    dPISRetido = mvardPISRetido
End Property

Public Property Let dISSRetido(ByVal vData As Double)
    mvardISSRetido = vData
End Property

Public Property Get dISSRetido() As Double
    dISSRetido = mvardISSRetido
End Property

Public Property Let dCOFINSRetido(ByVal vData As Double)
    mvardCOFINSRetido = vData
End Property

Public Property Get dCOFINSRetido() As Double
    dCOFINSRetido = mvardCOFINSRetido
End Property

Public Property Let dCSLLRetido(ByVal vData As Double)
    mvardCSLLRetido = vData
End Property

Public Property Get dCSLLRetido() As Double
    dCSLLRetido = mvardCSLLRetido
End Property

Public Property Let iPISRetidoManual(ByVal vData As Integer)
    mvariPISRetidoManual = vData
End Property

Public Property Get iPISRetidoManual() As Integer
    iPISRetidoManual = mvariPISRetidoManual
End Property

Public Property Let iISSRetidoManual(ByVal vData As Integer)
    mvariISSRetidoManual = vData
End Property

Public Property Get iISSRetidoManual() As Integer
    iISSRetidoManual = mvariISSRetidoManual
End Property

Public Property Let iCOFINSRetidoManual(ByVal vData As Integer)
    mvariCOFINSRetidoManual = vData
End Property

Public Property Get iCOFINSRetidoManual() As Integer
    iCOFINSRetidoManual = mvariCOFINSRetidoManual
End Property

Public Property Let iCSLLRetidoManual(ByVal vData As Integer)
    mvariCSLLRetidoManual = vData
End Property

Public Property Get iCSLLRetidoManual() As Integer
    iCSLLRetidoManual = mvariCSLLRetidoManual
End Property

Public Property Set objNFiscal(ByVal vData As Object)
    Set mvarobjNFiscal = vData
End Property

Public Property Set objTributacao(ByVal vData As Object)
    Set mvarobjTributacao = vData
End Property

Public Property Get objTributacao() As ClassTributacaoDoc
    Set objTributacao = mvarobjTributacao
End Property

Public Property Let iISSBaseManual(ByVal vData As Integer)
    mvariISSBaseManual = vData
End Property

Public Property Get iISSBaseManual() As Integer
    iISSBaseManual = mvariISSBaseManual
End Property

Public Property Let iIRRFBaseManual(ByVal vData As Integer)
    mvariIRRFBaseManual = vData
End Property

Public Property Get iIRRFBaseManual() As Integer
    iIRRFBaseManual = mvariIRRFBaseManual
End Property

Public Property Let iICMSCreditoManual(ByVal vData As Integer)
    mvariICMSCreditoManual = vData
End Property

Public Property Get iICMSCreditoManual() As Integer
    iICMSCreditoManual = mvariICMSCreditoManual
End Property

Public Property Let iIPICreditoManual(ByVal vData As Integer)
    mvariIPICreditoManual = vData
End Property

Public Property Get iIPICreditoManual() As Integer
    iIPICreditoManual = mvariIPICreditoManual
End Property

Public Property Set colTiposTribICMS(ByVal vData As Object)
    Set mvarcolTiposTribICMS = vData
End Property

Public Property Get colTiposTribICMS() As Collection
    Set colTiposTribICMS = mvarcolTiposTribICMS
End Property

Public Property Set colTiposTribIPI(ByVal vData As Object)
    Set mvarcolTiposTribIPI = vData
End Property

Public Property Get colTiposTribIPI() As Collection
    Set colTiposTribIPI = mvarcolTiposTribIPI
End Property

Public Property Set colTiposTribPISCOFINS(ByVal vData As Object)
    Set mvarcolTiposTribPISCOFINS = vData
End Property

Public Property Get colTiposTribPISCOFINS() As Collection
    Set colTiposTribPISCOFINS = mvarcolTiposTribPISCOFINS
End Property

Public Sub PISCOFINSTipo_Obter(vValor As Variant)
    Dim iCodigo As Integer
    iCodigo = vValor.iTipo
    For Each vValor In colTiposTribPISCOFINS
        If iCodigo = vValor.iTipo Then Exit For
    Next
End Sub

Public Sub ICMSSimplesTipo_Obter(vValor As Variant)
    Dim iCodigo As Integer
    iCodigo = vValor.iTipo
    For Each vValor In colTiposTribICMSSimples
        If iCodigo = vValor.iTipo Then Exit For
    Next
End Sub

Public Sub ISSTipo_Obter(vValor As Variant)
    Dim iCodigo As Integer
    iCodigo = vValor.iTipo
    For Each vValor In colTiposTribISS
        If iCodigo = vValor.iTipo Then Exit For
    Next
End Sub

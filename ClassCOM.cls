VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCOM"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialCompra As Integer
Private mvariCompradorAumentaQuant As Integer
Private mvariConsideraQuantCotacaoAnterior As Integer
Private mvardConsumoMedioMax As Double
Private mvariControleAlcada As Integer
Private mvariMesesConsumoMedio As Integer
Private mvariMesesMediaTempoRessup As Integer
Private mvardPercentMaisQuantCotacaoAnterior As Double
Private mvardPercentMaisReceb As Double
Private mvardPercentMenosQuantCotacaoAnterior As Double
Private mvardPercentMenosReceb As Double
Private mvariRecebForaFaixa As Integer
Private mvardResiduo As Double
Private mvardTaxaFinanceiraEmpresa As Double
Private mvariTemFaixaReceb As Integer
Private mvardTempoRessupMax As Double
Private mvariNumComprasMediaAtraso As Integer
Private mvariNumComprasTempoRessup As Integer
Private mvariNFDiferentePC As Integer
Dim mvarsFormatoPrecoUnitario As String
Private mvariEmailAlteraDataEnvioPC As Integer
Private mvariReqCompraDescProdBloq As Integer
Private mvariPedCompraBloqEnvioSemAprov As Integer
Private mvariPCExibeDeliveryDate As Integer

Public Property Let iPCExibeDeliveryDate(ByVal vData As Integer)
    mvariPCExibeDeliveryDate = vData
End Property

Public Property Get iPCExibeDeliveryDate() As Integer
    iPCExibeDeliveryDate = mvariPCExibeDeliveryDate
End Property

Public Property Let iPedCompraBloqEnvioSemAprov(ByVal vData As Integer)
    mvariPedCompraBloqEnvioSemAprov = vData
End Property

Public Property Get iPedCompraBloqEnvioSemAprov() As Integer
    iPedCompraBloqEnvioSemAprov = mvariPedCompraBloqEnvioSemAprov
End Property

Public Property Let iReqCompraDescProdBloq(ByVal vData As Integer)
    mvariReqCompraDescProdBloq = vData
End Property

Public Property Get iReqCompraDescProdBloq() As Integer
    iReqCompraDescProdBloq = mvariReqCompraDescProdBloq
End Property

Public Property Let iEmailAlteraDataEnvioPC(ByVal vData As Integer)
    mvariEmailAlteraDataEnvioPC = vData
End Property

Public Property Get iEmailAlteraDataEnvioPC() As Integer
    iEmailAlteraDataEnvioPC = mvariEmailAlteraDataEnvioPC
End Property

Public Property Let sFormatoPrecoUnitario(ByVal vData As String)
    mvarsFormatoPrecoUnitario = vData
End Property

Public Property Get sFormatoPrecoUnitario() As String
    sFormatoPrecoUnitario = mvarsFormatoPrecoUnitario
End Property

Public Property Let iNFDiferentePC(ByVal vData As Integer)
Attribute iNFDiferentePC.VB_Description = "Guarda a opção da empresa de aceitar ounão diferença entre os valores do Pedido de Compras e das Notas Fiscais (Compras, Fatura Compras e Rececimento de Compras)."
    mvariNFDiferentePC = vData
End Property
Public Property Get iNFDiferentePC() As Integer
    iNFDiferentePC = mvariNFDiferentePC
End Property

Public Function Inicializa() As Long

Dim objConfiguraCOM As New ClassConfiguraCOM
Dim lErro As Long

On Error GoTo Erro_Inicializa
   
    'Lê as configurações em ComprasConfig
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then Error 33651
    
    mvariFilialCompra = objConfiguraCOM.iFilialCompra
    mvariCompradorAumentaQuant = objConfiguraCOM.iCompradorAumentaQuant
    mvariConsideraQuantCotacaoAnterior = objConfiguraCOM.iConsideraQuantCotacaoAnterior
    mvardConsumoMedioMax = objConfiguraCOM.dConsumoMedioMax
    mvariControleAlcada = objConfiguraCOM.iControleAlcada
    mvariMesesConsumoMedio = objConfiguraCOM.iMesesConsumoMedio
    mvariMesesMediaTempoRessup = objConfiguraCOM.iMesesMediaTempoRessup
    mvardPercentMaisQuantCotacaoAnterior = objConfiguraCOM.dPercentMaisQuantCotacaoAnterior
    mvardPercentMaisReceb = objConfiguraCOM.dPercentMaisReceb
    mvardPercentMenosQuantCotacaoAnterior = objConfiguraCOM.dPercentMenosQuantCotacaoAnterior
    mvardPercentMenosReceb = objConfiguraCOM.dPercentMenosReceb
    mvariRecebForaFaixa = objConfiguraCOM.iRecebForaFaixa
    mvardResiduo = objConfiguraCOM.dResiduo
    mvardTaxaFinanceiraEmpresa = objConfiguraCOM.dTaxaFinanceiraEmpresa
    mvariTemFaixaReceb = objConfiguraCOM.iTemFaixaReceb
    mvardTempoRessupMax = objConfiguraCOM.dTempoRessupMax
    mvariNumComprasMediaAtraso = objConfiguraCOM.iNumComprasMediaAtraso
    mvariNumComprasTempoRessup = objConfiguraCOM.iNumComprasTempoRessup
    mvariNFDiferentePC = objConfiguraCOM.iNFDiferentePC
    mvarsFormatoPrecoUnitario = objConfiguraCOM.sFormatoPrecoUnitario 'Inserido por Wagner
    mvariEmailAlteraDataEnvioPC = objConfiguraCOM.iEmailAlteraDataEnvioPC
    mvariReqCompraDescProdBloq = objConfiguraCOM.iReqCompraDescProdBloq
    mvariPedCompraBloqEnvioSemAprov = objConfiguraCOM.iPedCompraBloqEnvioSemAprov
    mvariPCExibeDeliveryDate = objConfiguraCOM.iPCExibeDeliveryDate
    
    Inicializa = SUCESSO
    
    Exit Function
    
Erro_Inicializa:

    Inicializa = Err
    
    Select Case Err
    
        Case 33651
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRASCONFIG", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147429)
        
    End Select

    Exit Function

End Function
Function Gravar() As Long

    'DECIDIR O QUE GRAVAR (instalação)

End Function

Function Gravar_Trans() As Long
  'Grava no bd a configuração atual
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, colConfig As New ColComprasConfig
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim X As New ClassCOMGlobal

On Error GoTo Erro_Gravar_Trans

    objConfiguraCOM.dConsumoMedioMax = X.gobjCOM.dConsumoMedioMax
    objConfiguraCOM.dPercentMaisQuantCotacaoAnterior = X.gobjCOM.dPercentMaisQuantCotacaoAnterior
    objConfiguraCOM.dPercentMaisReceb = X.gobjCOM.dPercentMaisReceb
    objConfiguraCOM.dPercentMenosQuantCotacaoAnterior = X.gobjCOM.dPercentMenosQuantCotacaoAnterior
    objConfiguraCOM.dPercentMenosReceb = X.gobjCOM.dPercentMenosReceb
    objConfiguraCOM.dResiduo = X.gobjCOM.dResiduo
    objConfiguraCOM.dTaxaFinanceiraEmpresa = X.gobjCOM.dTaxaFinanceiraEmpresa
    objConfiguraCOM.dTempoRessupMax = X.gobjCOM.dTempoRessupMax
    objConfiguraCOM.iCompradorAumentaQuant = X.gobjCOM.iCompradorAumentaQuant
    objConfiguraCOM.iConsideraQuantCotacaoAnterior = X.gobjCOM.iConsideraQuantCotacaoAnterior
    objConfiguraCOM.iControleAlcada = X.gobjCOM.iControleAlcada
    objConfiguraCOM.iFilialCompra = X.gobjCOM.iFilialCompra
    objConfiguraCOM.iMesesConsumoMedio = X.gobjCOM.iMesesConsumoMedio
    objConfiguraCOM.iMesesMediaTempoRessup = X.gobjCOM.iMesesMediaTempoRessup
    objConfiguraCOM.iNFDiferentePC = X.gobjCOM.iNFDiferentePC
    objConfiguraCOM.iNumComprasMediaAtraso = X.gobjCOM.iNumComprasMediaAtraso
    objConfiguraCOM.iNumComprasTempoRessup = X.gobjCOM.iNumComprasTempoRessup
    objConfiguraCOM.iRecebForaFaixa = X.gobjCOM.iRecebForaFaixa
    objConfiguraCOM.iTemFaixaReceb = X.gobjCOM.iTemFaixaReceb
    objConfiguraCOM.iReqCompraDescProdBloq = X.gobjCOM.iReqCompraDescProdBloq
    objConfiguraCOM.iPedCompraBloqEnvioSemAprov = X.gobjCOM.iPedCompraBloqEnvioSemAprov
    objConfiguraCOM.iPCExibeDeliveryDate = X.gobjCOM.iPCExibeDeliveryDate
    
    lErro = CF("ConfiguraCOM_GravarTrans", objConfiguraCOM)
    If lErro <> SUCESSO Then Error 41817

    Gravar_Trans = SUCESSO

    Exit Function

Erro_Gravar_Trans:

    Gravar_Trans = Err

    Select Case Err

        Case 41817

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147430)

    End Select

    Exit Function

End Function


Public Property Let iNumComprasTempoRessup(ByVal vData As Integer)
    mvariNumComprasTempoRessup = vData
End Property
Public Property Get iNumComprasTempoRessup() As Integer
    iNumComprasTempoRessup = mvariNumComprasTempoRessup
End Property
Public Property Let iNumComprasMediaAtraso(ByVal vData As Integer)
    mvariNumComprasMediaAtraso = vData
End Property
Public Property Get iNumComprasMediaAtraso() As Integer
    iNumComprasMediaAtraso = mvariNumComprasMediaAtraso
End Property
Public Property Let dTempoRessupMax(ByVal vData As Double)
    mvardTempoRessupMax = vData
End Property
Public Property Get dTempoRessupMax() As Double
    dTempoRessupMax = mvardTempoRessupMax
End Property
Public Property Let iTemFaixaReceb(ByVal vData As Integer)
    mvariTemFaixaReceb = vData
End Property
Public Property Get iTemFaixaReceb() As Integer
    iTemFaixaReceb = mvariTemFaixaReceb
End Property
Public Property Let dTaxaFinanceiraEmpresa(ByVal vData As Double)
    mvardTaxaFinanceiraEmpresa = vData
End Property
Public Property Get dTaxaFinanceiraEmpresa() As Double
    dTaxaFinanceiraEmpresa = mvardTaxaFinanceiraEmpresa
End Property
Public Property Let dResiduo(ByVal vData As Double)
    mvardResiduo = vData
End Property
Public Property Get dResiduo() As Double
    dResiduo = mvardResiduo
End Property
Public Property Let iRecebForaFaixa(ByVal vData As Integer)
    mvariRecebForaFaixa = vData
End Property
Public Property Get iRecebForaFaixa() As Integer
    iRecebForaFaixa = mvariRecebForaFaixa
End Property
Public Property Let dPercentMenosReceb(ByVal vData As Double)
    mvardPercentMenosReceb = vData
End Property
Public Property Get dPercentMenosReceb() As Double
    dPercentMenosReceb = mvardPercentMenosReceb
End Property
Public Property Let dPercentMenosQuantCotacaoAnterior(ByVal vData As Double)
    mvardPercentMenosQuantCotacaoAnterior = vData
End Property
Public Property Get dPercentMenosQuantCotacaoAnterior() As Double
    dPercentMenosQuantCotacaoAnterior = mvardPercentMenosQuantCotacaoAnterior
End Property
Public Property Let dPercentMaisReceb(ByVal vData As Double)
    mvardPercentMaisReceb = vData
End Property
Public Property Get dPercentMaisReceb() As Double
    dPercentMaisReceb = mvardPercentMaisReceb
End Property
Public Property Let dPercentMaisQuantCotacaoAnterior(ByVal vData As Double)
    mvardPercentMaisQuantCotacaoAnterior = vData
End Property
Public Property Get dPercentMaisQuantCotacaoAnterior() As Double
    dPercentMaisQuantCotacaoAnterior = mvardPercentMaisQuantCotacaoAnterior
End Property
Public Property Let iMesesMediaTempoRessup(ByVal vData As Integer)
    mvariMesesMediaTempoRessup = vData
End Property
Public Property Get iMesesMediaTempoRessup() As Integer
    iMesesMediaTempoRessup = mvariMesesMediaTempoRessup
End Property
Public Property Let iMesesConsumoMedio(ByVal vData As Integer)
    mvariMesesConsumoMedio = vData
End Property
Public Property Get iMesesConsumoMedio() As Integer
    iMesesConsumoMedio = mvariMesesConsumoMedio
End Property
Public Property Let iControleAlcada(ByVal vData As Integer)
    mvariControleAlcada = vData
End Property
Public Property Get iControleAlcada() As Integer
    iControleAlcada = mvariControleAlcada
End Property
Public Property Let dConsumoMedioMax(ByVal vData As Double)
    mvardConsumoMedioMax = vData
End Property
Public Property Get dConsumoMedioMax() As Double
    dConsumoMedioMax = mvardConsumoMedioMax
End Property
Public Property Let iConsideraQuantCotacaoAnterior(ByVal vData As Integer)
    mvariConsideraQuantCotacaoAnterior = vData
End Property
Public Property Get iConsideraQuantCotacaoAnterior() As Integer
    iConsideraQuantCotacaoAnterior = mvariConsideraQuantCotacaoAnterior
End Property
Public Property Let iCompradorAumentaQuant(ByVal vData As Integer)
    mvariCompradorAumentaQuant = vData
End Property
Public Property Get iCompradorAumentaQuant() As Integer
    iCompradorAumentaQuant = mvariCompradorAumentaQuant
End Property
Public Property Let iFilialCompra(ByVal vData As Integer)
    mvariFilialCompra = vData
End Property
Public Property Get iFilialCompra() As Integer
    iFilialCompra = mvariFilialCompra
End Property


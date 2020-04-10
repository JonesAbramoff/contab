VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFilialCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarlCodCliente As Long
Private mvariCodFilial As Integer
Private mvarsNome As String
Private mvarsCGC As String
Private mvarsInscricaoEstadual As String
Private mvarsInscricaoMunicipal As String
Private mvarlEndereco As Long
Private mvarlEnderecoEntrega As Long
Private mvarlEnderecoCobranca As Long
Private mvariCodTransportadora As Integer
Private mvarsObservacao As String
Private mvarsContaContabil As String
Private mvariVendedor As Integer
Private mvardComissaoVendas As Double
Private mvariRegiao As Integer
Private mvariFreqVisitas As Integer
Private mvardtDataUltVisita As Date
Private mvariCodCobrador As Integer
Private mvarcolCategoriaItem As New Collection
Private mvariTipoFrete As Integer
Private mvarsNomeReduzidoCli As String
Private mvarsRazaoSocialCli As String
Private mvarlRevendedor As Long
Private mvariICMSBaseCalculoComIPI As Integer
Private mvarsInscricaoSuframa As String
Private mvarsRG As String
Private mvariFilialEmpresaLoja As Integer
Private mvariCodFilialLoja As Integer
Private mvarlCodClienteLoja As Long
Private mvariAtivo As Integer
Private mvariTransferido As Integer
Private mvariTipo As Integer
Private mvardLimiteCredito As Double
Private mvarsGuia As String
Private mvariCodTranspRedesp As Integer
Private mvariDetPagFrete As Integer
Private mvariCodMensagem As Integer
Private mvarlCodExterno As Long
Private mvariRegimeTributario As Integer
Private mvariIEIsento As Integer
Private mvariIENaoContrib As Integer

'nfe 3.10
Public sIdEstrangeiro As String

Public Property Let iIENaoContrib(ByVal vData As Integer)
    mvariIENaoContrib = vData
End Property

Public Property Get iIENaoContrib() As Integer
    If mvariIENaoContrib = -1 Then
        iIENaoContrib = IIf(Len(Trim(mvarsInscricaoEstadual)) = 0, MARCADO, DESMARCADO)
    Else
        iIENaoContrib = mvariIENaoContrib
    End If
End Property

Public Property Let iIEIsento(ByVal vData As Integer)
    mvariIEIsento = vData
End Property

Public Property Get iIEIsento() As Integer
    If mvariIEIsento = -1 Then
        iIEIsento = IIf(Len(Trim(mvarsInscricaoEstadual)) = 0, MARCADO, DESMARCADO)
    Else
        iIEIsento = mvariIEIsento
    End If
End Property

Public Property Let iRegimeTributario(ByVal vData As Integer)
    mvariRegimeTributario = vData
End Property

Public Property Get iRegimeTributario() As Integer
    iRegimeTributario = mvariRegimeTributario
End Property

Public Property Let lCodExterno(ByVal vData As Long)
    mvarlCodExterno = vData
End Property

Public Property Get lCodExterno() As Long
    lCodExterno = mvarlCodExterno
End Property

Public Property Let iCodMensagem(ByVal vData As Integer)
    mvariCodMensagem = vData
End Property

Public Property Get iCodMensagem() As Integer
    iCodMensagem = mvariCodMensagem
End Property

Public Property Let iDetPagFrete(ByVal vData As Integer)
    mvariDetPagFrete = vData
End Property

Public Property Get iDetPagFrete() As Integer
    iDetPagFrete = mvariDetPagFrete
End Property

Public Property Let iCodTranspRedesp(ByVal vData As Integer)
    mvariCodTranspRedesp = vData
End Property

Public Property Get iCodTranspRedesp() As Integer
    iCodTranspRedesp = mvariCodTranspRedesp
End Property

Public Property Let sGuia(ByVal vData As String)
    mvarsGuia = vData
End Property

Public Property Get sGuia() As String
    sGuia = mvarsGuia
End Property

Public Property Let dLimiteCredito(ByVal vData As Double)
    mvardLimiteCredito = vData
End Property

Public Property Get dLimiteCredito() As Double
    dLimiteCredito = mvardLimiteCredito
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iTransferido(ByVal vData As Integer)
    mvariTransferido = vData
End Property

Public Property Get iTransferido() As Integer
    iTransferido = mvariTransferido
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let lCodClienteLoja(ByVal vData As Long)
    mvarlCodClienteLoja = vData
End Property

Public Property Get lCodClienteLoja() As Long
    lCodClienteLoja = mvarlCodClienteLoja
End Property

Public Property Let iCodFilialLoja(ByVal vData As Integer)
    mvariCodFilialLoja = vData
End Property

Public Property Get iCodFilialLoja() As Integer
    iCodFilialLoja = mvariCodFilialLoja
End Property

Public Property Let iFilialEmpresaLoja(ByVal vData As Integer)
    mvariFilialEmpresaLoja = vData
End Property

Public Property Get iFilialEmpresaLoja() As Integer
    iFilialEmpresaLoja = mvariFilialEmpresaLoja
End Property

Public Property Let sRG(ByVal vData As String)
    mvarsRG = vData
End Property

Public Property Get sRG() As String
    sRG = mvarsRG
End Property

Public Property Let sInscricaoSuframa(ByVal vData As String)
    mvarsInscricaoSuframa = vData
End Property

Public Property Get sInscricaoSuframa() As String
    sInscricaoSuframa = mvarsInscricaoSuframa
End Property

Public Property Let iICMSBaseCalculoComIPI(ByVal vData As Integer)
    mvariICMSBaseCalculoComIPI = vData
End Property

Public Property Get iICMSBaseCalculoComIPI() As Integer
    iICMSBaseCalculoComIPI = mvariICMSBaseCalculoComIPI
End Property

Public Property Let lRevendedor(ByVal vData As Long)
    mvarlRevendedor = vData
End Property

Public Property Get lRevendedor() As Long
    lRevendedor = mvarlRevendedor
End Property

Public Property Let sRazaoSocialCli(ByVal vData As String)
    mvarsRazaoSocialCli = vData
End Property

Public Property Get sRazaoSocialCli() As String
    sRazaoSocialCli = mvarsRazaoSocialCli
End Property

Public Property Let sNomeReduzidoCli(ByVal vData As String)
    mvarsNomeReduzidoCli = vData
End Property

Public Property Get sNomeReduzidoCli() As String
    sNomeReduzidoCli = mvarsNomeReduzidoCli
End Property

Public Property Let iTipoFrete(ByVal vData As Integer)
    mvariTipoFrete = vData
End Property

Public Property Get iTipoFrete() As Integer
    iTipoFrete = mvariTipoFrete
End Property

Public Property Let iCodCobrador(ByVal vData As Integer)
    mvariCodCobrador = vData
End Property

Public Property Get iCodCobrador() As Integer
    iCodCobrador = mvariCodCobrador
End Property

Public Property Let dtDataUltVisita(ByVal vData As Date)
    mvardtDataUltVisita = vData
End Property

Public Property Get dtDataUltVisita() As Date
    dtDataUltVisita = mvardtDataUltVisita
End Property

Public Property Let iFreqVisitas(ByVal vData As Integer)
    mvariFreqVisitas = vData
End Property

Public Property Get iFreqVisitas() As Integer
    iFreqVisitas = mvariFreqVisitas
End Property

Public Property Let iRegiao(ByVal vData As Integer)
    mvariRegiao = vData
End Property

Public Property Get iRegiao() As Integer
    iRegiao = mvariRegiao
End Property

Public Property Let dComissaoVendas(ByVal vData As Double)
    mvardComissaoVendas = vData
End Property

Public Property Get dComissaoVendas() As Double
    dComissaoVendas = mvardComissaoVendas
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let sContaContabil(ByVal vData As String)
    mvarsContaContabil = vData
End Property

Public Property Get sContaContabil() As String
    sContaContabil = mvarsContaContabil
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let iCodTransportadora(ByVal vData As Integer)
    mvariCodTransportadora = vData
End Property

Public Property Get iCodTransportadora() As Integer
    iCodTransportadora = mvariCodTransportadora
End Property

Public Property Let lEnderecoCobranca(ByVal vData As Long)
    mvarlEnderecoCobranca = vData
End Property

Public Property Get lEnderecoCobranca() As Long
    lEnderecoCobranca = mvarlEnderecoCobranca
End Property

Public Property Let lEnderecoEntrega(ByVal vData As Long)
    mvarlEnderecoEntrega = vData
End Property

Public Property Get lEnderecoEntrega() As Long
    lEnderecoEntrega = mvarlEnderecoEntrega
End Property

Public Property Let lEndereco(ByVal vData As Long)
    mvarlEndereco = vData
End Property

Public Property Get lEndereco() As Long
    lEndereco = mvarlEndereco
End Property

Public Property Let sInscricaoMunicipal(ByVal vData As String)
    mvarsInscricaoMunicipal = vData
End Property

Public Property Get sInscricaoMunicipal() As String
    sInscricaoMunicipal = mvarsInscricaoMunicipal
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let sCgc(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

Public Property Let lCodCliente(ByVal vData As Long)
    mvarlCodCliente = vData
End Property

Public Property Get lCodCliente() As Long
    lCodCliente = mvarlCodCliente
End Property

Public Property Set colCategoriaItem(ByVal vData As Object)
    Set mvarcolCategoriaItem = vData
End Property

Public Property Get colCategoriaItem() As Collection
    Set colCategoriaItem = mvarcolCategoriaItem
End Property


Private Sub Class_Initialize()
    mvariIENaoContrib = -1
    mvariIEIsento = -1

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassConhecimentoFrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntNFiscal As Long 'local copy
Private mvardFretePeso As Double 'local copy
Private mvardFreteValor As Double 'local copy
Private mvardSEC As Double 'local copy
Private mvardDespacho As Double 'local copy
Private mvardPedagio As Double 'local copy
Private mvardOutrosValores As Double 'local copy
Private mvardAliquotas As Double 'local copy
Private mvardBaseCalculo As Double 'local copy
Private mvardValorTotal As Double 'local copy
Private mvardValorICMS As Double 'local copy
Private mvardPesoMercadoria As Double 'local copy
Private mvardValorMercadoria As Double 'local copy
Private mvarsNotasFiscais As String 'local copy
Private mvarsObservacao As String 'local copy
Private mvarsColeta As String 'local copy
Private mvarsEntrega As String 'local copy
Private mvarsCalculadoAte As String 'local copy
Private mvarsNaturezaCarga As String 'local copy
Private mvarsLocalVeiculo As String 'local copy
Private mvarsRemetente As String 'local copy
Private mvarsEnderecoRemetente As String 'local copy
Private mvarsMunicipioRemetente As String 'local copy
Private mvarsUFRemetente As String 'local copy
Private mvarsCepRemetente As String 'local copy
Private mvarsCGCRemetente As String 'local copy
Private mvarsInscEstadualRemetente As String 'local copy
Private mvarsDestinatario As String 'local copy
Private mvarsEnderecoDestinatario As String 'local copy
Private mvarsMunicipioDestinatario As String 'local copy
Private mvarsUFDestinatario As String 'local copy
Private mvarsCepDestinatario As String 'local copy
Private mvarsCGCDestinatario As String 'local copy
Private mvarsInscEstadualDestinatario As String 'local copy
Private mvarsMarcaVeiculo As String 'local copy
Private mvariICMSIncluso As Integer 'local copy
Private mvardValorINSS As Double
Private mvariINSSRetido As Integer
Private mvariIncluiPedagio As Integer

Public Property Let iIncluiPedagio(ByVal vData As Integer)
    mvariIncluiPedagio = vData
End Property

Public Property Get iIncluiPedagio() As Integer
    iIncluiPedagio = mvariIncluiPedagio
End Property

Public Property Let iINSSRetido(ByVal vData As Integer)
    mvariINSSRetido = vData
End Property

Public Property Get iINSSRetido() As Integer
    iINSSRetido = mvariINSSRetido
End Property

Public Property Let dValorINSS(ByVal vData As Double)
    mvardValorINSS = vData
End Property

Public Property Get dValorINSS() As Double
    dValorINSS = mvardValorINSS
End Property

Public Property Let iICMSIncluso(ByVal vData As Integer)
    mvariICMSIncluso = vData
End Property

Public Property Get iICMSIncluso() As Integer
    iICMSIncluso = mvariICMSIncluso
End Property
Public Property Let sMarcaVeiculo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sMarcaVeiculo = 5
    mvarsMarcaVeiculo = vData
End Property


Public Property Get sMarcaVeiculo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sMarcaVeiculo
    sMarcaVeiculo = mvarsMarcaVeiculo
End Property



Public Property Let sInscEstadualDestinatario(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sInscEstadualDestinatario = 5
    mvarsInscEstadualDestinatario = vData
End Property


Public Property Get sInscEstadualDestinatario() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sInscEstadualDestinatario
    sInscEstadualDestinatario = mvarsInscEstadualDestinatario
End Property



Public Property Let sCGCDestinatario(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCGCDestinatario = 5
    mvarsCGCDestinatario = vData
End Property


Public Property Get sCGCDestinatario() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCGCDestinatario
    sCGCDestinatario = mvarsCGCDestinatario
End Property



Public Property Let sCepDestinatario(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCepDestinatario = 5
    mvarsCepDestinatario = vData
End Property


Public Property Get sCepDestinatario() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCepDestinatario
    sCepDestinatario = mvarsCepDestinatario
End Property



Public Property Let sUFDestinatario(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUFDestinatario = 5
    mvarsUFDestinatario = vData
End Property


Public Property Get sUFDestinatario() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUFDestinatario
    sUFDestinatario = mvarsUFDestinatario
End Property



Public Property Let sMunicipioDestinatario(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sMunicipioDestinatario = 5
    mvarsMunicipioDestinatario = vData
End Property


Public Property Get sMunicipioDestinatario() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sMunicipioDestinatario
    sMunicipioDestinatario = mvarsMunicipioDestinatario
End Property



Public Property Let sEnderecoDestinatario(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sEnderecoDestinatario = 5
    mvarsEnderecoDestinatario = vData
End Property


Public Property Get sEnderecoDestinatario() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sEnderecoDestinatario
    sEnderecoDestinatario = mvarsEnderecoDestinatario
End Property



Public Property Let sDestinatario(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDestinatario = 5
    mvarsDestinatario = vData
End Property


Public Property Get sDestinatario() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDestinatario
    sDestinatario = mvarsDestinatario
End Property



Public Property Let sInscEstadualRemetente(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sInscEstadualRemetente = 5
    mvarsInscEstadualRemetente = vData
End Property


Public Property Get sInscEstadualRemetente() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sInscEstadualRemetente
    sInscEstadualRemetente = mvarsInscEstadualRemetente
End Property



Public Property Let sCGCRemetente(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCGCRemetente = 5
    mvarsCGCRemetente = vData
End Property


Public Property Get sCGCRemetente() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCGCRemetente
    sCGCRemetente = mvarsCGCRemetente
End Property



Public Property Let sCepRemetente(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCepRemetente = 5
    mvarsCepRemetente = vData
End Property


Public Property Get sCepRemetente() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCepRemetente
    sCepRemetente = mvarsCepRemetente
End Property



Public Property Let sUFRemetente(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUFRemetente = 5
    mvarsUFRemetente = vData
End Property


Public Property Get sUFRemetente() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUFRemetente
    sUFRemetente = mvarsUFRemetente
End Property



Public Property Let sMunicipioRemetente(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sMunicipioRemetente = 5
    mvarsMunicipioRemetente = vData
End Property


Public Property Get sMunicipioRemetente() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sMunicipioRemetente
    sMunicipioRemetente = mvarsMunicipioRemetente
End Property



Public Property Let sEnderecoRemetente(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sEnderecoRemetente = 5
    mvarsEnderecoRemetente = vData
End Property


Public Property Get sEnderecoRemetente() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sEnderecoRemetente
    sEnderecoRemetente = mvarsEnderecoRemetente
End Property



Public Property Let sRemetente(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sRemetente = 5
    mvarsRemetente = vData
End Property


Public Property Get sRemetente() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sRemetente
    sRemetente = mvarsRemetente
End Property



Public Property Let sLocalVeiculo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sLocalVeiculo = 5
    mvarsLocalVeiculo = vData
End Property


Public Property Get sLocalVeiculo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sLocalVeiculo
    sLocalVeiculo = mvarsLocalVeiculo
End Property



Public Property Let sNaturezaCarga(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNaturezaCarga = 5
    mvarsNaturezaCarga = vData
End Property


Public Property Get sNaturezaCarga() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNaturezaCarga
    sNaturezaCarga = mvarsNaturezaCarga
End Property



Public Property Let sCalculadoAte(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCalculadoAte = 5
    mvarsCalculadoAte = vData
End Property


Public Property Get sCalculadoAte() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCalculadoAte
    sCalculadoAte = mvarsCalculadoAte
End Property



Public Property Let sEntrega(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sEntrega = 5
    mvarsEntrega = vData
End Property


Public Property Get sEntrega() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sEntrega
    sEntrega = mvarsEntrega
End Property



Public Property Let sColeta(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sColeta = 5
    mvarsColeta = vData
End Property


Public Property Get sColeta() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sColeta
    sColeta = mvarsColeta
End Property



Public Property Let sObservacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sObservacao = 5
    mvarsObservacao = vData
End Property


Public Property Get sObservacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sObservacao
    sObservacao = mvarsObservacao
End Property



Public Property Let sNotasFiscais(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNotasFiscais = 5
    mvarsNotasFiscais = vData
End Property


Public Property Get sNotasFiscais() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNotasFiscais
    sNotasFiscais = mvarsNotasFiscais
End Property



Public Property Let dValorMercadoria(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorMercadoria = 5
    mvardValorMercadoria = vData
End Property


Public Property Get dValorMercadoria() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorMercadoria
    dValorMercadoria = mvardValorMercadoria
End Property



Public Property Let dPesoMercadoria(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPesoMercadoria = 5
    mvardPesoMercadoria = vData
End Property


Public Property Get dPesoMercadoria() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPesoMercadoria
    dPesoMercadoria = mvardPesoMercadoria
End Property



Public Property Let dValorICMS(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorICMS = 5
    mvardValorICMS = vData
End Property


Public Property Get dValorICMS() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorICMS
    dValorICMS = mvardValorICMS
End Property



Public Property Let dValorTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorTotal = 5
    mvardValorTotal = vData
End Property


Public Property Get dValorTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorTotal
    dValorTotal = mvardValorTotal
End Property



Public Property Let dBaseCalculo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dBaseCalculo = 5
    mvardBaseCalculo = vData
End Property


Public Property Get dBaseCalculo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dBaseCalculo
    dBaseCalculo = mvardBaseCalculo
End Property



Public Property Let dAliquotas(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dAliquotas = 5
    mvardAliquotas = vData
End Property


Public Property Get dAliquotas() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dAliquotas
    dAliquotas = mvardAliquotas
End Property



Public Property Let dOutrosValores(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dOutrosValores = 5
    mvardOutrosValores = vData
End Property


Public Property Get dOutrosValores() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dOutrosValores
    dOutrosValores = mvardOutrosValores
End Property



Public Property Let dPedagio(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPedagio = 5
    mvardPedagio = vData
End Property


Public Property Get dPedagio() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPedagio
    dPedagio = mvardPedagio
End Property



Public Property Let dDespacho(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dDespacho = 5
    mvardDespacho = vData
End Property


Public Property Get dDespacho() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dDespacho
    dDespacho = mvardDespacho
End Property



Public Property Let dSEC(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dSEC = 5
    mvardSEC = vData
End Property


Public Property Get dSEC() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dSEC
    dSEC = mvardSEC
End Property



Public Property Let dFreteValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dFreteValor = 5
    mvardFreteValor = vData
End Property


Public Property Get dFreteValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dFreteValor
    dFreteValor = mvardFreteValor
End Property



Public Property Let dFretePeso(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dFretePeso = 5
    mvardFretePeso = vData
End Property


Public Property Get dFretePeso() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dFretePeso
    dFretePeso = mvardFretePeso
End Property



Public Property Let lNumIntNFiscal(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntNFiscal = 5
    mvarlNumIntNFiscal = vData
End Property


Public Property Get lNumIntNFiscal() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntNFiscal
    lNumIntNFiscal = mvarlNumIntNFiscal
End Property




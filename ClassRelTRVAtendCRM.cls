VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelTRVAtendCRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel As Long
Private mvariSeq As Integer
Private mvariAtendenteCod As Integer
Private mvarsAtendenteNome As String
Private mvarsAtendenteCodUsu As String
Private mvarlTotalCli As Long
Private mvarlTotalVou As Long
Private mvarlTotalVouCanc As Long
Private mvarlTotalPax As Long
Private mvarlTotalCliInat As Long
Private mvarlTotalCliNovos As Long
Private mvarlTotalCliReativ As Long
Private mvarlTotalCliInatReativ As Long
Private mvarlTotalCliContact As Long
Private mvarlTotalCliInatContact As Long
Private mvarlTotalContRealizados As Long
Private mvardPercCliReativ As Double
Private mvardPercCliContact As Double
Private mvarlMaiorCliCod As Long
Private mvarsMaiorCliNome As String
Private mvardMaiorCliVendFat As Double
Private mvardMaiorCliVendLiq As Double
Private mvardMaiorCliVendBruto As Double
Private mvarsMaiorProdCod As String
Private mvarsMaiorProdDesc As String
Private mvarlMaiorProdQtd As Long
Private mvardVendaMediaFat As Double
Private mvardVendaMediaLiq As Double
Private mvardVendaMediaBruto As Double
Private mvardPercDescMedio As Double
Private mvardTotalVendFat As Double
Private mvardTotalVendLiq As Double
Private mvardTotalVendBruto As Double
Private mvardPercCancMedio As Double
Private mvardTotalInvestido As Double
Private mvarlTotalContAtivo As Long
Private mvarlTotalContReceptivo As Long
Private mvarlTotalPDV As Long
    
Public Property Let lTotalPDV(ByVal vData As Long)
    mvarlTotalPDV = vData
End Property

Public Property Get lTotalPDV() As Long
    lTotalPDV = mvarlTotalPDV
End Property

Public Property Let lTotalContReceptivo(ByVal vData As Long)
    mvarlTotalContReceptivo = vData
End Property

Public Property Get lTotalContReceptivo() As Long
    lTotalContReceptivo = mvarlTotalContReceptivo
End Property

Public Property Let lTotalContAtivo(ByVal vData As Long)
    mvarlTotalContAtivo = vData
End Property

Public Property Get lTotalContAtivo() As Long
    lTotalContAtivo = mvarlTotalContAtivo
End Property


Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iAtendenteCod(ByVal vData As Integer)
    mvariAtendenteCod = vData
End Property

Public Property Get iAtendenteCod() As Integer
    iAtendenteCod = mvariAtendenteCod
End Property

Public Property Let sAtendenteCodUsu(ByVal vData As String)
    mvarsAtendenteCodUsu = vData
End Property

Public Property Get sAtendenteCodUsu() As String
    sAtendenteCodUsu = mvarsAtendenteCodUsu
End Property

Public Property Let sAtendenteNome(ByVal vData As String)
    mvarsAtendenteNome = vData
End Property

Public Property Get sAtendenteNome() As String
    sAtendenteNome = mvarsAtendenteNome
End Property

Public Property Let lTotalCli(ByVal vData As Long)
    mvarlTotalCli = vData
End Property

Public Property Get lTotalCli() As Long
    lTotalCli = mvarlTotalCli
End Property

Public Property Let lTotalVou(ByVal vData As Long)
    mvarlTotalVou = vData
End Property

Public Property Get lTotalVou() As Long
    lTotalVou = mvarlTotalVou
End Property

Public Property Let lTotalVouCanc(ByVal vData As Long)
    mvarlTotalVouCanc = vData
End Property

Public Property Get lTotalVouCanc() As Long
    lTotalVouCanc = mvarlTotalVouCanc
End Property

Public Property Let lTotalPax(ByVal vData As Long)
    mvarlTotalPax = vData
End Property

Public Property Get lTotalPax() As Long
    lTotalPax = mvarlTotalPax
End Property

Public Property Let lTotalCliInat(ByVal vData As Long)
    mvarlTotalCliInat = vData
End Property

Public Property Get lTotalCliInat() As Long
    lTotalCliInat = mvarlTotalCliInat
End Property

Public Property Let lTotalCliNovos(ByVal vData As Long)
    mvarlTotalCliNovos = vData
End Property

Public Property Get lTotalCliNovos() As Long
    lTotalCliNovos = mvarlTotalCliNovos
End Property

Public Property Let lTotalCliReativ(ByVal vData As Long)
    mvarlTotalCliReativ = vData
End Property

Public Property Get lTotalCliReativ() As Long
    lTotalCliReativ = mvarlTotalCliReativ
End Property

Public Property Let lTotalCliInatReativ(ByVal vData As Long)
    mvarlTotalCliInatReativ = vData
End Property

Public Property Get lTotalCliInatReativ() As Long
    lTotalCliInatReativ = mvarlTotalCliInatReativ
End Property

Public Property Let lTotalCliContact(ByVal vData As Long)
    mvarlTotalCliContact = vData
End Property

Public Property Get lTotalCliContact() As Long
    lTotalCliContact = mvarlTotalCliContact
End Property

Public Property Let lTotalCliInatContact(ByVal vData As Long)
    mvarlTotalCliInatContact = vData
End Property

Public Property Get lTotalCliInatContact() As Long
    lTotalCliInatContact = mvarlTotalCliInatContact
End Property

Public Property Let lTotalContRealizados(ByVal vData As Long)
    mvarlTotalContRealizados = vData
End Property

Public Property Get lTotalContRealizados() As Long
    lTotalContRealizados = mvarlTotalContRealizados
End Property

Public Property Let dPercCliReativ(ByVal vData As Double)
    mvardPercCliReativ = vData
End Property

Public Property Get dPercCliReativ() As Double
    dPercCliReativ = mvardPercCliReativ
End Property

Public Property Let dPercCliContact(ByVal vData As Double)
    mvardPercCliContact = vData
End Property

Public Property Get dPercCliContact() As Double
    dPercCliContact = mvardPercCliContact
End Property

Public Property Let lMaiorCliCod(ByVal vData As Long)
    mvarlMaiorCliCod = vData
End Property

Public Property Get lMaiorCliCod() As Long
    lMaiorCliCod = mvarlMaiorCliCod
End Property

Public Property Let sMaiorCliNome(ByVal vData As String)
    mvarsMaiorCliNome = vData
End Property

Public Property Get sMaiorCliNome() As String
    sMaiorCliNome = mvarsMaiorCliNome
End Property

Public Property Let dMaiorCliVendFat(ByVal vData As Double)
    mvardMaiorCliVendFat = vData
End Property

Public Property Get dMaiorCliVendFat() As Double
    dMaiorCliVendFat = mvardMaiorCliVendFat
End Property

Public Property Let dMaiorCliVendLiq(ByVal vData As Double)
    mvardMaiorCliVendLiq = vData
End Property

Public Property Get dMaiorCliVendLiq() As Double
    dMaiorCliVendLiq = mvardMaiorCliVendLiq
End Property

Public Property Let dMaiorCliVendBruto(ByVal vData As Double)
    mvardMaiorCliVendBruto = vData
End Property

Public Property Get dMaiorCliVendBruto() As Double
    dMaiorCliVendBruto = mvardMaiorCliVendBruto
End Property

Public Property Let sMaiorProdCod(ByVal vData As String)
    mvarsMaiorProdCod = vData
End Property

Public Property Get sMaiorProdCod() As String
    sMaiorProdCod = mvarsMaiorProdCod
End Property

Public Property Let sMaiorProdDesc(ByVal vData As String)
    mvarsMaiorProdDesc = vData
End Property

Public Property Get sMaiorProdDesc() As String
    sMaiorProdDesc = mvarsMaiorProdDesc
End Property

Public Property Let lMaiorProdQtd(ByVal vData As Long)
    mvarlMaiorProdQtd = vData
End Property

Public Property Get lMaiorProdQtd() As Long
    lMaiorProdQtd = mvarlMaiorProdQtd
End Property

Public Property Let dVendaMediaFat(ByVal vData As Double)
    mvardVendaMediaFat = vData
End Property

Public Property Get dVendaMediaFat() As Double
    dVendaMediaFat = mvardVendaMediaFat
End Property

Public Property Let dVendaMediaLiq(ByVal vData As Double)
    mvardVendaMediaLiq = vData
End Property

Public Property Get dVendaMediaLiq() As Double
    dVendaMediaLiq = mvardVendaMediaLiq
End Property

Public Property Let dVendaMediaBruto(ByVal vData As Double)
    mvardVendaMediaBruto = vData
End Property

Public Property Get dVendaMediaBruto() As Double
    dVendaMediaBruto = mvardVendaMediaBruto
End Property

Public Property Let dPercDescMedio(ByVal vData As Double)
    mvardPercDescMedio = vData
End Property

Public Property Get dPercDescMedio() As Double
    dPercDescMedio = mvardPercDescMedio
End Property

Public Property Let dTotalVendFat(ByVal vData As Double)
    mvardTotalVendFat = vData
End Property

Public Property Get dTotalVendFat() As Double
    dTotalVendFat = mvardTotalVendFat
End Property

Public Property Let dTotalVendLiq(ByVal vData As Double)
    mvardTotalVendLiq = vData
End Property

Public Property Get dTotalVendLiq() As Double
    dTotalVendLiq = mvardTotalVendLiq
End Property

Public Property Let dTotalVendBruto(ByVal vData As Double)
    mvardTotalVendBruto = vData
End Property

Public Property Get dTotalVendBruto() As Double
    dTotalVendBruto = mvardTotalVendBruto
End Property

Public Property Let dPercCancMedio(ByVal vData As Double)
    mvardPercCancMedio = vData
End Property

Public Property Get dPercCancMedio() As Double
    dPercCancMedio = mvardPercCancMedio
End Property

Public Property Let dTotalInvestido(ByVal vData As Double)
    mvardTotalInvestido = vData
End Property

Public Property Get dTotalInvestido() As Double
    dTotalInvestido = mvardTotalInvestido
End Property

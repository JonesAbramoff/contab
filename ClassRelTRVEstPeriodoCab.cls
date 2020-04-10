VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassRelTRVEstPeriodoCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel AS Long
Private mvarlCQtdAg AS Long
Private mvarlPQtdAg AS Long
Private mvarlCQtdVou AS Long
Private mvarlPQtdVou AS Long
Private mvarlCQtdPax AS Long
Private mvarlPQtdPax AS Long
Private mvardCValorFat AS Double
Private mvardPValorFat AS Double
Private mvardCValorLiq AS Double
Private mvardPValorLiq AS Double
Private mvardCValorBruto AS Double
Private mvardPValorBruto AS Double
Private mvardCValorInvestido AS Double
Private mvardPValorInvestido AS Double
Private mvarlCCliNovos AS Long
Private mvarlPCliNovos AS Long
Private mvarlCCliReativ AS Long
Private mvarlPCliReativ AS Long
Private mvardCPercDescMedio AS Double
Private mvardPPercDescMedio AS Double
Private mvarsCMaiorProdCod AS String
Private mvarsPMaiorProdCod AS String
Private mvarsCMaiorProdDesc AS String
Private mvarsPMaiorProdDesc AS String
Private mvarlCMaiorProdQtd AS Long
Private mvarlPMaiorProdQtd AS Long
Private mvardCMaiorProdVlrFat AS Double
Private mvardPMaiorProdVlrFat AS Double
Private mvardCMaiorProdVlrLiq AS Double
Private mvardPMaiorProdVlrLiq AS Double
Private mvardCMaiorProdVlrBruto AS Double
Private mvardPMaiorProdVlrBruto AS Double
Private mvarlCContatosCall AS Long
Private mvarlPContatosCall AS Long
Private mvarlCContatosCobr AS Long
Private mvarlPContatosCobr AS Long
Private mvarlCContatosOutros AS Long
Private mvarlPContatosOutros AS Long
Private mvarlCMaiorCliCod AS Long
Private mvarlPMaiorCliCod AS Long
Private mvarsCMaiorCliNome AS String
Private mvarsPMaiorCliNome AS String
Private mvardCMaiorCliValorFat AS Double
Private mvardPMaiorCliValorFat AS Double
Private mvardCMaiorCliValorLiq AS Double
Private mvardPMaiorCliValorLiq AS Double
Private mvardCMaiorCliValorBruto AS Double
Private mvardPMaiorCliValorBruto AS Double
Private mvardCPercVouCanc AS Double
Private mvardPPercVouCanc AS Double

Public Property Let lNumIntRel (ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel () AS Long
    lNumIntRel= mvarlNumIntRel
End Property

Public Property Let lCQtdAg (ByVal vData As Long)
    mvarlCQtdAg = vData
End Property

Public Property Get lCQtdAg () AS Long
    lCQtdAg= mvarlCQtdAg
End Property

Public Property Let lPQtdAg (ByVal vData As Long)
    mvarlPQtdAg = vData
End Property

Public Property Get lPQtdAg () AS Long
    lPQtdAg= mvarlPQtdAg
End Property

Public Property Let lCQtdVou (ByVal vData As Long)
    mvarlCQtdVou = vData
End Property

Public Property Get lCQtdVou () AS Long
    lCQtdVou= mvarlCQtdVou
End Property

Public Property Let lPQtdVou (ByVal vData As Long)
    mvarlPQtdVou = vData
End Property

Public Property Get lPQtdVou () AS Long
    lPQtdVou= mvarlPQtdVou
End Property

Public Property Let lCQtdPax (ByVal vData As Long)
    mvarlCQtdPax = vData
End Property

Public Property Get lCQtdPax () AS Long
    lCQtdPax= mvarlCQtdPax
End Property

Public Property Let lPQtdPax (ByVal vData As Long)
    mvarlPQtdPax = vData
End Property

Public Property Get lPQtdPax () AS Long
    lPQtdPax= mvarlPQtdPax
End Property

Public Property Let dCValorFat (ByVal vData As Double)
    mvardCValorFat = vData
End Property

Public Property Get dCValorFat () AS Double
    dCValorFat= mvardCValorFat
End Property

Public Property Let dPValorFat (ByVal vData As Double)
    mvardPValorFat = vData
End Property

Public Property Get dPValorFat () AS Double
    dPValorFat= mvardPValorFat
End Property

Public Property Let dCValorLiq (ByVal vData As Double)
    mvardCValorLiq = vData
End Property

Public Property Get dCValorLiq () AS Double
    dCValorLiq= mvardCValorLiq
End Property

Public Property Let dPValorLiq (ByVal vData As Double)
    mvardPValorLiq = vData
End Property

Public Property Get dPValorLiq () AS Double
    dPValorLiq= mvardPValorLiq
End Property

Public Property Let dCValorBruto (ByVal vData As Double)
    mvardCValorBruto = vData
End Property

Public Property Get dCValorBruto () AS Double
    dCValorBruto= mvardCValorBruto
End Property

Public Property Let dPValorBruto (ByVal vData As Double)
    mvardPValorBruto = vData
End Property

Public Property Get dPValorBruto () AS Double
    dPValorBruto= mvardPValorBruto
End Property

Public Property Let dCValorInvestido (ByVal vData As Double)
    mvardCValorInvestido = vData
End Property

Public Property Get dCValorInvestido () AS Double
    dCValorInvestido= mvardCValorInvestido
End Property

Public Property Let dPValorInvestido (ByVal vData As Double)
    mvardPValorInvestido = vData
End Property

Public Property Get dPValorInvestido () AS Double
    dPValorInvestido= mvardPValorInvestido
End Property

Public Property Let lCCliNovos (ByVal vData As Long)
    mvarlCCliNovos = vData
End Property

Public Property Get lCCliNovos () AS Long
    lCCliNovos= mvarlCCliNovos
End Property

Public Property Let lPCliNovos (ByVal vData As Long)
    mvarlPCliNovos = vData
End Property

Public Property Get lPCliNovos () AS Long
    lPCliNovos= mvarlPCliNovos
End Property

Public Property Let lCCliReativ (ByVal vData As Long)
    mvarlCCliReativ = vData
End Property

Public Property Get lCCliReativ () AS Long
    lCCliReativ= mvarlCCliReativ
End Property

Public Property Let lPCliReativ (ByVal vData As Long)
    mvarlPCliReativ = vData
End Property

Public Property Get lPCliReativ () AS Long
    lPCliReativ= mvarlPCliReativ
End Property

Public Property Let dCPercDescMedio (ByVal vData As Double)
    mvardCPercDescMedio = vData
End Property

Public Property Get dCPercDescMedio () AS Double
    dCPercDescMedio= mvardCPercDescMedio
End Property

Public Property Let dPPercDescMedio (ByVal vData As Double)
    mvardPPercDescMedio = vData
End Property

Public Property Get dPPercDescMedio () AS Double
    dPPercDescMedio= mvardPPercDescMedio
End Property

Public Property Let sCMaiorProdCod (ByVal vData As String)
    mvarsCMaiorProdCod = vData
End Property

Public Property Get sCMaiorProdCod () AS String
    sCMaiorProdCod= mvarsCMaiorProdCod
End Property

Public Property Let sPMaiorProdCod (ByVal vData As String)
    mvarsPMaiorProdCod = vData
End Property

Public Property Get sPMaiorProdCod () AS String
    sPMaiorProdCod= mvarsPMaiorProdCod
End Property

Public Property Let sCMaiorProdDesc (ByVal vData As String)
    mvarsCMaiorProdDesc = vData
End Property

Public Property Get sCMaiorProdDesc () AS String
    sCMaiorProdDesc= mvarsCMaiorProdDesc
End Property

Public Property Let sPMaiorProdDesc (ByVal vData As String)
    mvarsPMaiorProdDesc = vData
End Property

Public Property Get sPMaiorProdDesc () AS String
    sPMaiorProdDesc= mvarsPMaiorProdDesc
End Property

Public Property Let lCMaiorProdQtd (ByVal vData As Long)
    mvarlCMaiorProdQtd = vData
End Property

Public Property Get lCMaiorProdQtd () AS Long
    lCMaiorProdQtd= mvarlCMaiorProdQtd
End Property

Public Property Let lPMaiorProdQtd (ByVal vData As Long)
    mvarlPMaiorProdQtd = vData
End Property

Public Property Get lPMaiorProdQtd () AS Long
    lPMaiorProdQtd= mvarlPMaiorProdQtd
End Property

Public Property Let dCMaiorProdVlrFat (ByVal vData As Double)
    mvardCMaiorProdVlrFat = vData
End Property

Public Property Get dCMaiorProdVlrFat () AS Double
    dCMaiorProdVlrFat= mvardCMaiorProdVlrFat
End Property

Public Property Let dPMaiorProdVlrFat (ByVal vData As Double)
    mvardPMaiorProdVlrFat = vData
End Property

Public Property Get dPMaiorProdVlrFat () AS Double
    dPMaiorProdVlrFat= mvardPMaiorProdVlrFat
End Property

Public Property Let dCMaiorProdVlrLiq (ByVal vData As Double)
    mvardCMaiorProdVlrLiq = vData
End Property

Public Property Get dCMaiorProdVlrLiq () AS Double
    dCMaiorProdVlrLiq= mvardCMaiorProdVlrLiq
End Property

Public Property Let dPMaiorProdVlrLiq (ByVal vData As Double)
    mvardPMaiorProdVlrLiq = vData
End Property

Public Property Get dPMaiorProdVlrLiq () AS Double
    dPMaiorProdVlrLiq= mvardPMaiorProdVlrLiq
End Property

Public Property Let dCMaiorProdVlrBruto (ByVal vData As Double)
    mvardCMaiorProdVlrBruto = vData
End Property

Public Property Get dCMaiorProdVlrBruto () AS Double
    dCMaiorProdVlrBruto= mvardCMaiorProdVlrBruto
End Property

Public Property Let dPMaiorProdVlrBruto (ByVal vData As Double)
    mvardPMaiorProdVlrBruto = vData
End Property

Public Property Get dPMaiorProdVlrBruto () AS Double
    dPMaiorProdVlrBruto= mvardPMaiorProdVlrBruto
End Property

Public Property Let lCContatosCall (ByVal vData As Long)
    mvarlCContatosCall = vData
End Property

Public Property Get lCContatosCall () AS Long
    lCContatosCall= mvarlCContatosCall
End Property

Public Property Let lPContatosCall (ByVal vData As Long)
    mvarlPContatosCall = vData
End Property

Public Property Get lPContatosCall () AS Long
    lPContatosCall= mvarlPContatosCall
End Property

Public Property Let lCContatosCobr (ByVal vData As Long)
    mvarlCContatosCobr = vData
End Property

Public Property Get lCContatosCobr () AS Long
    lCContatosCobr= mvarlCContatosCobr
End Property

Public Property Let lPContatosCobr (ByVal vData As Long)
    mvarlPContatosCobr = vData
End Property

Public Property Get lPContatosCobr () AS Long
    lPContatosCobr= mvarlPContatosCobr
End Property

Public Property Let lCContatosOutros (ByVal vData As Long)
    mvarlCContatosOutros = vData
End Property

Public Property Get lCContatosOutros () AS Long
    lCContatosOutros= mvarlCContatosOutros
End Property

Public Property Let lPContatosOutros (ByVal vData As Long)
    mvarlPContatosOutros = vData
End Property

Public Property Get lPContatosOutros () AS Long
    lPContatosOutros= mvarlPContatosOutros
End Property

Public Property Let lCMaiorCliCod (ByVal vData As Long)
    mvarlCMaiorCliCod = vData
End Property

Public Property Get lCMaiorCliCod () AS Long
    lCMaiorCliCod= mvarlCMaiorCliCod
End Property

Public Property Let lPMaiorCliCod (ByVal vData As Long)
    mvarlPMaiorCliCod = vData
End Property

Public Property Get lPMaiorCliCod () AS Long
    lPMaiorCliCod= mvarlPMaiorCliCod
End Property

Public Property Let sCMaiorCliNome (ByVal vData As String)
    mvarsCMaiorCliNome = vData
End Property

Public Property Get sCMaiorCliNome () AS String
    sCMaiorCliNome= mvarsCMaiorCliNome
End Property

Public Property Let sPMaiorCliNome (ByVal vData As String)
    mvarsPMaiorCliNome = vData
End Property

Public Property Get sPMaiorCliNome () AS String
    sPMaiorCliNome= mvarsPMaiorCliNome
End Property

Public Property Let dCMaiorCliValorFat (ByVal vData As Double)
    mvardCMaiorCliValorFat = vData
End Property

Public Property Get dCMaiorCliValorFat () AS Double
    dCMaiorCliValorFat= mvardCMaiorCliValorFat
End Property

Public Property Let dPMaiorCliValorFat (ByVal vData As Double)
    mvardPMaiorCliValorFat = vData
End Property

Public Property Get dPMaiorCliValorFat () AS Double
    dPMaiorCliValorFat= mvardPMaiorCliValorFat
End Property

Public Property Let dCMaiorCliValorLiq (ByVal vData As Double)
    mvardCMaiorCliValorLiq = vData
End Property

Public Property Get dCMaiorCliValorLiq () AS Double
    dCMaiorCliValorLiq= mvardCMaiorCliValorLiq
End Property

Public Property Let dPMaiorCliValorLiq (ByVal vData As Double)
    mvardPMaiorCliValorLiq = vData
End Property

Public Property Get dPMaiorCliValorLiq () AS Double
    dPMaiorCliValorLiq= mvardPMaiorCliValorLiq
End Property

Public Property Let dCMaiorCliValorBruto (ByVal vData As Double)
    mvardCMaiorCliValorBruto = vData
End Property

Public Property Get dCMaiorCliValorBruto () AS Double
    dCMaiorCliValorBruto= mvardCMaiorCliValorBruto
End Property

Public Property Let dPMaiorCliValorBruto (ByVal vData As Double)
    mvardPMaiorCliValorBruto = vData
End Property

Public Property Get dPMaiorCliValorBruto () AS Double
    dPMaiorCliValorBruto= mvardPMaiorCliValorBruto
End Property

Public Property Let dCPercVouCanc (ByVal vData As Double)
    mvardCPercVouCanc = vData
End Property

Public Property Get dCPercVouCanc () AS Double
    dCPercVouCanc= mvardCPercVouCanc
End Property

Public Property Let dPPercVouCanc (ByVal vData As Double)
    mvardPPercVouCanc = vData
End Property

Public Property Get dPPercVouCanc () AS Double
    dPPercVouCanc= mvardPPercVouCanc
End Property

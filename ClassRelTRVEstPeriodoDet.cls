VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassRelTRVEstPeriodoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel AS Long
Private mvarlSeq AS Long
Private mvariTipo AS Integer
Private mvarsTipoNome AS String
Private mvarsTextoLinha AS String
Private mvarlPQtdVou AS Long
Private mvardPPercVou AS Double
Private mvarlPQtdPax AS Long
Private mvardPPercPax AS Double
Private mvardPValorFat AS Double
Private mvardPPercValorFat AS Double
Private mvardPValorLiq AS Double
Private mvardPPercValorLiq AS Double
Private mvardPValorBruto AS Double
Private mvardPPercValorBruto AS Double
Private mvardPValorInvestido AS Double
Private mvardPPercDescMedio AS Double
Private mvarlCQtdVou AS Long
Private mvardCPercVou AS Double
Private mvarlCQtdPax AS Long
Private mvardCPercPax AS Double
Private mvardCValorFat AS Double
Private mvardCPercValorFat AS Double
Private mvardCValorLiq AS Double
Private mvardCPercValorLiq AS Double
Private mvardCValorBruto AS Double
Private mvardCPercValorBruto AS Double
Private mvardCValorInvestido AS Double
Private mvardCPercDescMedio AS Double

Public Property Let lNumIntRel (ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel () AS Long
    lNumIntRel= mvarlNumIntRel
End Property

Public Property Let lSeq (ByVal vData As Long)
    mvarlSeq = vData
End Property

Public Property Get lSeq () AS Long
    lSeq= mvarlSeq
End Property

Public Property Let iTipo (ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo () AS Integer
    iTipo= mvariTipo
End Property

Public Property Let sTipoNome (ByVal vData As String)
    mvarsTipoNome = vData
End Property

Public Property Get sTipoNome () AS String
    sTipoNome= mvarsTipoNome
End Property

Public Property Let sTextoLinha (ByVal vData As String)
    mvarsTextoLinha = vData
End Property

Public Property Get sTextoLinha () AS String
    sTextoLinha= mvarsTextoLinha
End Property

Public Property Let lPQtdVou (ByVal vData As Long)
    mvarlPQtdVou = vData
End Property

Public Property Get lPQtdVou () AS Long
    lPQtdVou= mvarlPQtdVou
End Property

Public Property Let dPPercVou (ByVal vData As Double)
    mvardPPercVou = vData
End Property

Public Property Get dPPercVou () AS Double
    dPPercVou= mvardPPercVou
End Property

Public Property Let lPQtdPax (ByVal vData As Long)
    mvarlPQtdPax = vData
End Property

Public Property Get lPQtdPax () AS Long
    lPQtdPax= mvarlPQtdPax
End Property

Public Property Let dPPercPax (ByVal vData As Double)
    mvardPPercPax = vData
End Property

Public Property Get dPPercPax () AS Double
    dPPercPax= mvardPPercPax
End Property

Public Property Let dPValorFat (ByVal vData As Double)
    mvardPValorFat = vData
End Property

Public Property Get dPValorFat () AS Double
    dPValorFat= mvardPValorFat
End Property

Public Property Let dPPercValorFat (ByVal vData As Double)
    mvardPPercValorFat = vData
End Property

Public Property Get dPPercValorFat () AS Double
    dPPercValorFat= mvardPPercValorFat
End Property

Public Property Let dPValorLiq (ByVal vData As Double)
    mvardPValorLiq = vData
End Property

Public Property Get dPValorLiq () AS Double
    dPValorLiq= mvardPValorLiq
End Property

Public Property Let dPPercValorLiq (ByVal vData As Double)
    mvardPPercValorLiq = vData
End Property

Public Property Get dPPercValorLiq () AS Double
    dPPercValorLiq= mvardPPercValorLiq
End Property

Public Property Let dPValorBruto (ByVal vData As Double)
    mvardPValorBruto = vData
End Property

Public Property Get dPValorBruto () AS Double
    dPValorBruto= mvardPValorBruto
End Property

Public Property Let dPPercValorBruto (ByVal vData As Double)
    mvardPPercValorBruto = vData
End Property

Public Property Get dPPercValorBruto () AS Double
    dPPercValorBruto= mvardPPercValorBruto
End Property

Public Property Let dPValorInvestido (ByVal vData As Double)
    mvardPValorInvestido = vData
End Property

Public Property Get dPValorInvestido () AS Double
    dPValorInvestido= mvardPValorInvestido
End Property

Public Property Let dPPercDescMedio (ByVal vData As Double)
    mvardPPercDescMedio = vData
End Property

Public Property Get dPPercDescMedio () AS Double
    dPPercDescMedio= mvardPPercDescMedio
End Property

Public Property Let lCQtdVou (ByVal vData As Long)
    mvarlCQtdVou = vData
End Property

Public Property Get lCQtdVou () AS Long
    lCQtdVou= mvarlCQtdVou
End Property

Public Property Let dCPercVou (ByVal vData As Double)
    mvardCPercVou = vData
End Property

Public Property Get dCPercVou () AS Double
    dCPercVou= mvardCPercVou
End Property

Public Property Let lCQtdPax (ByVal vData As Long)
    mvarlCQtdPax = vData
End Property

Public Property Get lCQtdPax () AS Long
    lCQtdPax= mvarlCQtdPax
End Property

Public Property Let dCPercPax (ByVal vData As Double)
    mvardCPercPax = vData
End Property

Public Property Get dCPercPax () AS Double
    dCPercPax= mvardCPercPax
End Property

Public Property Let dCValorFat (ByVal vData As Double)
    mvardCValorFat = vData
End Property

Public Property Get dCValorFat () AS Double
    dCValorFat= mvardCValorFat
End Property

Public Property Let dCPercValorFat (ByVal vData As Double)
    mvardCPercValorFat = vData
End Property

Public Property Get dCPercValorFat () AS Double
    dCPercValorFat= mvardCPercValorFat
End Property

Public Property Let dCValorLiq (ByVal vData As Double)
    mvardCValorLiq = vData
End Property

Public Property Get dCValorLiq () AS Double
    dCValorLiq= mvardCValorLiq
End Property

Public Property Let dCPercValorLiq (ByVal vData As Double)
    mvardCPercValorLiq = vData
End Property

Public Property Get dCPercValorLiq () AS Double
    dCPercValorLiq= mvardCPercValorLiq
End Property

Public Property Let dCValorBruto (ByVal vData As Double)
    mvardCValorBruto = vData
End Property

Public Property Get dCValorBruto () AS Double
    dCValorBruto= mvardCValorBruto
End Property

Public Property Let dCPercValorBruto (ByVal vData As Double)
    mvardCPercValorBruto = vData
End Property

Public Property Get dCPercValorBruto () AS Double
    dCPercValorBruto= mvardCPercValorBruto
End Property

Public Property Let dCValorInvestido (ByVal vData As Double)
    mvardCValorInvestido = vData
End Property

Public Property Get dCValorInvestido () AS Double
    dCValorInvestido= mvardCValorInvestido
End Property

Public Property Let dCPercDescMedio (ByVal vData As Double)
    mvardCPercDescMedio = vData
End Property

Public Property Get dCPercDescMedio () AS Double
    dCPercDescMedio= mvardCPercDescMedio
End Property

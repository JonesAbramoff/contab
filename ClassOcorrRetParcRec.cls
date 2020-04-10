VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOcorrRetParcRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntParc As Long
Private mvarlNumSeqOcorr As Long
Private mvarlNumIntBorderoRet As Long
Private mvariCodMovimento As Integer
Private mvariCodOcorrencia1 As Integer
Private mvariCodOcorrencia2 As Integer
Private mvariCodOcorrencia3 As Integer
Private mvariCodOcorrencia4 As Integer
Private mvariCodOcorrencia5 As Integer
Private mvariCodOcorrencia6 As Integer
Private mvariCodComplementar As Integer
Private mvardtDataComplementar As Date
Private mvardValorComplementar As Double
Private mvarsNumTitCobrador As String
Private mvarlNumSequencial As Long
Private mvarlNumIntRetCobrErros As Long

Public Property Let iCodMovimento(ByVal vData As Integer)
    mvariCodMovimento = vData
End Property

Public Property Get iCodMovimento() As Integer
    iCodMovimento = mvariCodMovimento
End Property

Public Property Let lNumSequencial(ByVal vData As Long)
    mvarlNumSequencial = vData
End Property

Public Property Get lNumSequencial() As Long
    lNumSequencial = mvarlNumSequencial
End Property

Public Property Let lNumIntRetCobrErros(ByVal vData As Long)
    mvarlNumIntRetCobrErros = vData
End Property

Public Property Get lNumIntRetCobrErros() As Long
    lNumIntRetCobrErros = mvarlNumIntRetCobrErros
End Property

Public Property Let sNumTitCobrador(ByVal vData As String)
    mvarsNumTitCobrador = vData
End Property

Public Property Get sNumTitCobrador() As String
    sNumTitCobrador = mvarsNumTitCobrador
End Property

Public Property Let dValorComplementar(ByVal vData As Double)
    mvardValorComplementar = vData
End Property

Public Property Get dValorComplementar() As Double
    dValorComplementar = mvardValorComplementar
End Property

Public Property Let dtDataComplementar(ByVal vData As Date)
    mvardtDataComplementar = vData
End Property

Public Property Get dtDataComplementar() As Date
    dtDataComplementar = mvardtDataComplementar
End Property

Public Property Let iCodComplementar(ByVal vData As Integer)
    mvariCodComplementar = vData
End Property

Public Property Get iCodComplementar() As Integer
    iCodComplementar = mvariCodComplementar
End Property

Public Property Let iCodOcorrencia6(ByVal vData As Integer)
    mvariCodOcorrencia6 = vData
End Property

Public Property Get iCodOcorrencia6() As Integer
    iCodOcorrencia6 = mvariCodOcorrencia6
End Property

Public Property Let iCodOcorrencia5(ByVal vData As Integer)
    mvariCodOcorrencia5 = vData
End Property

Public Property Get iCodOcorrencia5() As Integer
    iCodOcorrencia5 = mvariCodOcorrencia5
End Property

Public Property Let iCodOcorrencia4(ByVal vData As Integer)
    mvariCodOcorrencia4 = vData
End Property

Public Property Get iCodOcorrencia4() As Integer
    iCodOcorrencia4 = mvariCodOcorrencia4
End Property
Public Property Let iCodOcorrencia3(ByVal vData As Integer)
    mvariCodOcorrencia3 = vData
End Property

Public Property Get iCodOcorrencia3() As Integer
    iCodOcorrencia3 = mvariCodOcorrencia3
End Property

Public Property Let iCodOcorrencia2(ByVal vData As Integer)
    mvariCodOcorrencia2 = vData
End Property

Public Property Get iCodOcorrencia2() As Integer
    iCodOcorrencia2 = mvariCodOcorrencia2
End Property

Public Property Let iCodOcorrencia1(ByVal vData As Integer)
    mvariCodOcorrencia1 = vData
End Property

Public Property Get iCodOcorrencia1() As Integer
    iCodOcorrencia1 = mvariCodOcorrencia1
End Property

Public Property Let lNumIntBorderoRet(ByVal vData As Long)
    mvarlNumIntBorderoRet = vData
End Property

Public Property Get lNumIntBorderoRet() As Long
    lNumIntBorderoRet = mvarlNumIntBorderoRet
End Property

Public Property Let lNumSeqOcorr(ByVal vData As Long)
    mvarlNumSeqOcorr = vData
End Property

Public Property Get lNumSeqOcorr() As Long
    lNumSeqOcorr = mvarlNumSeqOcorr
End Property

Public Property Let lNumIntParc(ByVal vData As Long)
    mvarlNumIntParc = vData
End Property

Public Property Get lNumIntParc() As Long
    lNumIntParc = mvarlNumIntParc
End Property

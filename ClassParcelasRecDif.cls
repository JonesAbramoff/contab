VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassParcelasRecDif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc AS Long
Private mvarlNumIntParc AS Long
Private mvariSeq AS Integer
Private mvardtDataRegistro AS Date
Private mvariCodTipoDif AS Integer
Private mvardValorDiferenca AS Double
Private mvarsObservacao AS String
Private mvariNumSeqOcorr AS Integer

Public Property Let lNumIntDoc (ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc () AS Long
    lNumIntDoc= mvarlNumIntDoc
End Property

Public Property Let lNumIntParc (ByVal vData As Long)
    mvarlNumIntParc = vData
End Property

Public Property Get lNumIntParc () AS Long
    lNumIntParc= mvarlNumIntParc
End Property

Public Property Let iSeq (ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq () AS Integer
    iSeq= mvariSeq
End Property

Public Property Let dtDataRegistro (ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro () AS Date
    dtDataRegistro= mvardtDataRegistro
End Property

Public Property Let iCodTipoDif (ByVal vData As Integer)
    mvariCodTipoDif = vData
End Property

Public Property Get iCodTipoDif () AS Integer
    iCodTipoDif= mvariCodTipoDif
End Property

Public Property Let dValorDiferenca (ByVal vData As Double)
    mvardValorDiferenca = vData
End Property

Public Property Get dValorDiferenca () AS Double
    dValorDiferenca= mvardValorDiferenca
End Property

Public Property Let sObservacao (ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao () AS String
    sObservacao= mvarsObservacao
End Property

Public Property Let iNumSeqOcorr (ByVal vData As Integer)
    mvariNumSeqOcorr = vData
End Property

Public Property Get iNumSeqOcorr () AS Integer
    iNumSeqOcorr= mvariNumSeqOcorr
End Property

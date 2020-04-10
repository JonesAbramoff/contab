VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTranfCartCobr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntParc As Long
Private mvariNumSeqOcorr As Integer
Private mvariCobrador As Integer
Private mvariCarteiraCobranca As Integer
Private mvardtData As Date
Private mvardtDataRegistro As Date

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iCarteiraCobranca(ByVal vData As Integer)
    mvariCarteiraCobranca = vData
End Property

Public Property Get iCarteiraCobranca() As Integer
    iCarteiraCobranca = mvariCarteiraCobranca
End Property

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property

Public Property Let iNumSeqOcorr(ByVal vData As Integer)
    mvariNumSeqOcorr = vData
End Property

Public Property Get iNumSeqOcorr() As Integer
    iNumSeqOcorr = mvariNumSeqOcorr
End Property

Public Property Let lNumIntParc(ByVal vData As Long)
    mvarlNumIntParc = vData
End Property

Public Property Get lNumIntParc() As Long
    lNumIntParc = mvarlNumIntParc
End Property

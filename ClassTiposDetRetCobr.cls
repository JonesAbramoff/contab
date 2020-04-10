VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassTiposDetRetCobr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlBanco AS Long
Private mvariCodigoMovto AS Integer
Private mvariCodigoDetalhe AS Integer
Private mvarsDescricao AS String
Private mvariAcao AS Integer
Private mvariAcaoManual AS Integer
Private mvariCodTipoDiferenca AS Integer

Public Property Let lBanco (ByVal vData As Long)
    mvarlBanco = vData
End Property

Public Property Get lBanco () AS Long
    lBanco= mvarlBanco
End Property

Public Property Let iCodigoMovto (ByVal vData As Integer)
    mvariCodigoMovto = vData
End Property

Public Property Get iCodigoMovto () AS Integer
    iCodigoMovto= mvariCodigoMovto
End Property

Public Property Let iCodigoDetalhe (ByVal vData As Integer)
    mvariCodigoDetalhe = vData
End Property

Public Property Get iCodigoDetalhe () AS Integer
    iCodigoDetalhe= mvariCodigoDetalhe
End Property

Public Property Let sDescricao (ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao () AS String
    sDescricao= mvarsDescricao
End Property

Public Property Let iAcao (ByVal vData As Integer)
    mvariAcao = vData
End Property

Public Property Get iAcao () AS Integer
    iAcao= mvariAcao
End Property

Public Property Let iAcaoManual (ByVal vData As Integer)
    mvariAcaoManual = vData
End Property

Public Property Get iAcaoManual () AS Integer
    iAcaoManual= mvariAcaoManual
End Property

Public Property Let iCodTipoDiferenca (ByVal vData As Integer)
    mvariCodTipoDiferenca = vData
End Property

Public Property Get iCodTipoDiferenca () AS Integer
    iCodTipoDiferenca= mvariCodTipoDiferenca
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassTestesQualidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariCodigo AS Integer
Private mvarsNomeReduzido AS String
Private mvarsEspecificacao AS String
Private mvariTipoResultado AS Integer
Private mvardLimiteDe AS Double
Private mvardLimiteAte AS Double
Private mvarsMetodoUsado AS String
Private mvarsObservacao AS String
Private mvariNoCertificado AS Integer

Public Property Let iCodigo (ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo () AS Integer
    iCodigo= mvariCodigo
End Property

Public Property Let sNomeReduzido (ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido () AS String
    sNomeReduzido= mvarsNomeReduzido
End Property

Public Property Let sEspecificacao (ByVal vData As String)
    mvarsEspecificacao = vData
End Property

Public Property Get sEspecificacao () AS String
    sEspecificacao= mvarsEspecificacao
End Property

Public Property Let iTipoResultado (ByVal vData As Integer)
    mvariTipoResultado = vData
End Property

Public Property Get iTipoResultado () AS Integer
    iTipoResultado= mvariTipoResultado
End Property

Public Property Let dLimiteDe (ByVal vData As Double)
    mvardLimiteDe = vData
End Property

Public Property Get dLimiteDe () AS Double
    dLimiteDe= mvardLimiteDe
End Property

Public Property Let dLimiteAte (ByVal vData As Double)
    mvardLimiteAte = vData
End Property

Public Property Get dLimiteAte () AS Double
    dLimiteAte= mvardLimiteAte
End Property

Public Property Let sMetodoUsado (ByVal vData As String)
    mvarsMetodoUsado = vData
End Property

Public Property Get sMetodoUsado () AS String
    sMetodoUsado= mvarsMetodoUsado
End Property

Public Property Let sObservacao (ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao () AS String
    sObservacao= mvarsObservacao
End Property

Public Property Let iNoCertificado (ByVal vData As Integer)
    mvariNoCertificado = vData
End Property

Public Property Get iNoCertificado () AS Integer
    iNoCertificado= mvariNoCertificado
End Property

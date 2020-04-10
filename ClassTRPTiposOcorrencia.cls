VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPTiposOcorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvariEstornaAporteVou As Integer
Private mvariConsideraComisInt As Integer
Private mvariAlteraComiVou As Integer
Private mvariAlteraCMCC As Integer
Private mvariAlteraCMC As Integer
Private mvariAlteraCMR As Integer
Private mvariAlteraOVER As Integer
Private mvariAlteraCMA As Integer
Private mvariAceitaVlrPositivo As Integer
Private mvariAceitaVlrNegativo As Integer
Private mvariIncideSobre As Integer

Public Property Let iIncideSobre(ByVal vData As Integer)
    mvariIncideSobre = vData
End Property

Public Property Get iIncideSobre() As Integer
    iIncideSobre = mvariIncideSobre
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iEstornaAporteVou(ByVal vData As Integer)
    mvariEstornaAporteVou = vData
End Property

Public Property Get iEstornaAporteVou() As Integer
    iEstornaAporteVou = mvariEstornaAporteVou
End Property

Public Property Let iConsideraComisInt(ByVal vData As Integer)
    mvariConsideraComisInt = vData
End Property

Public Property Get iConsideraComisInt() As Integer
    iConsideraComisInt = mvariConsideraComisInt
End Property

Public Property Let iAlteraComiVou(ByVal vData As Integer)
    mvariAlteraComiVou = vData
End Property

Public Property Get iAlteraComiVou() As Integer
    iAlteraComiVou = mvariAlteraComiVou
End Property

Public Property Let iAlteraCMCC(ByVal vData As Integer)
    mvariAlteraCMCC = vData
End Property

Public Property Get iAlteraCMCC() As Integer
    iAlteraCMCC = mvariAlteraCMCC
End Property

Public Property Let iAlteraCMC(ByVal vData As Integer)
    mvariAlteraCMC = vData
End Property

Public Property Get iAlteraCMC() As Integer
    iAlteraCMC = mvariAlteraCMC
End Property

Public Property Let iAlteraCMR(ByVal vData As Integer)
    mvariAlteraCMR = vData
End Property

Public Property Get iAlteraCMR() As Integer
    iAlteraCMR = mvariAlteraCMR
End Property

Public Property Let iAlteraOVER(ByVal vData As Integer)
    mvariAlteraOVER = vData
End Property

Public Property Get iAlteraOVER() As Integer
    iAlteraOVER = mvariAlteraOVER
End Property

Public Property Let iAlteraCMA(ByVal vData As Integer)
    mvariAlteraCMA = vData
End Property

Public Property Get iAlteraCMA() As Integer
    iAlteraCMA = mvariAlteraCMA
End Property

Public Property Let iAceitaVlrPositivo(ByVal vData As Integer)
    mvariAceitaVlrPositivo = vData
End Property

Public Property Get iAceitaVlrPositivo() As Integer
    iAceitaVlrPositivo = mvariAceitaVlrPositivo
End Property

Public Property Let iAceitaVlrNegativo(ByVal vData As Integer)
    mvariAceitaVlrNegativo = vData
End Property

Public Property Get iAceitaVlrNegativo() As Integer
    iAceitaVlrNegativo = mvariAceitaVlrNegativo
End Property

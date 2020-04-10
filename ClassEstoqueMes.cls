VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEstoqueMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariAno As Integer
Private mvariMes As Integer
Private mvariFechamento As Integer
Private mvariAlterado As Integer
Private mvariCustoProdApurado As Integer
Private mvardCustoMatPrim As Double
Private mvarlHorasMaquina As Long
Private mvardGastosIndiretos As Double
Private mvardGastosDiretos As Double
Private mvarsNomeFilialEmpresa As String

Public dQuantFator1 As Double
Public dCustoFator1 As Double
Public dQuantFator2 As Double
Public dCustoFator2 As Double
Public dQuantFator3 As Double
Public dCustoFator3 As Double
Public dQuantFator4 As Double
Public dCustoFator4 As Double
Public dQuantFator5 As Double
Public dCustoFator5 As Double
Public dQuantFator6 As Double
Public dCustoFator6 As Double

Public Property Let sNomeFilialEmpresa(ByVal vData As String)
    mvarsNomeFilialEmpresa = vData
End Property

Public Property Get sNomeFilialEmpresa() As String
    sNomeFilialEmpresa = mvarsNomeFilialEmpresa
End Property

Public Property Let dGastosDiretos(ByVal vData As Double)
    mvardGastosDiretos = vData
End Property

Public Property Get dGastosDiretos() As Double
    dGastosDiretos = mvardGastosDiretos
End Property

Public Property Let dGastosIndiretos(ByVal vData As Double)
    mvardGastosIndiretos = vData
End Property

Public Property Get dGastosIndiretos() As Double
    dGastosIndiretos = mvardGastosIndiretos
End Property

Public Property Let lHorasMaquina(ByVal vData As Long)
    mvarlHorasMaquina = vData
End Property

Public Property Get lHorasMaquina() As Long
    lHorasMaquina = mvarlHorasMaquina
End Property

Public Property Let dCustoMatPrim(ByVal vData As Double)
    mvardCustoMatPrim = vData
End Property

Public Property Get dCustoMatPrim() As Double
    dCustoMatPrim = mvardCustoMatPrim
End Property

Public Property Let iCustoProdApurado(ByVal vData As Integer)
Attribute iCustoProdApurado.VB_Description = "0=nao apurado 1=apurado"
    mvariCustoProdApurado = vData
End Property

Public Property Get iCustoProdApurado() As Integer
    iCustoProdApurado = mvariCustoProdApurado
End Property

Public Property Let iAlterado(ByVal vData As Integer)
Attribute iAlterado.VB_Description = "0=nao alterado 1=alterado (necessita de recalculo antes de fechar)"
    mvariAlterado = vData
End Property

Public Property Get iAlterado() As Integer
    iAlterado = mvariAlterado
End Property

Public Property Let iFechamento(ByVal vData As Integer)
Attribute iFechamento.VB_Description = "0=aberto  1=fechado (nao admite movimentações)"
    mvariFechamento = vData
End Property

Public Property Get iFechamento() As Integer
    iFechamento = mvariFechamento
End Property

Public Property Let iMes(ByVal vData As Integer)
Attribute iMes.VB_Description = "Mes (1 a 12)"
    mvariMes = vData
End Property

Public Property Get iMes() As Integer
    iMes = mvariMes
End Property

Public Property Let iAno(ByVal vData As Integer)
Attribute iAno.VB_Description = "Ano"
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
Attribute iFilialEmpresa.VB_Description = "identifica a filial ou estabelecimento da empresa"
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPProdTarifa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsProduto As String
Private mvariDiasDe As Integer
Private mvariDiasAte As Integer
Private mvariDiario As Integer
Private mvardValor As Double
Private mvardValorAdicional As Double
Private mvardtVigencia As Date

Public Property Let dtVigencia(ByVal vData As Date)
    mvardtVigencia = vData
End Property

Public Property Get dtVigencia() As Date
    dtVigencia = mvardtVigencia
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iDiasDe(ByVal vData As Integer)
    mvariDiasDe = vData
End Property

Public Property Get iDiasDe() As Integer
    iDiasDe = mvariDiasDe
End Property

Public Property Let iDiasAte(ByVal vData As Integer)
    mvariDiasAte = vData
End Property

Public Property Get iDiasAte() As Integer
    iDiasAte = mvariDiasAte
End Property

Public Property Let iDiario(ByVal vData As Integer)
    mvariDiario = vData
End Property

Public Property Get iDiario() As Integer
    iDiario = mvariDiario
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dValorAdicional(ByVal vData As Double)
    mvardValorAdicional = vData
End Property

Public Property Get dValorAdicional() As Double
    dValorAdicional = mvardValorAdicional
End Property

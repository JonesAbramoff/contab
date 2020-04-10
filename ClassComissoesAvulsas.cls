VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComissoesAvulsas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mvarlNumIntDoc As Long
Private mvariVendedor As Integer
Private mvardtData As Date
Private mvarsReferencia As String
Private mvariCodigoMotivo As Integer
Private mvardBaseCalculo As Double
Private mvardAliquota As Double
Private mvardValorComissao As Double

Public Property Let dValorComissao(ByVal vData As Double)
    mvardValorComissao = vData
End Property

Public Property Get dValorComissao() As Double
    dValorComissao = mvardValorComissao
End Property

Public Property Let dAliquota(ByVal vData As Double)
    mvardAliquota = vData
End Property

Public Property Get dAliquota() As Double
    dAliquota = mvardAliquota
End Property

Public Property Let dBaseCalculo(ByVal vData As Double)
    mvardBaseCalculo = vData
End Property

Public Property Get dBaseCalculo() As Double
    dBaseCalculo = mvardBaseCalculo
End Property

Public Property Let iCodigoMotivo(ByVal vData As Integer)
    mvariCodigoMotivo = vData
End Property

Public Property Get iCodigoMotivo() As Integer
    iCodigoMotivo = mvariCodigoMotivo
End Property

Public Property Let sReferencia(ByVal vData As String)
    mvarsReferencia = vData
End Property

Public Property Get sReferencia() As String
    sReferencia = mvarsReferencia
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property





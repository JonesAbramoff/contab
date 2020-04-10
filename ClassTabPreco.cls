VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTabPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCodigo As Long
Private mvarlCliente As Long
Private mvariOrigem As Integer
Private mvariDestino As Integer
Private mvardPedagio As Double
Private mvardAdValoren As Double
Private mvarsObservacao As String
Private mvarcolTabPrecoItens As New Collection
Private mvardtDataVigencia As Date
Private mvarsClienteNomeRed As String
Private mvarsOrigemCidade As String
Private mvarsOrigemUF As String
Private mvarsDestinoCidade As String
Private mvarsDestinoUF As String

Public Property Let sDestinoUF(ByVal vData As String)
    mvarsDestinoUF = vData
End Property

Public Property Get sDestinoUF() As String
    sDestinoUF = mvarsDestinoUF
End Property

Public Property Let sDestinoCidade(ByVal vData As String)
    mvarsDestinoCidade = vData
End Property

Public Property Get sDestinoCidade() As String
    sDestinoCidade = mvarsDestinoCidade
End Property

Public Property Let sOrigemUF(ByVal vData As String)
    mvarsOrigemUF = vData
End Property

Public Property Get sOrigemUF() As String
    sOrigemUF = mvarsOrigemUF
End Property

Public Property Let sOrigemCidade(ByVal vData As String)
    mvarsOrigemCidade = vData
End Property

Public Property Get sOrigemCidade() As String
    sOrigemCidade = mvarsOrigemCidade
End Property

Public Property Let sClienteNomeRed(ByVal vData As String)
    mvarsClienteNomeRed = vData
End Property

Public Property Get sClienteNomeRed() As String
    sClienteNomeRed = mvarsClienteNomeRed
End Property

Public Property Let dtDataVigencia(ByVal vData As Date)
    mvardtDataVigencia = vData
End Property

Public Property Get dtDataVigencia() As Date
    dtDataVigencia = mvardtDataVigencia
End Property

Public Property Set colTabPrecoItens(ByVal vData As Object)
    Set mvarcolTabPrecoItens = vData
End Property

Public Property Get colTabPrecoItens() As Collection
    Set colTabPrecoItens = mvarcolTabPrecoItens
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dAdValoren(ByVal vData As Double)
    mvardAdValoren = vData
End Property

Public Property Get dAdValoren() As Double
    dAdValoren = mvardAdValoren
End Property

Public Property Let dPedagio(ByVal vData As Double)
    mvardPedagio = vData
End Property

Public Property Get dPedagio() As Double
    dPedagio = mvardPedagio
End Property

Public Property Let iDestino(ByVal vData As Integer)
    mvariDestino = vData
End Property

Public Property Get iDestino() As Integer
    iDestino = mvariDestino
End Property

Public Property Let iOrigem(ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property



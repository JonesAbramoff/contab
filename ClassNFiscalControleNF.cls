VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFiscalControleNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumForm As Long
Private mvarlNumNFiscal As Long
Private mvardtDataRegistro As Date
Private mvardHoraRegistro As Double
Private mvarsObservacao As String
Private mvariFilialEmpresa As Integer
Private mvarsUsuario As String
Private mvarsSerie As String
Private mvardtDataEmissao As Date
Private mvarlNumPedidoVenda As Long
Private mvardValorTotal As Double
Private mvarsNomeReduzidoCli As String
Private mvarsNomeReduzidoForn As String
Private mvariTipo As Integer
Private mvardICMSValor As Double
Private mvardIPIValor As Double

Public Property Let dIPIValor(ByVal vData As Double)
    mvardIPIValor = vData
End Property

Public Property Get dIPIValor() As Double
    dIPIValor = mvardIPIValor
End Property

Public Property Let dICMSValor(ByVal vData As Double)
    mvardICMSValor = vData
End Property

Public Property Get dICMSValor() As Double
    dICMSValor = mvardICMSValor
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sNomeReduzidoForn(ByVal vData As String)
    mvarsNomeReduzidoForn = vData
End Property

Public Property Get sNomeReduzidoForn() As String
    sNomeReduzidoForn = mvarsNomeReduzidoForn
End Property

Public Property Let sNomeReduzidoCli(ByVal vData As String)
    mvarsNomeReduzidoCli = vData
End Property

Public Property Get sNomeReduzidoCli() As String
    sNomeReduzidoCli = mvarsNomeReduzidoCli
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let lNumPedidoVenda(ByVal vData As Long)
    mvarlNumPedidoVenda = vData
End Property

Public Property Get lNumPedidoVenda() As Long
    lNumPedidoVenda = mvarlNumPedidoVenda
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dHoraRegistro(ByVal vData As Double)
    mvardHoraRegistro = vData
End Property

Public Property Get dHoraRegistro() As Double
    dHoraRegistro = mvardHoraRegistro
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let lNumNFiscal(ByVal vData As Long)
    mvarlNumNFiscal = vData
End Property

Public Property Get lNumNFiscal() As Long
    lNumNFiscal = mvarlNumNFiscal
End Property

Public Property Let lNumForm(ByVal vData As Long)
    mvarlNumForm = vData
End Property

Public Property Get lNumForm() As Long
    lNumForm = mvarlNumForm
End Property



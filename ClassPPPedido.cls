VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPPPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'classe para importacao de pedido do aplicativo de pesquisa de preços
Option Explicit

Public IdCarrinho As Long

Public LojaCNPJ As String

Public Status As Integer

Public Email As String
Public Nome As String

Public DataRegistro As Date
Public HoraRegistro As Double

Public ValorTotal As Double

Public TelConfirmacao As String
Public Observacao As String
Public EnderecoLogradouro As String
Public EnderecoNumero As String
Public EnderecoComplemento As String
Public EnderecoBairro As String
Public EnderecoCidade As String
Public EnderecoUF As String
Public EnderecoCEP As String
Public EnderecoFone As String
Public EnderecoReferencia As String

Public TaxaEntrega As Double

Public FormaPagto As Integer
Public MeioPagto As Integer
Public LocalPagto As Integer
Public NumParcelas As Integer
Public NumeroCartao As String
Public Tid As String
Public AutorizacaoARP As String
Public NumeroBanco As String
Public NumeroAgencia As String
Public NumeroCheque As String
Public TrocoPara As Double
Public StatusPagto As Integer
Public FormaEntrega As Integer

Public colItens As New Collection


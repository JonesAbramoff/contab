VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTransacaoCTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsSigla As String
Private mvarsTransacao As String
Private mvariCodigo As Integer
Private mvarsTransacaoTela As String
Private mvariCliFornecedor As Integer
Private mvariSubTipo As Integer
Private mvariObrigatoriedade As Integer
Private mvariOrigemLcto As Integer
Private mvariGerencial As Integer

Public Property Let iGerencial(ByVal iData As Integer)
    mvariGerencial = iData
End Property

Public Property Get iGerencial() As Integer
    iGerencial = mvariGerencial
End Property

Public Property Let iOrigemLcto(ByVal iData As Integer)
    mvariOrigemLcto = iData
End Property

Public Property Get iOrigemLcto() As Integer
    iOrigemLcto = mvariOrigemLcto
End Property

Public Property Let iObrigatoriedade(ByVal iData As Integer)
    mvariObrigatoriedade = iData
End Property

Public Property Get iObrigatoriedade() As Integer
    iObrigatoriedade = mvariObrigatoriedade
End Property


Public Property Let iSubTipo(ByVal iData As Integer)
    mvariSubTipo = iData
End Property

Public Property Get iSubTipo() As Integer
    iSubTipo = mvariSubTipo
End Property

Public Property Let iCliFornecedor(ByVal iData As Integer)
    mvariCliFornecedor = iData
End Property

Public Property Get iCliFornecedor() As Integer
    iCliFornecedor = mvariCliFornecedor
End Property

Public Property Let sTransacaoTela(ByVal vData As String)
    mvarsTransacaoTela = vData
End Property

Public Property Get sTransacaoTela() As String
    sTransacaoTela = mvarsTransacaoTela
End Property

Public Property Let iCodigo(ByVal iData As Integer)
    mvariCodigo = iData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sTransacao(ByVal vData As String)
Attribute sTransacao.VB_Description = "Nome da Transacao"
    mvarsTransacao = vData
End Property

Public Property Get sTransacao() As String
    sTransacao = mvarsTransacao
End Property

Public Property Let sSigla(ByVal vData As String)
Attribute sSigla.VB_Description = "Sigla que representa o modulo. Ex: CTB/CPR/FAT..."
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property





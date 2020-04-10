VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRateioOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "se preenchido, contém a chave do registro de um rateio off-line."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvarsDescricao As String
Private mvarsCclOrigem As String
Private mvarsContaCre As String
Private mvariTipo As Integer
Private mvarsConta As String
Private mvarsCcl As String
Private mvardPercentual As Double
Private mvarlSeq As Long
Private mvardValor As Double

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let lSeq(ByVal vData As Long)
Attribute lSeq.VB_Description = "Numero Sequencial."
    mvarlSeq = vData
End Property

Public Property Get lSeq() As Long
    lSeq = mvarlSeq
End Property

Public Property Let dPercentual(ByVal vData As Double)
    mvardPercentual = vData
End Property

Public Property Get dPercentual() As Double
    dPercentual = mvardPercentual
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let sConta(ByVal vData As String)
Attribute sConta.VB_Description = "Número da conta que vai receber parte do valor rateado"
    mvarsConta = vData
End Property

Public Property Get sConta() As String
    sConta = mvarsConta
End Property

Public Property Let iTipo(ByVal vData As Integer)
Attribute iTipo.VB_Description = "Tipo do Rateio. M=Mensal (utiliza as movimentações do mes) S= saldo até aquele mes I=indireto. Usa como origem o centro de custo."
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sContaCre(ByVal vData As String)
Attribute sContaCre.VB_Description = "Conta que vai ser creditada do valor a ser rateado."
    mvarsContaCre = vData
End Property

Public Property Get sContaCre() As String
    sContaCre = mvarsContaCre
End Property

Public Property Let sCclOrigem(ByVal vData As String)
Attribute sCclOrigem.VB_Description = "Centro de Custo/Lucro origem do rateio. Só é válido se o tipo for = 1"
    mvarsCclOrigem = vData
End Property

Public Property Get sCclOrigem() As String
    sCclOrigem = mvarsCclOrigem
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let lCodigo(ByVal vData As Long)
Attribute lCodigo.VB_Description = "Conta Origem do Rateio. Seu conteudo só é valido se o tipo do rateio for 2 ou 3."
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Sub Limpa()
    
    mvarsCclOrigem = ""
    mvarsContaCre = ""
    mvarsConta = ""
End Sub


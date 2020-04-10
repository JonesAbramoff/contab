VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDocAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlDoc As Long
Private mvariSeq As Integer
Private mvarsDescricao As String
Private mvarsConta As String
Private mvarsCcl As String
Private mvarsHistorico As String
Private mvardValor As Double
Private mvarsOrigem As String
Private mvariSeqContraPartida As Integer

Public Property Let iSeqContraPartida(ByVal vData As Integer)
    mvariSeqContraPartida = vData
End Property

Public Property Get iSeqContraPartida() As Integer
    iSeqContraPartida = mvariSeqContraPartida
End Property

Public Property Let sOrigem(ByVal vData As String)
Attribute sOrigem.VB_Description = "CPR=Contas a Pagar/Receber, EST= Estoque, CTB=contabilidade etc."
    mvarsOrigem = vData
End Property

Public Property Get sOrigem() As String
    sOrigem = mvarsOrigem
End Property

Public Property Let dValor(ByVal vData As Double)
Attribute dValor.VB_Description = "valor do lançamento ou nulo (Se > 0 é crédito, senao é débito)"
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let sHistorico(ByVal vData As String)
Attribute sHistorico.VB_Description = "Texto do histórico ou nulo (Tamanho máximo 150  bytes)"
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let sCcl(ByVal vData As String)
Attribute sCcl.VB_Description = "Número do ccusto extra contabil ou nulo"
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let sConta(ByVal vData As String)
Attribute sConta.VB_Description = "Pode conter o número da conta ou nulo"
    mvarsConta = vData
End Property

Public Property Get sConta() As String
    sConta = mvarsConta
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descrição do documento automático"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iSeq(ByVal vData As Integer)
Attribute iSeq.VB_Description = "Numero Sequencial do lançamento no documento (pode variar de 1 a 9999)"
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let lDoc(ByVal vData As Long)
Attribute lDoc.VB_Description = "Código que identifica o documento automático"
    mvarlDoc = vData
End Property

Public Property Get lDoc() As Long
    lDoc = mvarlDoc
End Property




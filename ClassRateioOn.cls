VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassRateioOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvariSeq As Integer
Private mvarsConta As String
Private mvarsCcl As String
Private mvardPercentual As Double
Private mvarsHistorico As String
Private mvarsDescricao As String

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descrição do Rateio."
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sHistorico(ByVal vData As String)
Attribute sHistorico.VB_Description = "Texto do histórico (Tamanho máximo 150 bytes)"
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let dPercentual(ByVal vData As Double)
Attribute dPercentual.VB_Description = "Percentual é um número inteiro positivo (1 a 99)"
    mvardPercentual = vData
End Property

Public Property Get dPercentual() As Double
    dPercentual = mvardPercentual
End Property

Public Property Let sCcl(ByVal vData As String)
Attribute sCcl.VB_Description = "Número do ccusto extra contabil"
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let sConta(ByVal vData As String)
Attribute sConta.VB_Description = "Número da conta"
    mvarsConta = vData
End Property

Public Property Get sConta() As String
    sConta = mvarsConta
End Property

Public Property Let iSeq(ByVal vData As Integer)
Attribute iSeq.VB_Description = "Numero Sequencial."
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iCodigo(ByVal vData As Integer)
Attribute iCodigo.VB_Description = "Identificador do Rateio"
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property




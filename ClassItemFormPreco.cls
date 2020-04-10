VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemFormPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'##### ATENÇÂO: EXISTE FUNÇÃO 'COPIAR': SE INCLUIR ALGUM CAMPO É PRECISO ATUALIZA-LA #####

Private mvarlNumIntDoc As Long 'na tabela de ItensFormPreco
Private mvariTipoDocOrigem As Integer '0:item de orcamento de venda
Private mvarlNumIntDocOrigem As Long 'correspondente ao tipodocorigem
Private mvariSequencial As Integer 'para que possa ser listado sempre numa mesma sequencia
Private mvarsProduto As String  'é o insumo/componente/proprio item
Private mvarsUnidMed As String
Private mvardQtde As Double 'deste produto que será usada no produto para o qual está sendo calculado o preço
Private mvardCustoUnit As Double 'obtido da cotação
Private mvardPercentMargem As Double 'margem a ser aplicada sobre o custo para obter o preço
Private mvardPrecoUnit As Double 'calculado: custo unit * (1+margem)
Private mvardPrecoTotal As Double 'calculado: qtde * precounit
Private mvariSituacao As Integer '1:cotando, 2:cotado

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Get iTipoDocOrigem() As Integer
    iTipoDocOrigem = mvariTipoDocOrigem
End Property

Public Property Let iTipoDocOrigem(ByVal vData As Integer)
    mvariTipoDocOrigem = vData
End Property

Public Property Get iSequencial() As Integer
    iSequencial = mvariSequencial
End Property

Public Property Let iSequencial(ByVal vData As Integer)
    mvariSequencial = vData
End Property

Public Property Get iSituacao() As Integer
    iSituacao = mvariSituacao
End Property

Public Property Let iSituacao(ByVal vData As Integer)
    mvariSituacao = vData
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sUnidMed(ByVal vData As String)
    mvarsUnidMed = vData
End Property

Public Property Get sUnidMed() As String
    sUnidMed = mvarsUnidMed
End Property

Public Property Let dQtde(ByVal vData As Double)
    mvardQtde = vData
End Property

Public Property Get dQtde() As Double
    dQtde = mvardQtde
End Property

Public Property Let dCustoUnit(ByVal vData As Double)
    mvardCustoUnit = vData
End Property

Public Property Get dCustoUnit() As Double
    dCustoUnit = mvardCustoUnit
End Property

Public Property Let dPercentMargem(ByVal vData As Double)
    mvardPercentMargem = vData
End Property

Public Property Get dPercentMargem() As Double
    dPercentMargem = mvardPercentMargem
End Property

Public Property Let dPrecoUnit(ByVal vData As Double)
    mvardPrecoUnit = vData
End Property

Public Property Get dPrecoUnit() As Double
    dPrecoUnit = mvardPrecoUnit
End Property

Public Property Let dPrecoTotal(ByVal vData As Double)
    mvardPrecoTotal = vData
End Property

Public Property Get dPrecoTotal() As Double
    dPrecoTotal = mvardPrecoTotal
End Property

Public Function Copiar(ByVal objItemOrig As ClassItemFormPreco) As ClassItemFormPreco
    With Me
        .lNumIntDoc = objItemOrig.lNumIntDoc
        .iTipoDocOrigem = objItemOrig.iTipoDocOrigem
        .lNumIntDocOrigem = objItemOrig.lNumIntDocOrigem
        .iSequencial = objItemOrig.iSequencial
        .sProduto = objItemOrig.sProduto
        .sUnidMed = objItemOrig.sUnidMed
        .dQtde = objItemOrig.dQtde
        .dCustoUnit = objItemOrig.dCustoUnit
        .dPercentMargem = objItemOrig.dPercentMargem
        .dPrecoUnit = objItemOrig.dPrecoUnit
        .dPrecoTotal = objItemOrig.dPrecoTotal
        .iSituacao = objItemOrig.iSituacao
    End With
    Set Copiar = Me
End Function

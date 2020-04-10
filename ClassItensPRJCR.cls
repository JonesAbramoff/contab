VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItensPRJCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsDescricao As String
Private mvardValorOriginal As Double
Private mvardQuantidadeOriginal As Double
Private mvariTipoOrigem As Integer
Private mvarsItem As String
Private mvariNegativo As Integer
Private mvarsObservacao As String
Private mvarsNatureza As String
Private mvarsCcl As String
Private mvarlNumIntPRJ As Long
Private mvarlNumIntEtapa As Long
Private mvarlNumMovto As Long
Private mvarsTipoDoc As String
Private mvarsNumDoc As String
Private mvarsTipoCliForn As String
Private mvarsCliForn As String
Private mvarsNomeCliForn As String
Private mvarsNomeRedCliForn As String
Private mvardtDataEmissao As Date
Private mvardtDataVencimento As Date

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let sNomeRedCliForn(ByVal vData As String)
    mvarsNomeRedCliForn = vData
End Property

Public Property Get sNomeRedCliForn() As String
    sNomeRedCliForn = mvarsNomeRedCliForn
End Property

Public Property Let sNomeCliForn(ByVal vData As String)
    mvarsNomeCliForn = vData
End Property

Public Property Get sNomeCliForn() As String
    sNomeCliForn = mvarsNomeCliForn
End Property

Public Property Let sCliForn(ByVal vData As String)
    mvarsCliForn = vData
End Property

Public Property Get sCliForn() As String
    sCliForn = mvarsCliForn
End Property

Public Property Let sTipoCliForn(ByVal vData As String)
    mvarsTipoCliForn = vData
End Property

Public Property Get sTipoCliForn() As String
    sTipoCliForn = mvarsTipoCliForn
End Property

Public Property Let sNumDoc(ByVal vData As String)
    mvarsNumDoc = vData
End Property

Public Property Get sNumDoc() As String
    sNumDoc = mvarsNumDoc
End Property

Public Property Let sTipoDoc(ByVal vData As String)
    mvarsTipoDoc = vData
End Property

Public Property Get sTipoDoc() As String
    sTipoDoc = mvarsTipoDoc
End Property

Public Property Let lNumIntPRJ(ByVal vData As Long)
    mvarlNumIntPRJ = vData
End Property

Public Property Get lNumIntPRJ() As Long
    lNumIntPRJ = mvarlNumIntPRJ
End Property

Public Property Let lNumMovto(ByVal vData As Long)
    mvarlNumMovto = vData
End Property

Public Property Get lNumMovto() As Long
    lNumMovto = mvarlNumMovto
End Property

Public Property Let lNumIntEtapa(ByVal vData As Long)
    mvarlNumIntEtapa = vData
End Property

Public Property Get lNumIntEtapa() As Long
    lNumIntEtapa = mvarlNumIntEtapa
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let sNatureza(ByVal vData As String)
    mvarsNatureza = vData
End Property

Public Property Get sNatureza() As String
    sNatureza = mvarsNatureza
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let iNegativo(ByVal vData As Integer)
    mvariNegativo = vData
End Property

Public Property Get iNegativo() As Integer
    iNegativo = mvariNegativo
End Property

Public Property Let sItem(ByVal vData As String)
    mvarsItem = vData
End Property

Public Property Get sItem() As String
    sItem = mvarsItem
End Property

Public Property Let iTipoOrigem(ByVal vData As Integer)
    mvariTipoOrigem = vData
End Property

Public Property Get iTipoOrigem() As Integer
    iTipoOrigem = mvariTipoOrigem
End Property

Public Property Let dQuantidadeOriginal(ByVal vData As Double)
    mvardQuantidadeOriginal = vData
End Property

Public Property Get dQuantidadeOriginal() As Double
    dQuantidadeOriginal = mvardQuantidadeOriginal
End Property

Public Property Let dValorOriginal(ByVal vData As Double)
    mvardValorOriginal = vData
End Property

Public Property Get dValorOriginal() As Double
    dValorOriginal = mvardValorOriginal
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property


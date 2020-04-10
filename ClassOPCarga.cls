VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOPCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardCarga As Double
Private mvarsDescricao As String
Private mvarsItemOPDescricao As String
Private mvarsLote As String
Private mvarsProduto As String
Private mvarsItemOPProduto As String
Private mvardQuantidade As Double
Private mvardQtdItemOPProduzida As Double
Private mvardQtdItemOP As Double
Private mvardQtdCarga As Double
Private mvardAcerto As Double
Private mvardFormFinal As Double
Private mvardtDataIni As Date
Private mvardtDataFim As Date
Private mvariSeq As Integer
Private mvarlItemOP As Long
Private mvariOrd As Integer
Private mvarsCodigoOP As String
Private mvarsMaquina1 As String
Private mvarsMaquina2 As String
Private mvariFilialEmpresa As Integer
Private mvarsResponsavel As String
Private mvarsNomeFornecedor As String
Private mvarlFornecedor As Long
Private mvarlRequisicao(1 To 3) As Long
Private mvariAlmoxarifado(1 To 3) As Integer
Private mvarsAlmoxarifado(1 To 3) As String
Private mvariNumRequisicoes As Integer
Private mvardQuantReq(1 To 3) As Double
Private mvarsLocalizacao As String
Private mvarsGrupoPesagem As String

Public Property Let sGrupoPesagem(ByVal vData As String)
    mvarsGrupoPesagem = vData
End Property

Public Property Get sGrupoPesagem() As String
    sGrupoPesagem = mvarsGrupoPesagem
End Property

Public Property Let sLocalizacao(ByVal vData As String)
    mvarsLocalizacao = vData
End Property

Public Property Get sLocalizacao() As String
    sLocalizacao = mvarsLocalizacao
End Property

Public Property Let dQuantReq(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQuantReq(iIndice) = vData
End Property

Public Property Get dQuantReq(ByVal iIndice As Integer) As Double
    dQuantReq = mvardQuantReq(iIndice)
End Property

Public Property Let iNumRequisicoes(ByVal vData As Integer)
    mvariNumRequisicoes = vData
End Property

Public Property Get iNumRequisicoes() As Integer
    iNumRequisicoes = mvariNumRequisicoes
End Property

Public Property Let sAlmoxarifado(ByVal iIndice As Integer, ByVal vData As String)
    mvarsAlmoxarifado(iIndice) = vData
End Property

Public Property Get sAlmoxarifado(ByVal iIndice As Integer) As String
    sAlmoxarifado = mvarsAlmoxarifado(iIndice)
End Property

Public Property Let iAlmoxarifado(ByVal iIndice As Integer, ByVal vData As Integer)
    mvariAlmoxarifado(iIndice) = vData
End Property

Public Property Get iAlmoxarifado(ByVal iIndice As Integer) As Integer
    iAlmoxarifado = mvariAlmoxarifado(iIndice)
End Property

Public Property Let lRequisicao(ByVal iIndice As Integer, ByVal vData As Long)
    mvarlRequisicao(iIndice) = vData
End Property

Public Property Get lRequisicao(ByVal iIndice As Integer) As Long
    lRequisicao = mvarlRequisicao(iIndice)
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let sNomeFornecedor(ByVal vData As String)
    mvarsNomeFornecedor = vData
End Property

Public Property Get sNomeFornecedor() As String
    sNomeFornecedor = mvarsNomeFornecedor
End Property

Public Property Let sResponsavel(ByVal vData As String)
    mvarsResponsavel = vData
End Property

Public Property Get sResponsavel() As String
    sResponsavel = mvarsResponsavel
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sCodigoOP(ByVal vData As String)
    mvarsCodigoOP = vData
End Property

Public Property Get sCodigoOP() As String
    sCodigoOP = mvarsCodigoOP
End Property

Public Property Let sMaquina1(ByVal vData As String)
    mvarsMaquina1 = vData
End Property

Public Property Get sMaquina1() As String
    sMaquina1 = mvarsMaquina1
End Property

Public Property Let sMaquina2(ByVal vData As String)
    mvarsMaquina2 = vData
End Property

Public Property Get sMaquina2() As String
    sMaquina2 = mvarsMaquina2
End Property

Public Property Let iOrd(ByVal vData As Integer)
    mvariOrd = vData
End Property

Public Property Get iOrd() As Integer
    iOrd = mvariOrd
End Property

Public Property Let dtDataIni(ByVal vData As Date)
    mvardtDataIni = vData
End Property

Public Property Get dtDataIni() As Date
    dtDataIni = mvardtDataIni
End Property

Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Let lItemOP(ByVal vData As Long)
    mvarlItemOP = vData
End Property

Public Property Get lItemOP() As Long
    lItemOP = mvarlItemOP
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let dFormFinal(ByVal vData As Double)
    mvardFormFinal = vData
End Property

Public Property Get dFormFinal() As Double
    dFormFinal = mvardFormFinal
End Property

Public Property Let dAcerto(ByVal vData As Double)
    mvardAcerto = vData
End Property

Public Property Get dAcerto() As Double
    dAcerto = mvardAcerto
End Property

Public Property Let dQtdItemOP(ByVal vData As Double)
    mvardQtdItemOP = vData
End Property

Public Property Get dQtdItemOP() As Double
    dQtdItemOP = mvardQtdItemOP
End Property

Public Property Let dQtdItemOPProduzida(ByVal vData As Double)
    mvardQtdItemOPProduzida = vData
End Property

Public Property Get dQtdItemOPProduzida() As Double
    dQtdItemOPProduzida = mvardQtdItemOPProduzida
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dQtdCarga(ByVal vData As Double)
    mvardQtdCarga = vData
End Property

Public Property Get dQtdCarga() As Double
    dQtdCarga = mvardQtdCarga
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sItemOPProduto(ByVal vData As String)
    mvarsItemOPProduto = vData
End Property

Public Property Get sItemOPProduto() As String
    sItemOPProduto = mvarsItemOPProduto
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sItemOPDescricao(ByVal vData As String)
    mvarsItemOPDescricao = vData
End Property

Public Property Get sItemOPDescricao() As String
    sItemOPDescricao = mvarsItemOPDescricao
End Property

Public Property Let dCarga(ByVal vData As Double)
    mvardCarga = vData
End Property

Public Property Get dCarga() As Double
    dCarga = mvardCarga
End Property




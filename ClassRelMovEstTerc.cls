VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMovEstTerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel As Long
Private mvarlSeq As Long
Private mvarlNumMovto As Long
Private mvariEscaninho As Integer
Private mvariTipoTerc As Integer
Private mvarlCodTerc As Long
Private mvariFilialTerc As Integer
Private mvarsProduto As String
Private mvarsNomeEscaninho As String
Private mvarsNomeTerceiro As String
Private mvarsNomeTipoTerc As String
Private mvarsNomeFilialTerc As String
Private mvarsDocOrigem As String
Private mvarsDescTipoMovto As String
Private mvardtData As Date
Private mvarsUM As String
Private mvardQtdeEntrada As Double
Private mvardQtdeSaida As Double
Private mvardSaldo As Double
Private mvarsTipoEscaninho As String
Private mvarlQuebraTerc As Long
Private mvarsFaixaTerc As String
Private mvarsFaixaProd As String
Private mvarsFaixaCat As String
Private mvarsFaixaData As String
Private mvardSaldoInicial As Double
Private mvarsProdutoMask As String
Private mvarsTipoProd As String

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let lSeq(ByVal vData As Long)
    mvarlSeq = vData
End Property

Public Property Get lSeq() As Long
    lSeq = mvarlSeq
End Property

Public Property Let lNumMovto(ByVal vData As Long)
    mvarlNumMovto = vData
End Property

Public Property Get lNumMovto() As Long
    lNumMovto = mvarlNumMovto
End Property

Public Property Let iEscaninho(ByVal vData As Integer)
    mvariEscaninho = vData
End Property

Public Property Get iEscaninho() As Integer
    iEscaninho = mvariEscaninho
End Property

Public Property Let iTipoTerc(ByVal vData As Integer)
    mvariTipoTerc = vData
End Property

Public Property Get iTipoTerc() As Integer
    iTipoTerc = mvariTipoTerc
End Property

Public Property Let lCodTerc(ByVal vData As Long)
    mvarlCodTerc = vData
End Property

Public Property Get lCodTerc() As Long
    lCodTerc = mvarlCodTerc
End Property

Public Property Let iFilialTerc(ByVal vData As Integer)
    mvariFilialTerc = vData
End Property

Public Property Get iFilialTerc() As Integer
    iFilialTerc = mvariFilialTerc
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sNomeEscaninho(ByVal vData As String)
    mvarsNomeEscaninho = vData
End Property

Public Property Get sNomeEscaninho() As String
    sNomeEscaninho = mvarsNomeEscaninho
End Property

Public Property Let sNomeTerceiro(ByVal vData As String)
    mvarsNomeTerceiro = vData
End Property

Public Property Get sNomeTerceiro() As String
    sNomeTerceiro = mvarsNomeTerceiro
End Property

Public Property Let sNomeTipoTerc(ByVal vData As String)
    mvarsNomeTipoTerc = vData
End Property

Public Property Get sNomeTipoTerc() As String
    sNomeTipoTerc = mvarsNomeTipoTerc
End Property

Public Property Let sNomeFilialTerc(ByVal vData As String)
    mvarsNomeFilialTerc = vData
End Property

Public Property Get sNomeFilialTerc() As String
    sNomeFilialTerc = mvarsNomeFilialTerc
End Property

Public Property Let sDocOrigem(ByVal vData As String)
    mvarsDocOrigem = vData
End Property

Public Property Get sDocOrigem() As String
    sDocOrigem = mvarsDocOrigem
End Property

Public Property Let sDescTipoMovto(ByVal vData As String)
    mvarsDescTipoMovto = vData
End Property

Public Property Get sDescTipoMovto() As String
    sDescTipoMovto = mvarsDescTipoMovto
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dQtdeEntrada(ByVal vData As Double)
    mvardQtdeEntrada = vData
End Property

Public Property Get dQtdeEntrada() As Double
    dQtdeEntrada = mvardQtdeEntrada
End Property

Public Property Let dQtdeSaida(ByVal vData As Double)
    mvardQtdeSaida = vData
End Property

Public Property Get dQtdeSaida() As Double
    dQtdeSaida = mvardQtdeSaida
End Property

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
End Property

Public Property Let sTipoEscaninho(ByVal vData As String)
    mvarsTipoEscaninho = vData
End Property

Public Property Get sTipoEscaninho() As String
    sTipoEscaninho = mvarsTipoEscaninho
End Property

Public Property Let lQuebraTerc(ByVal vData As Long)
    mvarlQuebraTerc = vData
End Property

Public Property Get lQuebraTerc() As Long
    lQuebraTerc = mvarlQuebraTerc
End Property

Public Property Let sFaixaTerc(ByVal vData As String)
    mvarsFaixaTerc = vData
End Property

Public Property Get sFaixaTerc() As String
    sFaixaTerc = mvarsFaixaTerc
End Property

Public Property Let sFaixaProd(ByVal vData As String)
    mvarsFaixaProd = vData
End Property

Public Property Get sFaixaProd() As String
    sFaixaProd = mvarsFaixaProd
End Property

Public Property Let sFaixaCat(ByVal vData As String)
    mvarsFaixaCat = vData
End Property

Public Property Get sFaixaCat() As String
    sFaixaCat = mvarsFaixaCat
End Property

Public Property Let sFaixaData(ByVal vData As String)
    mvarsFaixaData = vData
End Property

Public Property Get sFaixaData() As String
    sFaixaData = mvarsFaixaData
End Property

Public Property Let dSaldoInicial(ByVal vData As Double)
    mvardSaldoInicial = vData
End Property

Public Property Get dSaldoInicial() As Double
    dSaldoInicial = mvardSaldoInicial
End Property

Public Property Let sProdutoMask(ByVal vData As String)
    mvarsProdutoMask = vData
End Property

Public Property Get sProdutoMask() As String
    sProdutoMask = mvarsProdutoMask
End Property

Public Property Let sTipoProd(ByVal vData As String)
    mvarsTipoProd = vData
End Property

Public Property Get sTipoProd() As String
    sTipoProd = mvarsTipoProd
End Property

Public Sub Copia(ByVal objDoc As ClassRelMovEstTerc)
    lNumIntRel = objDoc.lNumIntRel
    lSeq = objDoc.lSeq
    lNumMovto = objDoc.lNumMovto
    iEscaninho = objDoc.iEscaninho
    iTipoTerc = objDoc.iTipoTerc
    lCodTerc = objDoc.lCodTerc
    iFilialTerc = objDoc.iFilialTerc
    sProduto = objDoc.sProduto
    sNomeEscaninho = objDoc.sNomeEscaninho
    sNomeTerceiro = objDoc.sNomeTerceiro
    sNomeTipoTerc = objDoc.sNomeTipoTerc
    sNomeFilialTerc = objDoc.sNomeFilialTerc
    sDocOrigem = objDoc.sDocOrigem
    sDescTipoMovto = objDoc.sDescTipoMovto
    dtData = objDoc.dtData
    sUM = objDoc.sUM
    dQtdeEntrada = objDoc.dQtdeEntrada
    dQtdeSaida = objDoc.dQtdeSaida
    dSaldo = objDoc.dSaldo
    sTipoEscaninho = objDoc.sTipoEscaninho
    lQuebraTerc = objDoc.lQuebraTerc
    sFaixaTerc = objDoc.sFaixaTerc
    sFaixaProd = objDoc.sFaixaProd
    sFaixaCat = objDoc.sFaixaCat
    sFaixaData = objDoc.sFaixaData
    dSaldoInicial = objDoc.dSaldoInicial
    sProdutoMask = objDoc.sProdutoMask
    sTipoProd = objDoc.sTipoProd
End Sub

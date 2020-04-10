VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPadraoContabItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsModulo As String
Private mvariTransacao As Integer
Private mvarsModelo As String
Private mvariItem As Integer
Private mvarsConta As String
Private mvarsCcl As String
Private mvarsCredito As String
Private mvarsDebito As String
Private mvarsHistorico As String
Private mvarsProduto As String
Private mvariCusto As Integer
Private mvariAglutina As Integer
Private mvarsEscaninho_Custo As String
Private mvariSubTipo As Integer
Private mvariGerencial As Integer

Public Property Let iGerencial(ByVal vData As Integer)
    mvariGerencial = vData
End Property

Public Property Get iGerencial() As Integer
    iGerencial = mvariGerencial
End Property

Public Property Let iSubTipo(ByVal vData As Integer)
    mvariSubTipo = vData
End Property

Public Property Get iSubTipo() As Integer
    iSubTipo = mvariSubTipo
End Property

Public Property Let sEscaninho_Custo(ByVal vData As String)
    mvarsEscaninho_Custo = vData
End Property

Public Property Get sEscaninho_Custo() As String
    sEscaninho_Custo = mvarsEscaninho_Custo
End Property

Public Property Let iAglutina(ByVal vData As Integer)
    mvariAglutina = vData
End Property

Public Property Get iAglutina() As Integer
    iAglutina = mvariAglutina
End Property

Public Property Let iCusto(ByVal vData As Integer)
Attribute iCusto.VB_Description = "1=Indica que o lançamento pertence a custo 0=indica que é um lançamento normal"
    mvariCusto = vData
End Property

Public Property Get iCusto() As Integer
    iCusto = mvariCusto
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sHistorico(ByVal vData As String)
Attribute sHistorico.VB_Description = "Formula que gera o historico"
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let sDebito(ByVal vData As String)
Attribute sDebito.VB_Description = "Formula que gera o valor do débito"
    mvarsDebito = vData
End Property

Public Property Get sDebito() As String
    sDebito = mvarsDebito
End Property

Public Property Let sCredito(ByVal vData As String)
Attribute sCredito.VB_Description = "Formula que gera o valor do crédito"
    mvarsCredito = vData
End Property

Public Property Get sCredito() As String
    sCredito = mvarsCredito
End Property

Public Property Let sCcl(ByVal vData As String)
Attribute sCcl.VB_Description = "Formula que gera o centro de custo/lucro"
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let sConta(ByVal vData As String)
Attribute sConta.VB_Description = "Formula que gera a conta contábil"
    mvarsConta = vData
End Property

Public Property Get sConta() As String
    sConta = mvarsConta
End Property

Public Property Let iItem(ByVal vData As Integer)
Attribute iItem.VB_Description = "Item do grid"
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let sModelo(ByVal vData As String)
Attribute sModelo.VB_Description = "Nome do Modelo"
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let iTransacao(ByVal vData As Integer)
Attribute iTransacao.VB_Description = "Nome da Transacao"
    mvariTransacao = vData
End Property

Public Property Get iTransacao() As Integer
    iTransacao = mvariTransacao
End Property

Public Property Let sModulo(ByVal vData As String)
Attribute sModulo.VB_Description = "Sigla do Modulo"
    mvarsModulo = vData
End Property

Public Property Get sModulo() As String
    sModulo = mvarsModulo
End Property




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImportTabelasCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlTabela As Long
Private mvarlCodigo As Long
Private mvarsCampo As String
Private mvarsNomeExibicao As String
Private mvariTipo As Integer
Private mvariChave As Integer
Private mvariExibe As Integer
Private mvarsNomeIgual1 As String
Private mvarsNomeIgual2 As String
Private mvarsNomeIgual3 As String
Private mvarsNomeLike1 As String
Private mvarsNomeLike2 As String
Private mvarsNomeLike3 As String
Private mvarsValorPadrao As String
Private mvariObrigatorio As Integer
Private mvariTamMax As Integer

Public Property Let iTamMax(ByVal vData As Integer)
    mvariTamMax = vData
End Property

Public Property Get iTamMax() As Integer
    iTamMax = mvariTamMax
End Property

Public Property Let iObrigatorio(ByVal vData As Integer)
    mvariObrigatorio = vData
End Property

Public Property Get iObrigatorio() As Integer
    iObrigatorio = mvariObrigatorio
End Property

Public Property Let lTabela(ByVal vData As Long)
    mvarlTabela = vData
End Property

Public Property Get lTabela() As Long
    lTabela = mvarlTabela
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sCampo(ByVal vData As String)
    mvarsCampo = vData
End Property

Public Property Get sCampo() As String
    sCampo = mvarsCampo
End Property

Public Property Let sNomeExibicao(ByVal vData As String)
    mvarsNomeExibicao = vData
End Property

Public Property Get sNomeExibicao() As String
    sNomeExibicao = mvarsNomeExibicao
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iChave(ByVal vData As Integer)
    mvariChave = vData
End Property

Public Property Get iChave() As Integer
    iChave = mvariChave
End Property

Public Property Let iExibe(ByVal vData As Integer)
    mvariExibe = vData
End Property

Public Property Get iExibe() As Integer
    iExibe = mvariExibe
End Property

Public Property Let sNomeIgual1(ByVal vData As String)
    mvarsNomeIgual1 = vData
End Property

Public Property Get sNomeIgual1() As String
    sNomeIgual1 = mvarsNomeIgual1
End Property

Public Property Let sNomeIgual2(ByVal vData As String)
    mvarsNomeIgual2 = vData
End Property

Public Property Get sNomeIgual2() As String
    sNomeIgual2 = mvarsNomeIgual2
End Property

Public Property Let sNomeIgual3(ByVal vData As String)
    mvarsNomeIgual3 = vData
End Property

Public Property Get sNomeIgual3() As String
    sNomeIgual3 = mvarsNomeIgual3
End Property

Public Property Let sNomeLike1(ByVal vData As String)
    mvarsNomeLike1 = vData
End Property

Public Property Get sNomeLike1() As String
    sNomeLike1 = mvarsNomeLike1
End Property

Public Property Let sNomeLike2(ByVal vData As String)
    mvarsNomeLike2 = vData
End Property

Public Property Get sNomeLike2() As String
    sNomeLike2 = mvarsNomeLike2
End Property

Public Property Let sNomeLike3(ByVal vData As String)
    mvarsNomeLike3 = vData
End Property

Public Property Get sNomeLike3() As String
    sNomeLike3 = mvarsNomeLike3
End Property

Public Property Let sValorPadrao(ByVal vData As String)
    mvarsValorPadrao = vData
End Property

Public Property Get sValorPadrao() As String
    sValorPadrao = mvarsValorPadrao
End Property

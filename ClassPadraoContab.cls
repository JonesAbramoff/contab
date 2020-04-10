VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPadraoContab"
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
Private mvariPadrao As Integer
Private mvariSubTipo As Integer
Private mvariObrigatoriedade As Integer

Public Property Let iObrigatoriedade(ByVal vData As Integer)
    mvariObrigatoriedade = vData
End Property

Public Property Get iObrigatoriedade() As Integer
    iObrigatoriedade = mvariObrigatoriedade
End Property

Public Property Let iSubtipo(ByVal vData As Integer)
    mvariSubTipo = vData
End Property

Public Property Get iSubtipo() As Integer
    iSubtipo = mvariSubTipo
End Property

Public Property Let iPadrao(ByVal vData As Integer)
Attribute iPadrao.VB_Description = "1=é o modelo padrão desta transação 0=não é o modelo padrão"
    mvariPadrao = vData
End Property

Public Property Get iPadrao() As Integer
    iPadrao = mvariPadrao
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




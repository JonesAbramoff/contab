VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMenuItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsTitulo As String
Private mvariIdentificador As Integer
Private mvarsSiglaRotina As String
Private mvarsNomeTela As String
Private mvarsNomeControle As String
Private mvariIndiceControle As Integer
Private mvarsNomeControlePai As String
Private mvariIndiceControlePai As Integer
Private mvariSeparador As Integer
Private mvariIndiceControleNoPai As Integer

Public Property Let iIndiceControleNoPai(ByVal vData As Integer)
    mvariIndiceControleNoPai = vData
End Property

Public Property Get iIndiceControleNoPai() As Integer
    iIndiceControleNoPai = mvariIndiceControleNoPai
End Property

Public Property Let iSeparador(ByVal vData As Integer)
    mvariSeparador = vData
End Property

Public Property Get iSeparador() As Integer
    iSeparador = mvariSeparador
End Property

Public Property Let iIndiceControlePai(ByVal vData As Integer)
Attribute iIndiceControlePai.VB_Description = "Indice do pai do item de menu no VB"
    mvariIndiceControlePai = vData
End Property

Public Property Get iIndiceControlePai() As Integer
    iIndiceControlePai = mvariIndiceControlePai
End Property

Public Property Let sNomeControlePai(ByVal vData As String)
Attribute sNomeControlePai.VB_Description = "Nome do pai do  ítem de menu no VB"
    mvarsNomeControlePai = vData
End Property

Public Property Get sNomeControlePai() As String
    sNomeControlePai = mvarsNomeControlePai
End Property

Public Property Let iIndiceControle(ByVal vData As Integer)
Attribute iIndiceControle.VB_Description = "Indice do item de menu no VB"
    mvariIndiceControle = vData
End Property

Public Property Get iIndiceControle() As Integer
    iIndiceControle = mvariIndiceControle
End Property

Public Property Let sNomeControle(ByVal vData As String)
Attribute sNomeControle.VB_Description = "Nome do  ítem de menu no VB"
    mvarsNomeControle = vData
End Property

Public Property Get sNomeControle() As String
    sNomeControle = mvarsNomeControle
End Property

Public Property Let sNomeTela(ByVal vData As String)
Attribute sNomeTela.VB_Description = "Nome de Tela - este campo ou (exclusivo) o de SiglaRotina estará preenchido"
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Let sSiglaRotina(ByVal vData As String)
Attribute sSiglaRotina.VB_Description = "Sigla de Rotina"
    mvarsSiglaRotina = vData
End Property

Public Property Get sSiglaRotina() As String
    sSiglaRotina = mvarsSiglaRotina
End Property

Public Property Let iIdentificador(ByVal vData As Integer)
Attribute iIdentificador.VB_Description = "Identificador de Item de Menu"
    mvariIdentificador = vData
End Property

Public Property Get iIdentificador() As Integer
    iIdentificador = mvariIdentificador
End Property

Public Property Let sTitulo(ByVal vData As String)
Attribute sTitulo.VB_Description = "Titulo (caption) de ítem de menu"
    mvarsTitulo = vData
End Property

Public Property Get sTitulo() As String
    sTitulo = mvarsTitulo
End Property




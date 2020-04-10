VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoTribIPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariPermiteAliquota As Integer
Private mvariPermiteReducaoBase As Integer
Private mvariTipo As Integer
Private mvarsDescricao As String
Private mvariColunaNoLivro As Integer
Private mvariVersaoNFE As Integer
Private mvariAtivo As Integer
Private mvariCSTEntrada As Integer
Private mvariCSTSaida As Integer
Private mvariTipoCalculo As Integer

'**** ATENCAO: QUANDO INCLUIR NOVO CAMPO ATUALIZAR A SUB COPIAR ******'

Public Property Let iTipoCalculo(ByVal vData As Integer)
    mvariTipoCalculo = vData
End Property

Public Property Get iTipoCalculo() As Integer
    iTipoCalculo = mvariTipoCalculo
End Property

Public Property Let iCSTEntrada(ByVal vData As Integer)
    mvariCSTEntrada = vData
End Property

Public Property Get iCSTEntrada() As Integer
    iCSTEntrada = mvariCSTEntrada
End Property

Public Property Let iCSTSaida(ByVal vData As Integer)
    mvariCSTSaida = vData
End Property

Public Property Get iCSTSaida() As Integer
    iCSTSaida = mvariCSTSaida
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iVersaoNFE(ByVal vData As Integer)
    mvariVersaoNFE = vData
End Property

Public Property Get iVersaoNFE() As Integer
    iVersaoNFE = mvariVersaoNFE
End Property

Public Property Let iColunaNoLivro(ByVal vData As Integer)
    mvariColunaNoLivro = vData
End Property

Public Property Get iColunaNoLivro() As Integer
    iColunaNoLivro = mvariColunaNoLivro
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iPermiteReducaoBase(ByVal vData As Integer)
    mvariPermiteReducaoBase = vData
End Property

Public Property Get iPermiteReducaoBase() As Integer
    iPermiteReducaoBase = mvariPermiteReducaoBase
End Property

Public Property Let iPermiteAliquota(ByVal vData As Integer)
    mvariPermiteAliquota = vData
End Property

Public Property Get iPermiteAliquota() As Integer
    iPermiteAliquota = mvariPermiteAliquota
End Property

Public Sub Copiar(ByVal objTipoTribIPIOrig As ClassTipoTribIPI)
'copia os dados de objTributacaoTipo p/"Me"

    With objTipoTribIPIOrig
        mvariTipo = .iTipo
        mvarsDescricao = .sDescricao
        mvariPermiteAliquota = .iPermiteAliquota
        mvariPermiteReducaoBase = .iPermiteReducaoBase
        mvariColunaNoLivro = .iColunaNoLivro
        mvariVersaoNFE = .iVersaoNFE
        mvariCSTEntrada = .iCSTEntrada
        mvariCSTSaida = .iCSTSaida
        mvariTipoCalculo = .iTipoCalculo
    End With
    
End Sub


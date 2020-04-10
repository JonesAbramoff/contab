VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoTribICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariTipo As Integer
Private mvarsDescricao As String
Private mvariPermiteAliquota As Integer
Private mvariPermiteMargLucro As Integer
Private mvariPermiteReducaoBase As Integer
Private mvariColunaNoLivroEntrada As Integer
Private mvariColunaNoLivroSaida As Integer
Private mvariTipoTribCST As Integer
Private mvariVersaoNFE As Integer
Private mvariPermiteST As Integer
Private mvariPermitePercBaseOperProp As Integer
Private mvariPermiteUFICMSSTDevido As Integer
Private mvariPermiteVlrICMSIsento As Integer
Private mvariPermiteMotDesoneracao As Integer
Private mvariPermiteSTRetUFRem As Integer
Private mvariPermiteSTUFDest As Integer
Private mvariPermiteSTRetAnt As Integer

'**** ATENCAO: QUANDO INCLUIR NOVO CAMPO ATUALIZAR A SUB COPIAR ******'

Public Property Let iPermiteVlrICMSIsento(ByVal vData As Integer)
    mvariPermiteVlrICMSIsento = vData
End Property

Public Property Get iPermiteVlrICMSIsento() As Integer
    iPermiteVlrICMSIsento = mvariPermiteVlrICMSIsento
End Property

Public Property Let iPermiteMotDesoneracao(ByVal vData As Integer)
    mvariPermiteMotDesoneracao = vData
End Property

Public Property Get iPermiteMotDesoneracao() As Integer
    iPermiteMotDesoneracao = mvariPermiteMotDesoneracao
End Property

Public Property Let iPermiteSTRetUFRem(ByVal vData As Integer)
    mvariPermiteSTRetUFRem = vData
End Property

Public Property Get iPermiteSTRetUFRem() As Integer
    iPermiteSTRetUFRem = mvariPermiteSTRetUFRem
End Property

Public Property Let iPermiteSTUFDest(ByVal vData As Integer)
    mvariPermiteSTUFDest = vData
End Property

Public Property Get iPermiteSTUFDest() As Integer
    iPermiteSTUFDest = mvariPermiteSTUFDest
End Property

Public Property Let iPermiteSTRetAnt(ByVal vData As Integer)
    mvariPermiteSTRetAnt = vData
End Property

Public Property Get iPermiteSTRetAnt() As Integer
    iPermiteSTRetAnt = mvariPermiteSTRetAnt
End Property

Public Property Let iVersaoNFE(ByVal vData As Integer)
    mvariVersaoNFE = vData
End Property

Public Property Get iVersaoNFE() As Integer
    iVersaoNFE = mvariVersaoNFE
End Property

Public Property Let iPermiteST(ByVal vData As Integer)
    mvariPermiteST = vData
End Property

Public Property Get iPermiteST() As Integer
    iPermiteST = mvariPermiteST
End Property

Public Property Let iPermitePercBaseOperProp(ByVal vData As Integer)
    mvariPermitePercBaseOperProp = vData
End Property

Public Property Get iPermitePercBaseOperProp() As Integer
    iPermitePercBaseOperProp = mvariPermitePercBaseOperProp
End Property

Public Property Let iPermiteUFICMSSTDevido(ByVal vData As Integer)
    mvariPermiteUFICMSSTDevido = vData
End Property

Public Property Get iPermiteUFICMSSTDevido() As Integer
    iPermiteUFICMSSTDevido = mvariPermiteUFICMSSTDevido
End Property

Public Property Let iTipoTribCST(ByVal vData As Integer)
    mvariTipoTribCST = vData
End Property

Public Property Get iTipoTribCST() As Integer
    iTipoTribCST = mvariTipoTribCST
End Property

Public Property Let iColunaNoLivroSaida(ByVal vData As Integer)
    mvariColunaNoLivroSaida = vData
End Property

Public Property Get iColunaNoLivroSaida() As Integer
    iColunaNoLivroSaida = mvariColunaNoLivroSaida
End Property

Public Property Let iColunaNoLivroEntrada(ByVal vData As Integer)
    mvariColunaNoLivroEntrada = vData
End Property

Public Property Get iColunaNoLivroEntrada() As Integer
    iColunaNoLivroEntrada = mvariColunaNoLivroEntrada
End Property

Public Property Let iPermiteReducaoBase(ByVal vData As Integer)
    mvariPermiteReducaoBase = vData
End Property

Public Property Get iPermiteReducaoBase() As Integer
    iPermiteReducaoBase = mvariPermiteReducaoBase
End Property

Public Property Let iPermiteMargLucro(ByVal vData As Integer)
    mvariPermiteMargLucro = vData
End Property

Public Property Get iPermiteMargLucro() As Integer
    iPermiteMargLucro = mvariPermiteMargLucro
End Property

Public Property Let iPermiteAliquota(ByVal vData As Integer)
    mvariPermiteAliquota = vData
End Property

Public Property Get iPermiteAliquota() As Integer
    iPermiteAliquota = mvariPermiteAliquota
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

Public Sub Copiar(ByVal objTipoTribICMSOrig As ClassTipoTribICMS)
'copia os dados de objTributacaoTipo p/"Me"

    With objTipoTribICMSOrig
        mvariTipo = .iTipo
        mvarsDescricao = .sDescricao
        mvariPermiteAliquota = .iPermiteAliquota
        mvariPermiteMargLucro = .iPermiteMargLucro
        mvariPermiteReducaoBase = .iPermiteReducaoBase
        mvariColunaNoLivroEntrada = .iColunaNoLivroEntrada
        mvariColunaNoLivroSaida = .iColunaNoLivroSaida
        mvariTipoTribCST = .iTipoTribCST
        mvariVersaoNFE = .iVersaoNFE
        mvariPermiteST = .iPermiteST
        mvariPermitePercBaseOperProp = .iPermitePercBaseOperProp
        mvariPermiteUFICMSSTDevido = .iPermiteUFICMSSTDevido
        mvariPermiteVlrICMSIsento = .iPermiteVlrICMSIsento
        mvariPermiteMotDesoneracao = .iPermiteMotDesoneracao
        mvariPermiteSTRetUFRem = .iPermiteSTRetUFRem
        mvariPermiteSTUFDest = .iPermiteSTUFDest
        mvariPermiteSTRetAnt = .iPermiteSTRetAnt
    End With
    
End Sub

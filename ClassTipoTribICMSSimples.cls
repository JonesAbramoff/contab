VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoTribICMSSimples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariTipo As Integer
Private mvarsDescricao As String
Private mvariPermiteCredito As Integer
Private mvariColunaNoLivroEntrada As Integer
Private mvariColunaNoLivroSaida As Integer
Private mvariCSOSN As Integer
Private mvariVersaoNFE As Integer
Private mvariPermiteST As Integer
Private mvariPermiteAliquota As Integer
Private mvariPermiteSTRetAnt As Integer

'**** ATENCAO: QUANDO INCLUIR NOVO CAMPO ATUALIZAR A SUB COPIAR ******'

Public Property Let iPermiteAliquota(ByVal vData As Integer)
    mvariPermiteAliquota = vData
End Property

Public Property Get iPermiteAliquota() As Integer
    iPermiteAliquota = mvariPermiteAliquota
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

Public Property Let iCSOSN(ByVal vData As Integer)
    mvariCSOSN = vData
End Property

Public Property Get iCSOSN() As Integer
    iCSOSN = mvariCSOSN
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

Public Property Let iPermiteCredito(ByVal vData As Integer)
    mvariPermiteCredito = vData
End Property

Public Property Get iPermiteCredito() As Integer
    iPermiteCredito = mvariPermiteCredito
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

Public Sub Copiar(ByVal objTipoTribICMSOrig As ClassTipoTribICMSSimples)
'copia os dados de objTributacaoTipo p/"Me"

    With objTipoTribICMSOrig
        mvariTipo = .iTipo
        mvarsDescricao = .sDescricao
        mvariPermiteCredito = .iPermiteCredito
        mvariColunaNoLivroEntrada = .iColunaNoLivroEntrada
        mvariColunaNoLivroSaida = .iColunaNoLivroSaida
        mvariCSOSN = .iCSOSN
        mvariVersaoNFE = .iVersaoNFE
        mvariPermiteST = .iPermiteST
    End With
    
End Sub

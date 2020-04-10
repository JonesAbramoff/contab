VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoTribPISCOFINS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariTipo As Integer
Private mvarsDescricao As String
Private mvariTipoCalculo As Integer
Private mvariVersaoNFE As Integer
Private mvariEntrada As Integer
Private mvariSaida As Integer

'**** ATENCAO: QUANDO INCLUIR NOVO CAMPO ATUALIZAR A SUB COPIAR ******'

Public Property Let iEntrada(ByVal vData As Integer)
    mvariEntrada = vData
End Property

Public Property Get iEntrada() As Integer
    iEntrada = mvariEntrada
End Property

Public Property Let iSaida(ByVal vData As Integer)
    mvariSaida = vData
End Property

Public Property Get iSaida() As Integer
    iSaida = mvariSaida
End Property

Public Property Let iVersaoNFE(ByVal vData As Integer)
    mvariVersaoNFE = vData
End Property

Public Property Get iVersaoNFE() As Integer
    iVersaoNFE = mvariVersaoNFE
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iTipoCalculo(ByVal vData As Integer)
    mvariTipoCalculo = vData
End Property

Public Property Get iTipoCalculo() As Integer
    iTipoCalculo = mvariTipoCalculo
End Property

Public Sub Copiar(ByVal objTipoTribPISCOFINSOrig As ClassTipoTribPISCOFINS)
'copia os dados de objTributacaoTipo p/"Me"

    With objTipoTribPISCOFINSOrig
    
        mvariTipo = .iTipo
        mvarsDescricao = .sDescricao
        mvariTipoCalculo = .iTipoCalculo
        mvariVersaoNFE = .iVersaoNFE
        mvariEntrada = .iEntrada
        mvariSaida = .iSaida
        
    End With
    
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFilialFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodFornecedor As Long
Private mvariCodFilial As Integer
Private mvarsNome As String
Private mvarsCGC As String
Private mvarsInscricaoEstadual As String
Private mvarsInscricaoMunicipal As String
Private mvarlEndereco As Long
Private mvarsObservacao As String
Private mvarsContaContabil As String
Private mvarsContaFornConsig As String
Private mvariBanco As Integer
Private mvarsAgencia As String
Private mvarsContaCorrente As String
Private mvarsContato As String
Private mvariTipoFrete As Integer
Private mvarcolCategoriaItem As New Collection

'Alteracao Daniel
Private mvarsInscricaoINSS As String
Private mvariTipoCobranca As Integer
Private mvarsContaDespesa As String
Private mvariGeraCredICMS As Integer
Private mvarsInscricaoSuframa As String
Private mvariRegimeTributario As Integer
Private mvariIEIsento As Integer
Private mvariIENaoContrib As Integer

'nfe 3.10
Public sIdEstrangeiro As String

Public Property Let iIENaoContrib(ByVal vData As Integer)
    mvariIENaoContrib = vData
End Property

Public Property Get iIENaoContrib() As Integer
    If mvariIENaoContrib = -1 Then
        iIENaoContrib = IIf(Len(Trim(mvarsInscricaoEstadual)) = 0, MARCADO, DESMARCADO)
    Else
        iIENaoContrib = mvariIENaoContrib
    End If
End Property

Public Property Let iIEIsento(ByVal vData As Integer)
    mvariIEIsento = vData
End Property

Public Property Get iIEIsento() As Integer
    If mvariIEIsento = -1 Then
        iIEIsento = IIf(Len(Trim(mvarsInscricaoEstadual)) = 0, MARCADO, DESMARCADO)
    Else
        iIEIsento = mvariIEIsento
    End If
End Property

Public Property Let iRegimeTributario(ByVal vData As Integer)
    mvariRegimeTributario = vData
End Property

Public Property Get iRegimeTributario() As Integer
    iRegimeTributario = mvariRegimeTributario
End Property

Public Property Let sInscricaoSuframa(ByVal vData As String)
    mvarsInscricaoSuframa = vData
End Property

Public Property Get sInscricaoSuframa() As String
    sInscricaoSuframa = mvarsInscricaoSuframa
End Property

Public Property Set colCategoriaItem(ByVal vData As Object)
    Set mvarcolCategoriaItem = vData
End Property

Public Property Get colCategoriaItem() As Collection
    Set colCategoriaItem = mvarcolCategoriaItem
End Property

Public Property Let iTipoFrete(ByVal vData As Integer)
    mvariTipoFrete = vData
End Property

Public Property Get iTipoFrete() As Integer
    iTipoFrete = mvariTipoFrete
End Property

Public Property Let sContato(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sContato = 5
    mvarsContato = vData
End Property


Public Property Get sContato() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sContato
    sContato = mvarsContato
End Property




Public Property Let sContaCorrente(ByVal vData As String)
    mvarsContaCorrente = vData
End Property

Public Property Get sContaCorrente() As String
    sContaCorrente = mvarsContaCorrente
End Property

Public Property Let sAgencia(ByVal vData As String)
    mvarsAgencia = vData
End Property

Public Property Get sAgencia() As String
    sAgencia = mvarsAgencia
End Property

Public Property Let iBanco(ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco() As Integer
    iBanco = mvariBanco
End Property

Public Property Let lCodFornecedor(ByVal vData As Long)
    mvarlCodFornecedor = vData
End Property

Public Property Get lCodFornecedor() As Long
    lCodFornecedor = mvarlCodFornecedor
End Property

Public Property Let sContaContabil(ByVal vData As String)
    mvarsContaContabil = vData
End Property

Public Property Get sContaContabil() As String
    sContaContabil = mvarsContaContabil
End Property

Public Property Let sContaFornConsig(ByVal vData As String)
    mvarsContaFornConsig = vData
End Property

Public Property Get sContaFornConsig() As String
    sContaFornConsig = mvarsContaFornConsig
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let lEndereco(ByVal vData As Long)
    mvarlEndereco = vData
End Property

Public Property Get lEndereco() As Long
    lEndereco = mvarlEndereco
End Property

Public Property Let sInscricaoMunicipal(ByVal vData As String)
    mvarsInscricaoMunicipal = vData
End Property

Public Property Get sInscricaoMunicipal() As String
    sInscricaoMunicipal = mvarsInscricaoMunicipal
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let sCgc(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

'Alteracao Daniel
Public Property Let sInscricaoINSS(ByVal vData As String)
    mvarsInscricaoINSS = vData
End Property

Public Property Get sInscricaoINSS() As String
    sInscricaoINSS = mvarsInscricaoINSS
End Property

Public Property Let sContaDespesa(ByVal vData As String)
    mvarsContaDespesa = vData
End Property

Public Property Get sContaDespesa() As String
    sContaDespesa = mvarsContaDespesa
End Property

Public Property Let iTipoCobranca(ByVal vData As Integer)
    mvariTipoCobranca = vData
End Property

Public Property Get iTipoCobranca() As Integer
    iTipoCobranca = mvariTipoCobranca
End Property

Public Property Let iGeraCredICMS(ByVal vData As Integer)
    mvariGeraCredICMS = vData
End Property

Public Property Get iGeraCredICMS() As Integer
    iGeraCredICMS = mvariGeraCredICMS
End Property

Private Sub Class_Initialize()
    mvariIEIsento = -1
    mvariIENaoContrib = -1

End Sub

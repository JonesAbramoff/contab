VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAdmExtFinArqsLidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsNomeArq As String
Private mvardtDataImportacao As Date
Private mvardHoraImportacao As Double
Private mvarsUsuario As String
Private mvardtDataAtualizado As Date
Private mvariFilialEmpresa As Integer
Private mvariBandeira As Integer
Private mvariNaoAtualizar As Integer
Private mvarcolMov As New Collection

'########################################
'AUXILIARES
Private mvariQtdDep As Integer
Private mvariQtdParc As Integer
Private mvariQtdParcEnc As Integer
Private mvardVlrDep As Double
Private mvardVlrParc As Double
Private mvardVlrParcEnc As Double
Private mvariTotalReg As Integer

Public Property Let iTotalReg(ByVal vData As Integer)
    mvariTotalReg = vData
End Property

Public Property Get iTotalReg() As Integer
    iTotalReg = mvariTotalReg
End Property

Public Property Let dVlrDep(ByVal vData As Double)
    mvardVlrDep = vData
End Property

Public Property Get dVlrDep() As Double
    dVlrDep = mvardVlrDep
End Property

Public Property Let dVlrParc(ByVal vData As Double)
    mvardVlrParc = vData
End Property

Public Property Get dVlrParc() As Double
    dVlrParc = mvardVlrParc
End Property

Public Property Let dVlrParcEnc(ByVal vData As Double)
    mvardVlrParcEnc = vData
End Property

Public Property Get dVlrParcEnc() As Double
    dVlrParcEnc = mvardVlrParcEnc
End Property

Public Property Let iQtdDep(ByVal vData As Integer)
    mvariQtdDep = vData
End Property

Public Property Get iQtdDep() As Integer
    iQtdDep = mvariQtdDep
End Property

Public Property Let iQtdParc(ByVal vData As Integer)
    mvariQtdParc = vData
End Property

Public Property Get iQtdParc() As Integer
    iQtdParc = mvariQtdParc
End Property

Public Property Let iQtdParcEnc(ByVal vData As Integer)
    mvariQtdParcEnc = vData
End Property

Public Property Get iQtdParcEnc() As Integer
    iQtdParcEnc = mvariQtdParcEnc
End Property
'AUXILIARES
'########################################

Public Property Set colMov(ByVal vData As Collection)
    Set mvarcolMov = vData
End Property

Public Property Get colMov() As Collection
    Set colMov = mvarcolMov
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sNomeArq(ByVal vData As String)
    mvarsNomeArq = vData
End Property

Public Property Get sNomeArq() As String
    sNomeArq = mvarsNomeArq
End Property

Public Property Let dtDataImportacao(ByVal vData As Date)
    mvardtDataImportacao = vData
End Property

Public Property Get dtDataImportacao() As Date
    dtDataImportacao = mvardtDataImportacao
End Property

Public Property Let dHoraImportacao(ByVal vData As Double)
    mvardHoraImportacao = vData
End Property

Public Property Get dHoraImportacao() As Double
    dHoraImportacao = mvardHoraImportacao
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let dtDataAtualizado(ByVal vData As Date)
    mvardtDataAtualizado = vData
End Property

Public Property Get dtDataAtualizado() As Date
    dtDataAtualizado = mvardtDataAtualizado
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iBandeira(ByVal vData As Integer)
    mvariBandeira = vData
End Property

Public Property Get iBandeira() As Integer
    iBandeira = mvariBandeira
End Property

Public Property Let iNaoAtualizar(ByVal vData As Integer)
    mvariNaoAtualizar = vData
End Property

Public Property Get iNaoAtualizar() As Integer
    iNaoAtualizar = mvariNaoAtualizar
End Property

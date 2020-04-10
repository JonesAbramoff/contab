VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqImportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntArq As Long
Private mvariTipoArq As Integer
Private mvarsNomeArquivo As String
Private mvardtDataImportacao As Date
Private mvardHoraImportacao As Double
Private mvardtDataAtualizacao As Date
Private mvardHoraAtualizacao As Double
Private mvarsUsuario As String
Private mvariManter As Integer

Public Property Let iManter(ByVal vData As Integer)
    mvariManter = vData
End Property

Public Property Get iManter() As Integer
    iManter = mvariManter
End Property

Public Property Let lNumIntArq(ByVal vData As Long)
    mvarlNumIntArq = vData
End Property

Public Property Get lNumIntArq() As Long
    lNumIntArq = mvarlNumIntArq
End Property

Public Property Let iTipoArq(ByVal vData As Integer)
    mvariTipoArq = vData
End Property

Public Property Get iTipoArq() As Integer
    iTipoArq = mvariTipoArq
End Property

Public Property Let sNomeArquivo(ByVal vData As String)
    mvarsNomeArquivo = vData
End Property

Public Property Get sNomeArquivo() As String
    sNomeArquivo = mvarsNomeArquivo
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

Public Property Let dtDataAtualizacao(ByVal vData As Date)
    mvardtDataAtualizacao = vData
End Property

Public Property Get dtDataAtualizacao() As Date
    dtDataAtualizacao = mvardtDataAtualizacao
End Property

Public Property Let dHoraAtualizacao(ByVal vData As Double)
    mvardHoraAtualizacao = vData
End Property

Public Property Get dHoraAtualizacao() As Double
    dHoraAtualizacao = mvardHoraAtualizacao
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

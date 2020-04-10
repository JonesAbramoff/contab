VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqExportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntArq As Long
Private mvarsNomeArquivo As String
Private mvariTipoArq As Integer
Private mvardtDataExportacao As Date
Private mvardHoraExportacao As Double
Private mvarsUsuario As String

Public Property Let lNumIntArq(ByVal vData As Long)
    mvarlNumIntArq = vData
End Property

Public Property Get lNumIntArq() As Long
    lNumIntArq = mvarlNumIntArq
End Property

Public Property Let sNomeArquivo(ByVal vData As String)
    mvarsNomeArquivo = vData
End Property

Public Property Get sNomeArquivo() As String
    sNomeArquivo = mvarsNomeArquivo
End Property

Public Property Let iTipoArq(ByVal vData As Integer)
    mvariTipoArq = vData
End Property

Public Property Get iTipoArq() As Integer
    iTipoArq = mvariTipoArq
End Property

Public Property Let dtDataExportacao(ByVal vData As Date)
    mvardtDataExportacao = vData
End Property

Public Property Get dtDataExportacao() As Date
    dtDataExportacao = mvardtDataExportacao
End Property

Public Property Let dHoraExportacao(ByVal vData As Double)
    mvardHoraExportacao = vData
End Property

Public Property Get dHoraExportacao() As Double
    dHoraExportacao = mvardHoraExportacao
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

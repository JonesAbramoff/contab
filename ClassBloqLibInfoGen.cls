VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBloqLibInfoGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvariSeqBloqueio As Integer
Private mvarsNomeRedCliForn As String
Private mvardtData As Date
Private mvardValor As Double
Private mvariTipoDeBloqueio As Integer
Private mvarsNomeRedTipoBloq As String
Private mvarsUsuario As String
Private mvardtDataBloqueio As Date
Private mvariMarcado As Integer
Private mvarsObservacao As String

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let iMarcado(ByVal vData As Integer)
    mvariMarcado = vData
End Property

Public Property Get iMarcado() As Integer
    iMarcado = mvariMarcado
End Property

Public Property Let dtDataBloqueio(ByVal vData As Date)
    mvardtDataBloqueio = vData
End Property

Public Property Get dtDataBloqueio() As Date
    dtDataBloqueio = mvardtDataBloqueio
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let sNomeRedTipoBloq(ByVal vData As String)
    mvarsNomeRedTipoBloq = vData
End Property

Public Property Get sNomeRedTipoBloq() As String
    sNomeRedTipoBloq = mvarsNomeRedTipoBloq
End Property

Public Property Let iTipoDeBloqueio(ByVal vData As Integer)
    mvariTipoDeBloqueio = vData
End Property

Public Property Get iTipoDeBloqueio() As Integer
    iTipoDeBloqueio = mvariTipoDeBloqueio
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sNomeRedCliForn(ByVal vData As String)
    mvarsNomeRedCliForn = vData
End Property

Public Property Get sNomeRedCliForn() As String
    sNomeRedCliForn = mvarsNomeRedCliForn
End Property

Public Property Let iSeqBloqueio(ByVal vData As Integer)
    mvariSeqBloqueio = vData
End Property

Public Property Get iSeqBloqueio() As Integer
    iSeqBloqueio = mvariSeqBloqueio
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property




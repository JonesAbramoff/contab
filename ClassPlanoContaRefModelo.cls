VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPlanoContaRefModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodigo As Long
Private mvarsDescricao As String
Private mvariOficial As Integer
Private mvariAnoVigencia As Integer
Private mvarsUsuarioCad As String
Private mvardtDataCad As Date
Private mvardHoraCad As Double
Private mvarsUsuarioAlt As String
Private mvardtDataAlt As Date
Private mvardHoraAlt As Double
Private mvarcolContas As New Collection
Private mvariTipo As Integer

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Set colContas(ByVal vData As Object)
    Set mvarcolContas = vData
End Property

Public Property Get colContas() As Collection
    Set colContas = mvarcolContas
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iOficial(ByVal vData As Integer)
    mvariOficial = vData
End Property

Public Property Get iOficial() As Integer
    iOficial = mvariOficial
End Property

Public Property Let iAnoVigencia(ByVal vData As Integer)
    mvariAnoVigencia = vData
End Property

Public Property Get iAnoVigencia() As Integer
    iAnoVigencia = mvariAnoVigencia
End Property

Public Property Let sUsuarioCad(ByVal vData As String)
    mvarsUsuarioCad = vData
End Property

Public Property Get sUsuarioCad() As String
    sUsuarioCad = mvarsUsuarioCad
End Property

Public Property Let dtDataCad(ByVal vData As Date)
    mvardtDataCad = vData
End Property

Public Property Get dtDataCad() As Date
    dtDataCad = mvardtDataCad
End Property

Public Property Let dHoraCad(ByVal vData As Double)
    mvardHoraCad = vData
End Property

Public Property Get dHoraCad() As Double
    dHoraCad = mvardHoraCad
End Property

Public Property Let sUsuarioAlt(ByVal vData As String)
    mvarsUsuarioAlt = vData
End Property

Public Property Get sUsuarioAlt() As String
    sUsuarioAlt = mvarsUsuarioAlt
End Property

Public Property Let dtDataAlt(ByVal vData As Date)
    mvardtDataAlt = vData
End Property

Public Property Get dtDataAlt() As Date
    dtDataAlt = mvardtDataAlt
End Property

Public Property Let dHoraAlt(ByVal vData As Double)
    mvardHoraAlt = vData
End Property

Public Property Get dHoraAlt() As Double
    dHoraAlt = mvardHoraAlt
End Property

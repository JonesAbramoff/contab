VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPTitulosExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsUsuario As String
Private mvardtData As Date
Private mvardHora As Double
Private mvariTipoDocOrigem As Integer
Private mvarlNumIntDocOrigem As Long
Private mvarlNumTitulo As Long
Private mvariExcluido As Integer
Private mvariExportado As Integer
Private mvarsMotivo As String
Private mvariTemQueContabilizar As Integer
Private mvardValorAporteCred As Double
Private mvardValorComissao As Double
Private mvardValorAporte As Double
Private mvardValorCreditos As Double
Private mvardValorDebitos As Double
Private mvardValorTarifa As Double
Private mvardValorDeducoes As Double
Private mvarcolFatAporteCred As New Collection

Public Property Set colFatAporteCred(ByVal vData As Collection)
    Set mvarcolFatAporteCred = vData
End Property

Public Property Get colFatAporteCred() As Collection
    Set colFatAporteCred = mvarcolFatAporteCred
End Property

Public Property Let dValorAporte(ByVal vData As Double)
    mvardValorAporte = vData
End Property

Public Property Get dValorAporte() As Double
    dValorAporte = mvardValorAporte
End Property

Public Property Let dValorComissao(ByVal vData As Double)
    mvardValorComissao = vData
End Property

Public Property Get dValorComissao() As Double
    dValorComissao = mvardValorComissao
End Property

Public Property Let dValorDebitos(ByVal vData As Double)
    mvardValorDebitos = vData
End Property

Public Property Get dValorDebitos() As Double
    dValorDebitos = mvardValorDebitos
End Property

Public Property Let dValorCreditos(ByVal vData As Double)
    mvardValorCreditos = vData
End Property

Public Property Get dValorCreditos() As Double
    dValorCreditos = mvardValorCreditos
End Property

Public Property Let dValorDeducoes(ByVal vData As Double)
    mvardValorDeducoes = vData
End Property

Public Property Get dValorDeducoes() As Double
    dValorDeducoes = mvardValorDeducoes
End Property

Public Property Let dValorTarifa(ByVal vData As Double)
    mvardValorTarifa = vData
End Property

Public Property Get dValorTarifa() As Double
    dValorTarifa = mvardValorTarifa
End Property

Public Property Let dValorAporteCred(ByVal vData As Double)
    mvardValorAporteCred = vData
End Property

Public Property Get dValorAporteCred() As Double
    dValorAporteCred = mvardValorAporteCred
End Property

Public Property Let iTemQueContabilizar(ByVal vData As Integer)
    mvariTemQueContabilizar = vData
End Property

Public Property Get iTemQueContabilizar() As Integer
    iTemQueContabilizar = mvariTemQueContabilizar
End Property

Public Property Let sMotivo(ByVal vData As String)
    mvarsMotivo = vData
End Property

Public Property Get sMotivo() As String
    sMotivo = mvarsMotivo
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dHora(ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora() As Double
    dHora = mvardHora
End Property

Public Property Let iTipoDocOrigem(ByVal vData As Integer)
    mvariTipoDocOrigem = vData
End Property

Public Property Get iTipoDocOrigem() As Integer
    iTipoDocOrigem = mvariTipoDocOrigem
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property

Public Property Let iExcluido(ByVal vData As Integer)
    mvariExcluido = vData
End Property

Public Property Get iExcluido() As Integer
    iExcluido = mvariExcluido
End Property

Public Property Let iExportado(ByVal vData As Integer)
    mvariExportado = vData
End Property

Public Property Get iExportado() As Integer
    iExportado = mvariExportado
End Property

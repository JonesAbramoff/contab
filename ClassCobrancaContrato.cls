VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCobrancaContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarlNumIntDoc As Long
Private mvardtDataRefIni As Date
Private mvardtDataRefFim As Date
Private mvardtDataCobrIni As Date
Private mvardtDataCobrFim As Date
Private mvardtDataGeracao As Date
Private mvarsUsuario As String

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let dtDataGeracao(ByVal vData As Date)
    mvardtDataGeracao = vData
End Property

Public Property Get dtDataGeracao() As Date
    dtDataGeracao = mvardtDataGeracao
End Property

Public Property Let dtDataCobrFim(ByVal vData As Date)
    mvardtDataCobrFim = vData
End Property

Public Property Get dtDataCobrFim() As Date
    dtDataCobrFim = mvardtDataCobrFim
End Property

Public Property Let dtDataCobrIni(ByVal vData As Date)
    mvardtDataCobrIni = vData
End Property

Public Property Get dtDataCobrIni() As Date
    dtDataCobrIni = mvardtDataCobrIni
End Property

Public Property Let dtDataRefFim(ByVal vData As Date)
    mvardtDataRefFim = vData
End Property

Public Property Get dtDataRefFim() As Date
    dtDataRefFim = mvardtDataRefFim
End Property

Public Property Let dtDataRefIni(ByVal vData As Date)
    mvardtDataRefIni = vData
End Property

Public Property Get dtDataRefIni() As Date
    dtDataRefIni = mvardtDataRefIni
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDEInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsNumero As String
Private mvardtData As Date
Private mvariFilialEmpresa As Integer
Private mvarsDescricao As String
Private mvariTipoDoc As Integer
Private mvariNatureza As Integer
Private mvarsNumConhEmbarque As String
Private mvarsUFEmbarque As String
Private mvarsLocalEmbarque As String
Private mvardtDataConhEmbarque As Date
Private mvariTipoConhEmbarque As Integer
Private mvariCodPais As Integer
Private mvardtDataAverbacao As Date
Private mvarcolRE As New Collection

'AUXILIAR
Private mvarsNumRegistro As String
Private mvardtDataRegistro As Date

Public Property Let sNumRegistro(ByVal vData As String)
    mvarsNumRegistro = vData
End Property

Public Property Get sNumRegistro() As String
    sNumRegistro = mvarsNumRegistro
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property
'AUXILIAR

Public Property Set colRE(ByVal vData As Object)
    Set mvarcolRE = vData
End Property

Public Property Get colRE() As Collection
    Set colRE = mvarcolRE
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sNumero(ByVal vData As String)
    mvarsNumero = vData
End Property

Public Property Get sNumero() As String
    sNumero = mvarsNumero
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let iNatureza(ByVal vData As Integer)
    mvariNatureza = vData
End Property

Public Property Get iNatureza() As Integer
    iNatureza = mvariNatureza
End Property

Public Property Let sNumConhEmbarque(ByVal vData As String)
    mvarsNumConhEmbarque = vData
End Property

Public Property Get sNumConhEmbarque() As String
    sNumConhEmbarque = mvarsNumConhEmbarque
End Property

Public Property Let sUFEmbarque(ByVal vData As String)
    mvarsUFEmbarque = vData
End Property

Public Property Get sUFEmbarque() As String
    sUFEmbarque = mvarsUFEmbarque
End Property

Public Property Let sLocalEmbarque(ByVal vData As String)
    mvarsLocalEmbarque = vData
End Property

Public Property Get sLocalEmbarque() As String
    sLocalEmbarque = mvarsLocalEmbarque
End Property

Public Property Let dtDataConhEmbarque(ByVal vData As Date)
    mvardtDataConhEmbarque = vData
End Property

Public Property Get dtDataConhEmbarque() As Date
    dtDataConhEmbarque = mvardtDataConhEmbarque
End Property

Public Property Let iTipoConhEmbarque(ByVal vData As Integer)
    mvariTipoConhEmbarque = vData
End Property

Public Property Get iTipoConhEmbarque() As Integer
    iTipoConhEmbarque = mvariTipoConhEmbarque
End Property

Public Property Let iCodPais(ByVal vData As Integer)
    mvariCodPais = vData
End Property

Public Property Get iCodPais() As Integer
    iCodPais = mvariCodPais
End Property

Public Property Let dtDataAverbacao(ByVal vData As Date)
    mvardtDataAverbacao = vData
End Property

Public Property Get dtDataAverbacao() As Date
    dtDataAverbacao = mvardtDataAverbacao
End Property

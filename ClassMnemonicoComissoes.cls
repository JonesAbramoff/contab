VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMnemonicoComissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarsMnemonico As String
Private mvarsDescricao As String
Private mvarsGrid As String
Private mvarsNomeBrowser As String
Private mvarsProjetoBrowser As String
Private mvarsClasseBrowser As String
Private mvarsPropertyBrowser As String
Private mvariTipo As Integer
Private mvariNumParam As Integer
Private mvariParam1 As Integer
Private mvariParam2 As Integer
Private mvariParam3 As Integer

Public Property Let iParam3(ByVal vData As Integer)
    mvariParam3 = vData
End Property

Public Property Get iParam3() As Integer
    iParam3 = mvariParam3
End Property

Public Property Let iParam2(ByVal vData As Integer)
    mvariParam2 = vData
End Property

Public Property Get iParam2() As Integer
    iParam2 = mvariParam2
End Property

Public Property Let iParam1(ByVal vData As Integer)
    mvariParam1 = vData
End Property

Public Property Get iParam1() As Integer
    iParam1 = mvariParam1
End Property

Public Property Let iNumParam(ByVal vData As Integer)
    mvariNumParam = vData
End Property

Public Property Get iNumParam() As Integer
    iNumParam = mvariNumParam
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sGrid(ByVal vData As String)
    mvarsGrid = vData
End Property

Public Property Get sGrid() As String
    sGrid = mvarsGrid
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sMnemonico(ByVal vData As String)
    mvarsMnemonico = vData
End Property

Public Property Get sMnemonico() As String
    sMnemonico = mvarsMnemonico
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property


Public Property Let sNomeBrowser(ByVal vData As String)
    mvarsNomeBrowser = vData
End Property

Public Property Get sNomeBrowser() As String
    sNomeBrowser = mvarsNomeBrowser
End Property

Public Property Let sProjetoBrowser(ByVal vData As String)
    mvarsProjetoBrowser = vData
End Property

Public Property Get sProjetoBrowser() As String
    sProjetoBrowser = mvarsProjetoBrowser
End Property

Public Property Let sClasseBrowser(ByVal vData As String)
    mvarsClasseBrowser = vData
End Property

Public Property Get sClasseBrowser() As String
    sClasseBrowser = mvarsClasseBrowser
End Property

Public Property Let sPropertyBrowser(ByVal vData As String)
    mvarsPropertyBrowser = vData
End Property

Public Property Get sPropertyBrowser() As String
    sPropertyBrowser = mvarsPropertyBrowser
End Property


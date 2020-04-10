VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCamposCustomizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocOrigem As Long
Private mvariTipoNumIntDocOrigem As Integer
Private mvardtData(1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES) As Date
Private mvarsTexto(1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES) As String
Private mvarlNumero(1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES) As Long
Private mvardValor(1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES) As Double

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let iTipoNumIntDocOrigem(ByVal vData As Integer)
    mvariTipoNumIntDocOrigem = vData
End Property

Public Property Get iTipoNumIntDocOrigem() As Integer
    iTipoNumIntDocOrigem = mvariTipoNumIntDocOrigem
End Property

Public Property Let dtData(ByVal iIndice As Integer, ByVal vData As Date)
    mvardtData(iIndice) = vData
End Property

Public Property Get dtData(ByVal iIndice As Integer) As Date
    dtData = mvardtData(iIndice)
End Property

Public Property Let sTexto(ByVal iIndice As Integer, ByVal vData As String)
    mvarsTexto(iIndice) = vData
End Property

Public Property Get sTexto(ByVal iIndice As Integer) As String
    sTexto = mvarsTexto(iIndice)
End Property

Public Property Let lNumero(ByVal iIndice As Integer, ByVal vData As Long)
    mvarlNumero(iIndice) = vData
End Property

Public Property Get lNumero(ByVal iIndice As Integer) As Long
    lNumero = mvarlNumero(iIndice)
End Property

Public Property Let dValor(ByVal iIndice As Integer, ByVal vData As Double)
    mvardValor(iIndice) = vData
End Property

Public Property Get dValor(ByVal iIndice As Integer) As Double
    dValor = mvardValor(iIndice)
End Property

Private Sub Class_Initialize()

Dim iIndice As Integer

    For iIndice = 1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES
        mvardtData(iIndice) = DATA_NULA
    Next
    
End Sub

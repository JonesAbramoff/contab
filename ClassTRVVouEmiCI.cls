VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVVouEmiCI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsTipVou As String
Private mvarsSerie As String
Private mvarlNumVou As Long
Private mvarlFornEmissor As Long
Private mvardPercCI As Double
Private mvardPercReal As Double
Private mvardPercCIAnt As Double

Public Property Let dPercCIAnt(ByVal vData As Double)
    mvardPercCIAnt = vData
End Property

Public Property Get dPercCIAnt() As Double
    dPercCIAnt = mvardPercCIAnt
End Property

Public Property Let sTipVou(ByVal vData As String)
    mvarsTipVou = vData
End Property

Public Property Get sTipVou() As String
    sTipVou = mvarsTipVou
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let lNumVou(ByVal vData As Long)
    mvarlNumVou = vData
End Property

Public Property Get lNumVou() As Long
    lNumVou = mvarlNumVou
End Property

Public Property Let lFornEmissor(ByVal vData As Long)
    mvarlFornEmissor = vData
End Property

Public Property Get lFornEmissor() As Long
    lFornEmissor = mvarlFornEmissor
End Property

Public Property Let dPercCI(ByVal vData As Double)
    mvardPercCI = vData
End Property

Public Property Get dPercCI() As Double
    dPercCI = mvardPercCI
End Property

Public Property Let dPercReal(ByVal vData As Double)
    mvardPercReal = vData
End Property

Public Property Get dPercReal() As Double
    dPercReal = mvardPercReal
End Property

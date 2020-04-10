VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCodTribMun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCidade As Long
Private mvarsProduto As String
Private mvarsCodTribMun As String
Private mvardAliquota As Double

Public Property Let dAliquota(ByVal vData As Double)
   mvardAliquota = vData
End Property

Public Property Get dAliquota() As Double
    dAliquota = mvardAliquota
End Property

Public Property Let lCidade(ByVal vData As Long)
    mvarlCidade = vData
End Property

Public Property Get lCidade() As Long
    lCidade = mvarlCidade
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sCodTribMun(ByVal vData As String)
    mvarsCodTribMun = vData
End Property

Public Property Get sCodTribMun() As String
    sCodTribMun = mvarsCodTribMun
End Property

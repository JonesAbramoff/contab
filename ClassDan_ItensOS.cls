VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassDan_ItensOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsOS AS String
Private mvariItem AS Integer
Private mvarsProduto AS String
Private mvardQuantidade AS Double

Public Property Let sOS (ByVal vData As String)
    mvarsOS = vData
End Property

Public Property Get sOS () AS String
    sOS= mvarsOS
End Property

Public Property Let iItem (ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem () AS Integer
    iItem= mvariItem
End Property

Public Property Let sProduto (ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto () AS String
    sProduto= mvarsProduto
End Property

Public Property Let dQuantidade (ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade () AS Double
    dQuantidade= mvardQuantidade
End Property

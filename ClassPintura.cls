VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassPintura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariCodigo AS Integer
Private mvarsDescricao AS String

Public Property Let iCodigo (ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo () AS Integer
    iCodigo= mvariCodigo
End Property

Public Property Let sDescricao (ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao () AS String
    sDescricao= mvarsDescricao
End Property

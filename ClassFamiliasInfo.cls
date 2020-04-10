VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassFamiliasInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodFamilia AS Long
Private mvariSeq AS Integer
Private mvariCodInfo AS Integer
Private mvariValor AS Integer

Public Property Let lCodFamilia (ByVal vData As Long)
    mvarlCodFamilia = vData
End Property

Public Property Get lCodFamilia () AS Long
    lCodFamilia= mvarlCodFamilia
End Property

Public Property Let iSeq (ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq () AS Integer
    iSeq= mvariSeq
End Property

Public Property Let iCodInfo (ByVal vData As Integer)
    mvariCodInfo = vData
End Property

Public Property Get iCodInfo () AS Integer
    iCodInfo= mvariCodInfo
End Property

Public Property Let iValor (ByVal vData As Integer)
    mvariValor = vData
End Property

Public Property Get iValor () AS Integer
    iValor= mvariValor
End Property

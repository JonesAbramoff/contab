VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassTRVOcrCasoArqImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntArq AS Long
Private mvardtData AS Date
Private mvardHora AS Double
Private mvarsNomeArq AS String

Public Property Let lNumIntArq (ByVal vData As Long)
    mvarlNumIntArq = vData
End Property

Public Property Get lNumIntArq () AS Long
    lNumIntArq= mvarlNumIntArq
End Property

Public Property Let dtData (ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData () AS Date
    dtData= mvardtData
End Property

Public Property Let dHora (ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora () AS Double
    dHora= mvardHora
End Property

Public Property Let sNomeArq (ByVal vData As String)
    mvarsNomeArq = vData
End Property

Public Property Get sNomeArq () AS String
    sNomeArq= mvarsNomeArq
End Property

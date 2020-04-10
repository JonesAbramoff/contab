VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInfoAdicDocItemDetExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariSeq As Integer
Private mvarsNumDrawback As String
Private mvarsNumRegistExport As String
Private mvarsChvNFe As String
Private mvardQuantExport As Double

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let sNumDrawback(ByVal vData As String)
    mvarsNumDrawback = vData
End Property

Public Property Get sNumDrawback() As String
    sNumDrawback = mvarsNumDrawback
End Property

Public Property Let sNumRegistExport(ByVal vData As String)
    mvarsNumRegistExport = vData
End Property

Public Property Get sNumRegistExport() As String
    sNumRegistExport = mvarsNumRegistExport
End Property

Public Property Let sChvNFe(ByVal vData As String)
    mvarsChvNFe = vData
End Property

Public Property Get sChvNFe() As String
    sChvNFe = mvarsChvNFe
End Property

Public Property Let dQuantExport(ByVal vData As Double)
    mvardQuantExport = vData
End Property

Public Property Get dQuantExport() As Double
    dQuantExport = mvardQuantExport
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInfoAdicExportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariTipoDoc As Integer
Private mvarlNumIntDoc As Long
Private mvarsUFEmbarque As String
Private mvarsLocalEmbarque As String
Private mvarlNumIntDE As Long
Private mvarsNumRE As String

Public Property Let sNumRE(ByVal vData As String)
    mvarsNumRE = vData
End Property

Public Property Get sNumRE() As String
    sNumRE = mvarsNumRE
End Property

Public Property Let lNumIntDE(ByVal vData As Long)
    mvarlNumIntDE = vData
End Property

Public Property Get lNumIntDE() As Long
    lNumIntDE = mvarlNumIntDE
End Property


Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
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

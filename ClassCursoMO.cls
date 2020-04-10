VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCursoMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocCurso As Long
Private mvariCodMO As Integer
Private mvarsAvaliacao As String
Private mvariAprovado As Integer

Public Property Let iAprovado(ByVal vData As Integer)
    mvariAprovado = vData
End Property

Public Property Get iAprovado() As Integer
    iAprovado = mvariAprovado
End Property

Public Property Let lNumIntDocCurso(ByVal vData As Long)
    mvarlNumIntDocCurso = vData
End Property

Public Property Get lNumIntDocCurso() As Long
    lNumIntDocCurso = mvarlNumIntDocCurso
End Property

Public Property Let iCodMO(ByVal vData As Integer)
    mvariCodMO = vData
End Property

Public Property Get iCodMO() As Integer
    iCodMO = mvariCodMO
End Property

Public Property Let sAvaliacao(ByVal vData As String)
    mvarsAvaliacao = vData
End Property

Public Property Get sAvaliacao() As String
    sAvaliacao = mvarsAvaliacao
End Property


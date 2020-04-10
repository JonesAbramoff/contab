VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCursoCertificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocCurso As Long
Private mvarlCodCertificado As Long

Public Property Let lNumIntDocCurso(ByVal vData As Long)
    mvarlNumIntDocCurso = vData
End Property

Public Property Get lNumIntDocCurso() As Long
    lNumIntDocCurso = mvarlNumIntDocCurso
End Property

Public Property Let lCodCertificado(ByVal vData As Long)
    mvarlCodCertificado = vData
End Property

Public Property Get lCodCertificado() As Long
    lCodCertificado = mvarlCodCertificado
End Property

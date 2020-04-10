VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMvLctoCtaRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel As Long
Private mvariSeqCtaPer As Integer
Private mvariSeqCtaDia As Integer
Private mvariSeqCtaLcto As Integer
Private mvariFilialEmpresaLcto As Integer
Private mvarsOrigemLcto As String
Private mvarlDocLcto As Long
Private mvariSeqLcto As Integer
Private mvardtDataLcto As Date
Private mvardValorLcto As Double
Private mvarsHistoricoLcto As String

Public Property Let sHistoricoLcto(ByVal vData As String)
    mvarsHistoricoLcto = vData
End Property

Public Property Get sHistoricoLcto() As String
    sHistoricoLcto = mvarsHistoricoLcto
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let iSeqCtaPer(ByVal vData As Integer)
    mvariSeqCtaPer = vData
End Property

Public Property Get iSeqCtaPer() As Integer
    iSeqCtaPer = mvariSeqCtaPer
End Property

Public Property Let iSeqCtaDia(ByVal vData As Integer)
    mvariSeqCtaDia = vData
End Property

Public Property Get iSeqCtaDia() As Integer
    iSeqCtaDia = mvariSeqCtaDia
End Property

Public Property Let iSeqCtaLcto(ByVal vData As Integer)
    mvariSeqCtaLcto = vData
End Property

Public Property Get iSeqCtaLcto() As Integer
    iSeqCtaLcto = mvariSeqCtaLcto
End Property

Public Property Let iFilialEmpresaLcto(ByVal vData As Integer)
    mvariFilialEmpresaLcto = vData
End Property

Public Property Get iFilialEmpresaLcto() As Integer
    iFilialEmpresaLcto = mvariFilialEmpresaLcto
End Property

Public Property Let sOrigemLcto(ByVal vData As String)
    mvarsOrigemLcto = vData
End Property

Public Property Get sOrigemLcto() As String
    sOrigemLcto = mvarsOrigemLcto
End Property

Public Property Let lDocLcto(ByVal vData As Long)
    mvarlDocLcto = vData
End Property

Public Property Get lDocLcto() As Long
    lDocLcto = mvarlDocLcto
End Property

Public Property Let iSeqLcto(ByVal vData As Integer)
    mvariSeqLcto = vData
End Property

Public Property Get iSeqLcto() As Integer
    iSeqLcto = mvariSeqLcto
End Property

Public Property Let dtDataLcto(ByVal vData As Date)
    mvardtDataLcto = vData
End Property

Public Property Get dtDataLcto() As Date
    dtDataLcto = mvardtDataLcto
End Property

Public Property Let dValorLcto(ByVal vData As Double)
    mvardValorLcto = vData
End Property

Public Property Get dValorLcto() As Double
    dValorLcto = mvardValorLcto
End Property

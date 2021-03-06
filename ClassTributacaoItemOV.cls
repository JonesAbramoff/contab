VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributacaoItemOV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'??? campos que deverao ser excluidos
'Private mvariFilialEmpresa As Integer
'Private mvarlCodPedido As Long
'Private mvariItem As Integer
'??? fim
Private mvarsNaturezaOp As String
Private mvariNaturezaOpManual As Integer
Private mvariTipoTributacao As Integer
Private mvariTipoTributacaoManual As Integer
Private mvariIPITipo As Integer
Private mvariIPITipoManual As Integer
Private mvardIPIBaseCalculo As Double
Private mvariIPIBaseManual As Integer
Private mvardIPIPercRedBase As Double
Private mvariIPIPercRedBaseManual As Integer
Private mvardIPIAliquota As Double
Private mvariIPIAliquotaManual As Integer
Private mvardIPIValor As Double
Private mvariIPIValorManual As Integer
Private mvariICMSTipo As Integer
Private mvariICMSTipoManual As Integer
Private mvardICMSBase As Double
Private mvariICMSBaseManual As Integer
Private mvardICMSPercRedBase As Double
Private mvariICMSPercRedBaseManual As Integer
Private mvardICMSAliquota As Double
Private mvariICMSAliquotaManual As Integer
Private mvardICMSValor As Double
Private mvariICMSValorManual As Integer
Private mvardICMSSubstBase As Double
Private mvariICMSSubstBaseManual As Integer
Private mvardICMSSubstAliquota As Double
Private mvariICMSSubstAliquotaManual As Integer
Private mvardICMSSubstValor As Double
Private mvariICMSSubstValorManual As Integer

Private mvarlNumIntDoc As Long

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

'Public Property Let iFilialEmpresa(ByVal vData As Integer)
'    mvariFilialEmpresa = vData
'End Property
'
'Public Property Get iFilialEmpresa() As Integer
'    iFilialEmpresa = mvariFilialEmpresa
'End Property
'
'Public Property Let lCodPedido(ByVal vData As Long)
'    mvarlCodPedido = vData
'End Property
'
'Public Property Get lCodPedido() As Long
'    lCodPedido = mvarlCodPedido
'End Property
'
'Public Property Let iItem(ByVal vData As Integer)
'    mvariItem = vData
'End Property
'
'Public Property Get iItem() As Integer
'    iItem = mvariItem
'End Property
'
Public Property Let iICMSSubstAliquotaManual(ByVal vData As Integer)
    mvariICMSSubstAliquotaManual = vData
End Property

Public Property Get iICMSSubstAliquotaManual() As Integer
    iICMSSubstAliquotaManual = mvariICMSSubstAliquotaManual
End Property

Public Property Let iICMSSubstValorManual(ByVal vData As Integer)
    mvariICMSSubstValorManual = vData
End Property

Public Property Get iICMSSubstValorManual() As Integer
    iICMSSubstValorManual = mvariICMSSubstValorManual
End Property

Public Property Let dICMSSubstValor(ByVal vData As Double)
    mvardICMSSubstValor = vData
End Property

Public Property Get dICMSSubstValor() As Double
    dICMSSubstValor = mvardICMSSubstValor
End Property

Public Property Let dICMSSubstAliquota(ByVal vData As Double)
    mvardICMSSubstAliquota = vData
End Property

Public Property Get dICMSSubstAliquota() As Double
    dICMSSubstAliquota = mvardICMSSubstAliquota
End Property

Public Property Let iICMSSubstBaseManual(ByVal vData As Integer)
    mvariICMSSubstBaseManual = vData
End Property

Public Property Get iICMSSubstBaseManual() As Integer
    iICMSSubstBaseManual = mvariICMSSubstBaseManual
End Property

Public Property Let dICMSSubstBase(ByVal vData As Double)
    mvardICMSSubstBase = vData
End Property

Public Property Get dICMSSubstBase() As Double
    dICMSSubstBase = mvardICMSSubstBase
End Property

Public Property Let iICMSValorManual(ByVal vData As Integer)
    mvariICMSValorManual = vData
End Property

Public Property Get iICMSValorManual() As Integer
    iICMSValorManual = mvariICMSValorManual
End Property

Public Property Let dICMSValor(ByVal vData As Double)
    mvardICMSValor = vData
End Property

Public Property Get dICMSValor() As Double
    dICMSValor = mvardICMSValor
End Property

Public Property Let iICMSAliquotaManual(ByVal vData As Integer)
    mvariICMSAliquotaManual = vData
End Property

Public Property Get iICMSAliquotaManual() As Integer
    iICMSAliquotaManual = mvariICMSAliquotaManual
End Property

Public Property Let dICMSAliquota(ByVal vData As Double)
    mvardICMSAliquota = vData
End Property

Public Property Get dICMSAliquota() As Double
    dICMSAliquota = mvardICMSAliquota
End Property

Public Property Let iICMSPercRedBaseManual(ByVal vData As Integer)
    mvariICMSPercRedBaseManual = vData
End Property

Public Property Get iICMSPercRedBaseManual() As Integer
    iICMSPercRedBaseManual = mvariICMSPercRedBaseManual
End Property

Public Property Let dICMSPercRedBase(ByVal vData As Double)
    mvardICMSPercRedBase = vData
End Property

Public Property Get dICMSPercRedBase() As Double
    dICMSPercRedBase = mvardICMSPercRedBase
End Property

Public Property Let iICMSBaseManual(ByVal vData As Integer)
    mvariICMSBaseManual = vData
End Property

Public Property Get iICMSBaseManual() As Integer
    iICMSBaseManual = mvariICMSBaseManual
End Property

Public Property Let dICMSBase(ByVal vData As Double)
    mvardICMSBase = vData
End Property

Public Property Get dICMSBase() As Double
    dICMSBase = mvardICMSBase
End Property

Public Property Let iICMSTipoManual(ByVal vData As Integer)
    mvariICMSTipoManual = vData
End Property

Public Property Get iICMSTipoManual() As Integer
    iICMSTipoManual = mvariICMSTipoManual
End Property

Public Property Let iICMSTipo(ByVal vData As Integer)
    mvariICMSTipo = vData
End Property

Public Property Get iICMSTipo() As Integer
    iICMSTipo = mvariICMSTipo
End Property

Public Property Let iIPIValorManual(ByVal vData As Integer)
    mvariIPIValorManual = vData
End Property

Public Property Get iIPIValorManual() As Integer
    iIPIValorManual = mvariIPIValorManual
End Property

Public Property Let dIPIValor(ByVal vData As Double)
    mvardIPIValor = vData
End Property

Public Property Get dIPIValor() As Double
    dIPIValor = mvardIPIValor
End Property

Public Property Let dIPIBaseCalculo(ByVal vData As Double)
    mvardIPIBaseCalculo = vData
End Property

Public Property Get dIPIBaseCalculo() As Double
    dIPIBaseCalculo = mvardIPIBaseCalculo
End Property

Public Property Let iIPIAliquotaManual(ByVal vData As Integer)
    mvariIPIAliquotaManual = vData
End Property

Public Property Get iIPIAliquotaManual() As Integer
    iIPIAliquotaManual = mvariIPIAliquotaManual
End Property

Public Property Let dIPIAliquota(ByVal vData As Double)
    mvardIPIAliquota = vData
End Property

Public Property Get dIPIAliquota() As Double
    dIPIAliquota = mvardIPIAliquota
End Property

Public Property Let iIPITipoManual(ByVal vData As Integer)
    mvariIPITipoManual = vData
End Property

Public Property Get iIPITipoManual() As Integer
    iIPITipoManual = mvariIPITipoManual
End Property

Public Property Let iIPITipo(ByVal vData As Integer)
    mvariIPITipo = vData
End Property

Public Property Get iIPITipo() As Integer
    iIPITipo = mvariIPITipo
End Property

Public Property Let iTipoTributacaoManual(ByVal vData As Integer)
    mvariTipoTributacaoManual = vData
End Property

Public Property Get iTipoTributacaoManual() As Integer
    iTipoTributacaoManual = mvariTipoTributacaoManual
End Property

Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

Public Property Let dIPIPercRedBase(ByVal vData As Double)
    mvardIPIPercRedBase = vData
End Property

Public Property Get dIPIPercRedBase() As Double
    dIPIPercRedBase = mvardIPIPercRedBase
End Property

Public Property Let iIPIPercRedBaseManual(ByVal vData As Integer)
    mvariIPIPercRedBaseManual = vData
End Property

Public Property Get iIPIPercRedBaseManual() As Integer
    iIPIPercRedBaseManual = mvariIPIPercRedBaseManual
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let iNaturezaOpManual(ByVal vData As Integer)
    mvariNaturezaOpManual = vData
End Property

Public Property Get iNaturezaOpManual() As Integer
    iNaturezaOpManual = mvariNaturezaOpManual
End Property

Public Property Let iIPIBaseManual(ByVal vData As Integer)
    mvariIPIBaseManual = vData
End Property

Public Property Get iIPIBaseManual() As Integer
    iIPIBaseManual = mvariIPIBaseManual
End Property




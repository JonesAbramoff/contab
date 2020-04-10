VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCCMovDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodCaixa As Integer
Private mvariFilialEmpresa As Integer
Private mvardtData As Date
Private mvariTipoMeioPagto As Integer
Private mvariAdmMeioPagto As Integer
Private mvariParcelamento As Integer
Private mvardDeb As Double
Private mvardCred As Double

Public Property Let dCred(ByVal vData As Double)
    mvardCred = vData
End Property

Public Property Get dCred() As Double
    dCred = mvardCred
End Property

Public Property Let dDeb(ByVal vData As Double)
    mvardDeb = vData
End Property

Public Property Get dDeb() As Double
    dDeb = mvardDeb
End Property

Public Property Let iParcelamento(ByVal vData As Integer)
    mvariParcelamento = vData
End Property

Public Property Get iParcelamento() As Integer
    iParcelamento = mvariParcelamento
End Property

Public Property Let iAdmMeioPagto(ByVal vData As Integer)
    mvariAdmMeioPagto = vData
End Property

Public Property Get iAdmMeioPagto() As Integer
    iAdmMeioPagto = mvariAdmMeioPagto
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iCodCaixa(ByVal vData As Integer)
    mvariCodCaixa = vData
End Property

Public Property Get iCodCaixa() As Integer
    iCodCaixa = mvariCodCaixa
End Property




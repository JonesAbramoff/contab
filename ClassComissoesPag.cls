VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComissoesPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardtDataBaixa As Date
Private mvariCodVendedorIni As Integer
Private mvariCodVendedorFim As Integer
Private mvardtComisGeradasDe As Date
Private mvardtComisGeradasAte As Date
Private mvariTipo As Integer
Private mvarlNumIntComissao As Long
Private mvariStatusAnterior As Integer
Private mvariStatusNovo As Integer
Private mvariFilialEmpresa As Integer

Public Property Let iFilialEmpresa(ByVal vData As Long)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Long
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntComissao(ByVal vData As Long)
    mvarlNumIntComissao = vData
End Property

Public Property Get lNumIntComissao() As Long
    lNumIntComissao = mvarlNumIntComissao
End Property

Public Property Let iCodVendedorIni(ByVal vData As Integer)
    mvariCodVendedorIni = vData
End Property

Public Property Get iCodVendedorIni() As Integer
    iCodVendedorIni = mvariCodVendedorIni
End Property

Public Property Get iCodVendedorFim() As Integer
    iCodVendedorFim = mvariCodVendedorFim
End Property

Public Property Get iStatusNovo() As Integer
    iStatusNovo = mvariStatusNovo
End Property

Public Property Get iStatusAnterior() As Integer
    iStatusAnterior = mvariStatusAnterior
End Property

Public Property Let iCodVendedorFim(ByVal vData As Integer)
    mvariCodVendedorFim = vData
End Property

Public Property Let iStatusAnterior(ByVal vData As Integer)
    mvariStatusAnterior = vData
End Property

Public Property Let iStatusNovo(ByVal vData As Integer)
    mvariStatusNovo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Let dtDataBaixa(ByVal vData As Date)
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let dtComisGeradasDe(ByVal vData As Date)
    mvardtComisGeradasDe = vData
End Property

Public Property Get dtComisGeradasDe() As Date
    dtComisGeradasDe = mvardtComisGeradasDe
End Property

Public Property Let dtComisGeradasAte(ByVal vData As Date)
    mvardtComisGeradasAte = vData
End Property

Public Property Get dtComisGeradasAte() As Date
    dtComisGeradasAte = mvardtComisGeradasAte
End Property



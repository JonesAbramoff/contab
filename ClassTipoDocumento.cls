VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarsDescricao As String
Private mvariContabiliza As Integer
Private mvariAumentaValorPagto As Integer
Private mvariEmNFFatPag As Integer
Private mvariEmCreditoPagForn As Integer
Private mvariEmTituloRec As Integer
Private mvariEmDebitosRecCli As Integer
Private mvarsSigla As String
Private mvariClasseDocCPR As Integer
Private mvarsDescricaoReduzida As String

Public Property Let sDescricaoReduzida(ByVal vData As String)
    mvarsDescricaoReduzida = vData
End Property

Public Property Get sDescricaoReduzida() As String
    sDescricaoReduzida = mvarsDescricaoReduzida
End Property

Public Property Let iClasseDocCPR(ByVal vData As Integer)
    mvariClasseDocCPR = vData
End Property

Public Property Get iClasseDocCPR() As Integer
    iClasseDocCPR = mvariClasseDocCPR
End Property

Public Property Let sSigla(ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property

Public Property Let iEmDebitosRecCli(ByVal vData As Integer)
    mvariEmDebitosRecCli = vData
End Property

Public Property Get iEmDebitosRecCli() As Integer
    iEmDebitosRecCli = mvariEmDebitosRecCli
End Property

Public Property Let iEmTituloRec(ByVal vData As Integer)
    mvariEmTituloRec = vData
End Property

Public Property Get iEmTituloRec() As Integer
    iEmTituloRec = mvariEmTituloRec
End Property

Public Property Let iEmCreditoPagForn(ByVal vData As Integer)
    mvariEmCreditoPagForn = vData
End Property

Public Property Get iEmCreditoPagForn() As Integer
    iEmCreditoPagForn = mvariEmCreditoPagForn
End Property

Public Property Let iEmNFFatPag(ByVal vData As Integer)
    mvariEmNFFatPag = vData
End Property

Public Property Get iEmNFFatPag() As Integer
    iEmNFFatPag = mvariEmNFFatPag
End Property

Public Property Let iAumentaValorPagto(ByVal vData As Integer)
    mvariAumentaValorPagto = vData
End Property

Public Property Get iAumentaValorPagto() As Integer
    iAumentaValorPagto = mvariAumentaValorPagto
End Property

Public Property Let iContabiliza(ByVal vData As Integer)
    mvariContabiliza = vData
End Property

Public Property Get iContabiliza() As Integer
    iContabiliza = mvariContabiliza
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoMovtoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvarsNomeReduzido As String
Private mvarsSigla As String
Private mvarsSignificadoRefInterna As String
Private mvariCredito As Integer
Private mvariPodeCaixaECF As Integer
Private mvariPodeCaixaCentral As Integer
Private mvariPodeCaixaGeral As Integer
Private mvariTMPLoja As Integer
Private mvariRelMovCaixa As Integer

Public Property Let iRelMovCaixa(ByVal vData As Integer)
    mvariRelMovCaixa = vData
End Property

Public Property Get iRelMovCaixa() As Integer
    iRelMovCaixa = mvariRelMovCaixa
End Property

Public Property Let iTMPLoja(ByVal vData As Integer)
    mvariTMPLoja = vData
End Property

Public Property Get iTMPLoja() As Integer
    iTMPLoja = mvariTMPLoja
End Property

Public Property Let iPodeCaixaGeral(ByVal vData As Integer)
    mvariPodeCaixaGeral = vData
End Property

Public Property Get iPodeCaixaGeral() As Integer
    iPodeCaixaGeral = mvariPodeCaixaGeral
End Property

Public Property Let iPodeCaixaCentral(ByVal vData As Integer)
    mvariPodeCaixaCentral = vData
End Property

Public Property Get iPodeCaixaCentral() As Integer
    iPodeCaixaCentral = mvariPodeCaixaCentral
End Property

Public Property Let iPodeCaixaECF(ByVal vData As Integer)
    mvariPodeCaixaECF = vData
End Property

Public Property Get iPodeCaixaECF() As Integer
    iPodeCaixaECF = mvariPodeCaixaECF
End Property

Public Property Let iCredito(ByVal vData As Integer)
    mvariCredito = vData
End Property

Public Property Get iCredito() As Integer
    iCredito = mvariCredito
End Property

Public Property Let sSignificadoRefInterna(ByVal vData As String)
    mvarsSignificadoRefInterna = vData
End Property

Public Property Get sSignificadoRefInterna() As String
    sSignificadoRefInterna = mvarsSignificadoRefInterna
End Property

Public Property Let sSigla(ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Sub Copiar(objTipoMovtoCaixa As ClassTipoMovtoCaixa)

    With objTipoMovtoCaixa
    
        mvariCodigo = .iCodigo
        mvarsDescricao = .sDescricao
        mvarsNomeReduzido = .sNomeReduzido
        mvarsSigla = .sSigla
        mvarsSignificadoRefInterna = .sSignificadoRefInterna
        mvariCredito = .iCredito
        mvariPodeCaixaECF = .iPodeCaixaECF
        mvariPodeCaixaCentral = .iPodeCaixaCentral
        mvariPodeCaixaGeral = .iPodeCaixaGeral
        mvariTMPLoja = .iTMPLoja
    
    End With

End Sub
